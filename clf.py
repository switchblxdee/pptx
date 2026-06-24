import numpy as np
import re
import pickle
import os
from typing import Dict, List, Optional, Tuple
from tqdm.auto import tqdm
from rank_bm25 import BM25Okapi

# Fuzzy: rapidfuzz если доступен, иначе мягкий фолбэк на difflib (без typo-robustness)
try:
    from rapidfuzz import fuzz as _rf_fuzz
    _HAS_RAPIDFUZZ = True
except ImportError:
    import difflib
    _HAS_RAPIDFUZZ = False


# Веса фьюжн-сигналов для матчинга родителя по имени. Тюнятся под таксономию.
DEFAULT_PARENT_MATCH_WEIGHTS = {
    "semantic":  0.25,   # косинус по эмбеддингам имён
    "bm25":      0.20,   # лексический, IDF-взвешенный
    "fuzzy":     0.25,   # token_set / partial ratio
    "coverage":  0.15,   # доля токенов кандидата, присутствующих в запросе
    "contained": 0.10,   # токены кандидата ⊆ токены запроса
    "prefix":    0.05,   # запрос начинается с имени кандидата
    # length_pref: >0 — при равенстве предпочесть более длинное (специфичное) имя,
    #              <0 — более короткое (базовый код). 0 — нейтрально (по умолчанию).
    "length_pref": 0.0,
}


def _length_pref_active(weights: Dict[str, float]) -> bool:
    return abs(weights.get("length_pref", 0.0)) > 1e-9


class HierarchicalEmbeddingClassifier:
    def __init__(self, embeddings_model, batch_size: int = 100, rrf_k: int = 60,
                 use_new: bool = False,
                 parent_match_weights: Optional[Dict[str, float]] = None,
                 cache_path: str = "/home/work/23101393@sigma.sbrf.ru/vs_code/instra_file/hierarchy_cache.pkl"):
        self.model = embeddings_model
        self.batch_size = batch_size
        self.rrf_k = rrf_k
        self.use_new = use_new
        self.cache_path = cache_path
        self.parent_match_weights = {**DEFAULT_PARENT_MATCH_WEIGHTS,
                                     **(parent_match_weights or {})}

        self.dim: Optional[int] = None

        self.hierarchy: Dict[str, Dict] = {}
        self.parent_names: List[str] = []
        self.parent_vectors: Optional[np.ndarray] = None
        self.parent_name_vectors: Optional[np.ndarray] = None

        self.all_subclass_texts: List[str] = []
        self.all_subclass_parents: List[str] = []
        self.all_subclass_vectors: Optional[np.ndarray] = None

        self.bm25_parent_index = None
        self.bm25_subclass_indexes: Dict[str, BM25Okapi] = {}
        self.tokenized_subclasses: Dict[str, List[List[str]]] = {}

        # Ленивый индекс для лексического/fuzzy матчинга имён родителей
        self.parent_names_normalized: Optional[List[str]] = None
        self.parent_name_tokens: Optional[List[set]] = None

    # ------------------------------------------------------------------ #
    # Утилиты
    # ------------------------------------------------------------------ #
    def _tokenize(self, text: str) -> List[str]:
        text = text.lower() if text else "пусто"
        return re.findall(r'\w+', text)

    def _normalize_name(self, text: str) -> str:
        return " ".join(self._tokenize(text))

    @staticmethod
    def _normalize(vec: np.ndarray) -> np.ndarray:
        n = np.linalg.norm(vec)
        return vec / n if n > 0 else vec

    @staticmethod
    def _normalize_matrix(mat: np.ndarray) -> np.ndarray:
        norms = np.linalg.norm(mat, axis=1, keepdims=True)
        norms = np.where(norms == 0, 1.0, norms)
        return mat / norms

    def _cosine_topk(self, query_vector: np.ndarray, matrix: Optional[np.ndarray],
                     top_k: int) -> List[Tuple[int, float]]:
        if matrix is None or len(matrix) == 0:
            return []
        q = self._normalize(np.asarray(query_vector, dtype=float))
        m = self._normalize_matrix(np.asarray(matrix, dtype=float))
        sims = m @ q
        k = min(top_k, len(sims))
        part = np.argpartition(sims, -k)[-k:]
        order = part[np.argsort(sims[part])[::-1]]
        return [(int(i), float(sims[i])) for i in order]

    def _cosine_all(self, query_vector: np.ndarray, matrix: Optional[np.ndarray]) -> np.ndarray:
        if matrix is None or len(matrix) == 0:
            return np.array([])
        q = self._normalize(np.asarray(query_vector, dtype=float))
        m = self._normalize_matrix(np.asarray(matrix, dtype=float))
        return m @ q

    def _fuzzy(self, q_norm: str, cand_norm: str) -> float:
        if not q_norm or not cand_norm:
            return 0.0
        if _HAS_RAPIDFUZZ:
            return max(_rf_fuzz.token_set_ratio(q_norm, cand_norm),
                       _rf_fuzz.partial_ratio(q_norm, cand_norm)) / 100.0
        return difflib.SequenceMatcher(None, q_norm, cand_norm).ratio()

    def _get_query_vector(self, query: str) -> np.ndarray:
        query_clean = query[:4096] if query else "пусто"
        return np.asarray(self.model.embed_query(query_clean), dtype=float)

    # ------------------------------------------------------------------ #
    # Построение
    # ------------------------------------------------------------------ #
    def _embed_subclasses(self, parent_name: str, subclasses: List[str]) -> np.ndarray:
        subclass_vectors: List[np.ndarray] = []
        for i in range(0, len(subclasses), self.batch_size):
            batch = subclasses[i:i + self.batch_size]
            batch_clean = [text[:4096] if text else "пусто" for text in batch]
            try:
                batch_embeddings = self.model.embed_documents(batch_clean)
                if len(batch_embeddings) != len(batch):
                    print(f"Warning: для '{parent_name}' получено "
                          f"{len(batch_embeddings)} вместо {len(batch)}")
                subclass_vectors.extend(np.asarray(v, dtype=float) for v in batch_embeddings)
            except Exception as e:
                print(f"Ошибка для родителя '{parent_name}': {e}")
                for sub in batch:
                    try:
                        vec = self.model.embed_query((sub or "пусто")[:4096])
                        subclass_vectors.append(np.asarray(vec, dtype=float))
                    except Exception:
                        subclass_vectors.append(np.zeros(self.dim or 768))

        if self.dim is None:
            for v in subclass_vectors:
                if v.size > 0 and np.linalg.norm(v) > 0:
                    self.dim = int(v.size)
                    break

        dim = self.dim or 768
        fixed = [v if v.size == dim else np.zeros(dim) for v in subclass_vectors]
        arr = np.array(fixed, dtype=float) if fixed else np.zeros((0, dim))

        if len(arr) < len(subclasses):
            pad = np.zeros((len(subclasses) - len(arr), dim))
            arr = np.vstack([arr, pad]) if len(arr) else pad
            print(f"Warning: для '{parent_name}' добавлен padding до {len(subclasses)}")
        elif len(arr) > len(subclasses):
            arr = arr[:len(subclasses)]
        return arr

    def _centroid(self, subclass_vectors: np.ndarray) -> np.ndarray:
        dim = self.dim or (subclass_vectors.shape[1] if subclass_vectors.ndim == 2 else 768)
        if subclass_vectors.size == 0:
            return np.zeros(dim)
        row_norms = np.linalg.norm(subclass_vectors, axis=1)
        mask = row_norms > 0
        if not mask.any():
            return np.zeros(dim)
        normed = self._normalize_matrix(subclass_vectors[mask])
        return self._normalize(normed.mean(axis=0))

    def _build_parent_match_index(self):
        self.parent_names_normalized = [self._normalize_name(n) for n in self.parent_names]
        self.parent_name_tokens = [set(self._tokenize(n)) for n in self.parent_names]

    def _ensure_parent_match_index(self):
        if self.parent_names_normalized is None or self.parent_name_tokens is None:
            self._build_parent_match_index()

    def build_hierarchy(self, data: Dict[str, List[str]]):
        if not self.use_new and os.path.exists(self.cache_path):
            print(f"Загружаем иерархию из {self.cache_path}...")
            self._load_cache()
            self._ensure_parent_match_index()
            return

        print("Строим иерархию заново...")
        self.hierarchy = {}
        self.parent_names = []
        self.all_subclass_texts = []
        self.all_subclass_parents = []
        parent_vectors_list: List[np.ndarray] = []
        all_subclass_vectors_list: List[np.ndarray] = []

        non_empty_data = {k: v for k, v in data.items() if v and len(v) > 0}
        print(f"Всего родительских классов: {len(data)}; с подклассами: {len(non_empty_data)}")

        print("Получаем эмбеддинги для подклассов...")
        for parent_name, subclasses in tqdm(non_empty_data.items(), desc="Обработка подклассов"):
            subclass_vectors = self._embed_subclasses(parent_name, subclasses)
            parent_vector = self._centroid(subclass_vectors)

            self.hierarchy[parent_name] = {
                'subclasses': subclasses,
                'vectors': subclass_vectors,
                'parent_vector': parent_vector,
            }
            self.parent_names.append(parent_name)
            parent_vectors_list.append(parent_vector)
            for sub, vec in zip(subclasses, subclass_vectors):
                self.all_subclass_texts.append(sub)
                self.all_subclass_parents.append(parent_name)
                all_subclass_vectors_list.append(vec)

        if parent_vectors_list:
            self.parent_vectors = np.array(parent_vectors_list)
            self.all_subclass_vectors = np.array(all_subclass_vectors_list)
            print(f"Готово! Обработано {len(self.parent_names)} родительских классов")
        else:
            self.parent_vectors = None
            self.all_subclass_vectors = None
            print("Не создано ни одного родительского класса")

        if self.parent_names:
            try:
                name_vecs = self.model.embed_documents(self.parent_names)
                self.parent_name_vectors = np.array([np.asarray(v, dtype=float) for v in name_vecs])
            except Exception as e:
                print(f"Не удалось построить parent_name_vectors: {e}")
                self.parent_name_vectors = None

        print("Строим BM25 индексы...")
        for parent_name, p_data in self.hierarchy.items():
            tokenized = [self._tokenize(sub) for sub in p_data['subclasses']]
            self.tokenized_subclasses[parent_name] = tokenized
            self.bm25_subclass_indexes[parent_name] = BM25Okapi(tokenized)
        if self.parent_names:
            self.bm25_parent_index = BM25Okapi([self._tokenize(n) for n in self.parent_names])
        print("BM25 индексы построены.")

        self._build_parent_match_index()
        self._save_cache()

    # ------------------------------------------------------------------ #
    # Кэш
    # ------------------------------------------------------------------ #
    def _save_cache(self):
        state = {
            'dim': self.dim,
            'hierarchy': self.hierarchy,
            'parent_names': self.parent_names,
            'parent_vectors': self.parent_vectors,
            'parent_name_vectors': self.parent_name_vectors,
            'all_subclass_texts': self.all_subclass_texts,
            'all_subclass_parents': self.all_subclass_parents,
            'all_subclass_vectors': self.all_subclass_vectors,
            'bm25_parent_index': self.bm25_parent_index,
            'bm25_subclass_indexes': self.bm25_subclass_indexes,
            'tokenized_subclasses': self.tokenized_subclasses,
            'parent_names_normalized': self.parent_names_normalized,
            'parent_name_tokens': self.parent_name_tokens,
        }
        with open(self.cache_path, 'wb') as f:
            pickle.dump(state, f)
        print(f"Иерархия сохранена в {self.cache_path}")

    def _load_cache(self):
        with open(self.cache_path, 'rb') as f:
            state = pickle.load(f)
        self.dim = state.get('dim')
        self.hierarchy = state['hierarchy']
        self.parent_names = state['parent_names']
        self.parent_vectors = state['parent_vectors']
        self.parent_name_vectors = state.get('parent_name_vectors')
        self.all_subclass_texts = state['all_subclass_texts']
        self.all_subclass_parents = state['all_subclass_parents']
        self.all_subclass_vectors = state.get('all_subclass_vectors')
        self.bm25_parent_index = state['bm25_parent_index']
        self.bm25_subclass_indexes = state['bm25_subclass_indexes']
        self.tokenized_subclasses = state['tokenized_subclasses']
        self.parent_names_normalized = state.get('parent_names_normalized')
        self.parent_name_tokens = state.get('parent_name_tokens')
        print("Иерархия загружена из кэша.")

    # ------------------------------------------------------------------ #
    # УНИВЕРСАЛЬНЫЙ матчер родителя по имени
    # ------------------------------------------------------------------ #
    def find_parent_universal(self, name: str, top_parents: int = 20,
                              weights: Optional[Dict[str, float]] = None,
                              return_signals: bool = False,
                              query_vector: Optional[np.ndarray] = None) -> List[Dict]:
        """
        Робастный матчинг по ИМЕНИ родителя из шумной строки
        ("СУППД Новости и вопросы" -> "СУППД"). Фьюжн semantic + bm25 + fuzzy + lexical-приоры.
        Веса см. DEFAULT_PARENT_MATCH_WEIGHTS; можно переопределить per-call.
        query_vector — предпосчитанный эмбеддинг `name` (чтобы не эмбеддить повторно).
        """
        if not name:
            return [{"error": "No parent!"}]
        if not self.parent_names:
            return []
        self._ensure_parent_match_index()
        w = {**self.parent_match_weights, **(weights or {})}
        use_len = _length_pref_active(w)

        q_norm = self._normalize_name(name)
        q_tokens = set(self._tokenize(name))
        q_padded = f" {q_norm} "
        n = len(self.parent_names)

        # semantic -> [0,1]
        sem = np.zeros(n)
        if self.parent_name_vectors is not None and len(self.parent_name_vectors) == n:
            qv = query_vector if query_vector is not None else self._get_query_vector(name)
            cos = self._cosine_all(qv, self.parent_name_vectors)
            if cos.size == n:
                sem = (cos + 1.0) / 2.0

        # bm25 -> нормировка по максимуму
        bm25 = np.zeros(n)
        if self.bm25_parent_index is not None:
            scores = np.asarray(self.bm25_parent_index.get_scores(self._tokenize(name)), dtype=float)
            if scores.size == n:
                mx = scores.max()
                bm25 = scores / mx if mx > 0 else scores

        results = []
        for i, p_name in enumerate(self.parent_names):
            p_norm = self.parent_names_normalized[i]
            p_tokens = self.parent_name_tokens[i]

            coverage = (len(p_tokens & q_tokens) / len(p_tokens)) if p_tokens else 0.0
            contained = 1.0 if p_tokens and p_tokens <= q_tokens else 0.0
            prefix = 1.0 if (q_norm == p_norm or q_norm.startswith(p_norm + " ")) else 0.0
            substr = 1.0 if (p_norm and f" {p_norm} " in q_padded) else 0.0
            fuzzy = self._fuzzy(q_norm, p_norm)
            length_pref = (len(p_tokens) / max(1, len(q_tokens))) if use_len else 0.0

            score = (w["semantic"] * sem[i]
                     + w["bm25"] * bm25[i]
                     + w["fuzzy"] * fuzzy
                     + w["coverage"] * coverage
                     + w["contained"] * contained
                     + w["prefix"] * max(prefix, substr)
                     + w["length_pref"] * length_pref)

            item = {"key": p_name, "similarity": float(score)}
            if return_signals:
                item["signals"] = {
                    "semantic": float(sem[i]), "bm25": float(bm25[i]), "fuzzy": float(fuzzy),
                    "coverage": float(coverage), "contained": float(contained),
                    "prefix": float(prefix), "substr": float(substr),
                }
            results.append(item)

        results.sort(key=lambda x: x["similarity"], reverse=True)
        return results[:top_parents]

    def find_sim_parent_by_parent(self, parent: str, top_parents: int = 20) -> List[Dict]:
        """Бэк-совместимый алиас на универсальный матчер по имени."""
        return self.find_parent_universal(parent, top_parents=top_parents)

    # ------------------------------------------------------------------ #
    # РОУТЕР: классификация жалобы -> родитель (имя + подклассы)
    # ------------------------------------------------------------------ #
    @staticmethod
    def _minmax(d: Dict[str, float]) -> Dict[str, float]:
        """Min-max нормировка значений словаря в [0,1] (константа -> нули)."""
        if not d:
            return {}
        vals = list(d.values())
        lo, hi = min(vals), max(vals)
        if hi - lo < 1e-12:
            return {k: 0.0 for k in d}
        return {k: (v - lo) / (hi - lo) for k, v in d.items()}

    def find_parent_for_query(self, query: str, top_k: int = 5,
                              query_vector: Optional[np.ndarray] = None,
                              w_name: float = 0.45, w_subclass: float = 0.55,
                              subclass_topm: int = 3,
                              return_signals: bool = False) -> List[Dict]:
        """
        Роутинг свободной жалобы к родителю. Сливает два сигнала:
          name     — лексика+семантика по ИМЕНИ родителя (ловит "сломался API GigaChat");
          subclass — среднее top-m косинусов запроса с подклассами родителя
                     (ловит чистый симптом: "401" ~ "Проблема с подключением по API").
        Оба сигнала min-max нормируются по кандидатам и взвешенно складываются.
        """
        if not self.parent_names:
            return []
        if query_vector is None:
            query_vector = self._get_query_vector(query)

        # --- name-сигнал (переиспользуем универсальный матчер, без повторного эмбеддинга) ---
        name_list = self.find_parent_universal(
            query, top_parents=len(self.parent_names), query_vector=query_vector)
        name_raw = {r["key"]: r["similarity"] for r in name_list if "key" in r}

        # --- subclass-сигнал: top-m косинусов по подклассам каждого родителя (max-pool) ---
        sub_raw = {p: 0.0 for p in self.parent_names}
        if self.all_subclass_vectors is not None and len(self.all_subclass_vectors) > 0:
            sims = self._cosine_all(query_vector, self.all_subclass_vectors)
            buckets: Dict[str, List[float]] = {p: [] for p in self.parent_names}
            for sim, p in zip(sims, self.all_subclass_parents):
                buckets[p].append(float(sim))
            for p, vals in buckets.items():
                if vals:
                    vals.sort(reverse=True)
                    m = min(subclass_topm, len(vals))
                    sub_raw[p] = sum(vals[:m]) / m

        name_n = self._minmax(name_raw)
        sub_n = self._minmax(sub_raw)

        fused = []
        for p in self.parent_names:
            score = w_name * name_n.get(p, 0.0) + w_subclass * sub_n.get(p, 0.0)
            item = {
                'parent': p,
                'similarity': float(score),
                'num_subclasses': len(self.hierarchy[p]['subclasses']),
            }
            if return_signals:
                item['signals'] = {
                    'name': float(name_n.get(p, 0.0)),
                    'subclass': float(sub_n.get(p, 0.0)),
                    'name_raw': float(name_raw.get(p, 0.0)),
                    'subclass_raw': float(sub_raw.get(p, 0.0)),
                }
            fused.append(item)

        fused.sort(key=lambda x: x['similarity'], reverse=True)
        return fused[:top_k]

    # ------------------------------------------------------------------ #
    # Поиск родителя по контенту запроса (центроиды)
    # ------------------------------------------------------------------ #
    def find_parent(self, query: str, top_k: int = 3,
                    query_vector: Optional[np.ndarray] = None) -> List[Dict]:
        if self.parent_vectors is None or len(self.parent_vectors) == 0:
            return []
        if query_vector is None:
            query_vector = self._get_query_vector(query)
        ranked = self._cosine_topk(query_vector, self.parent_vectors, top_k)
        return [{
            'parent': self.parent_names[idx],
            'similarity': sim,
            'num_subclasses': len(self.hierarchy[self.parent_names[idx]]['subclasses']),
        } for idx, sim in ranked]

    def find_parent_hybrid(self, query: str, top_k: int = 5,
                           query_vector: Optional[np.ndarray] = None) -> List[Dict]:
        if not self.parent_names:
            return []
        if query_vector is None:
            query_vector = self._get_query_vector(query)
        cosine_ranked = [(self.parent_names[idx], sim)
                         for idx, sim in self._cosine_topk(query_vector, self.parent_vectors, 20)]
        bm25_ranked: List[Tuple[str, float]] = []
        if self.bm25_parent_index:
            scores = self.bm25_parent_index.get_scores(self._tokenize(query))
            top_bm25 = np.argsort(scores)[::-1][:20]
            bm25_ranked = [(self.parent_names[idx], float(scores[idx]))
                           for idx in top_bm25 if scores[idx] > 0]
        fused = self._rrf_fusion(cosine_ranked, bm25_ranked, top_k=top_k)
        return [{
            'parent': name, 'rrf_score': rrf_score,
            'num_subclasses': len(self.hierarchy[name]['subclasses']),
        } for name, rrf_score in fused]

    # ------------------------------------------------------------------ #
    # Поиск подкласса
    # ------------------------------------------------------------------ #
    def find_subclass(self, query: str, parent_name: Optional[str] = None, top_k: int = 5,
                      query_vector: Optional[np.ndarray] = None) -> List[Dict]:
        if query_vector is None:
            query_vector = self._get_query_vector(query)
        if parent_name:
            if parent_name not in self.hierarchy:
                raise ValueError(f"Parent class '{parent_name}' not found")
            names = self.hierarchy[parent_name]['subclasses']
            vectors = self.hierarchy[parent_name]['vectors']
            ranked = self._cosine_topk(query_vector, vectors, top_k)
            return [{'subclass': names[idx], 'similarity': sim, 'parent': parent_name}
                    for idx, sim in ranked]
        if self.all_subclass_vectors is None or len(self.all_subclass_vectors) == 0:
            return []
        ranked = self._cosine_topk(query_vector, self.all_subclass_vectors, top_k)
        return [{
            'subclass': self.all_subclass_texts[idx], 'similarity': sim,
            'parent': self.all_subclass_parents[idx],
        } for idx, sim in ranked]

    def _bm25_search_subclasses(self, query: str, parent_name: str,
                                top_k: int) -> List[Tuple[int, float]]:
        bm25 = self.bm25_subclass_indexes[parent_name]
        scores = bm25.get_scores(self._tokenize(query))
        top_indices = np.argsort(scores)[::-1][:top_k]
        return [(int(idx), float(scores[idx])) for idx in top_indices if scores[idx] > 0]

    def find_subclass_hybrid(self, query: str, parent_name: str, top_k: int = 5,
                             query_vector: Optional[np.ndarray] = None) -> List[Dict]:
        if parent_name not in self.hierarchy:
            raise ValueError(f"Parent class '{parent_name}' not found")
        if query_vector is None:
            query_vector = self._get_query_vector(query)
        cosine_ranked = self._cosine_topk(query_vector,
                                          self.hierarchy[parent_name]['vectors'], top_k=50)
        bm25_ranked = self._bm25_search_subclasses(query, parent_name, top_k=50)
        fused = self._rrf_fusion(cosine_ranked, bm25_ranked, top_k=top_k)
        subclasses = self.hierarchy[parent_name]['subclasses']
        return [{'subclass': subclasses[idx], 'rrf_score': rrf_score, 'parent': parent_name}
                for idx, rrf_score in fused]

    def _rrf_fusion(self, ranked_list1, ranked_list2, top_k: int):
        rrf_scores: Dict = {}
        k = self.rrf_k
        for rank, (item_id, _) in enumerate(ranked_list1, start=1):
            rrf_scores[item_id] = rrf_scores.get(item_id, 0) + 1.0 / (k + rank)
        for rank, (item_id, _) in enumerate(ranked_list2, start=1):
            rrf_scores[item_id] = rrf_scores.get(item_id, 0) + 1.0 / (k + rank)
        return sorted(rrf_scores.items(), key=lambda x: x[1], reverse=True)[:top_k]

    # ------------------------------------------------------------------ #
    # Главный вход
    # ------------------------------------------------------------------ #
    def classify(self, query: str, parent_name: Optional[str] = None, find_parent: bool = False,
                 top_parents: int = 20, top_subclasses: int = 40,
                 use_hybrid: bool = False) -> List[Dict]:
        query_vector = self._get_query_vector(query)

        if parent_name:
            if parent_name not in self.hierarchy:
                if not find_parent:
                    return [{'error': f"Parent class '{parent_name}' not found"}]
                top_sim_parents = self.find_parent_universal(parent_name, top_parents=top_parents)
                results = []
                for p in top_sim_parents:
                    p_name = p.get("key")
                    if p_name is None:
                        continue
                    if use_hybrid:
                        subclasses = self.find_subclass_hybrid(
                            query, p_name, top_k=top_subclasses, query_vector=query_vector)
                    else:
                        subclasses = self.find_subclass(
                            query, p_name, top_k=top_subclasses, query_vector=query_vector)
                    results.append({
                        'parent': {
                            'parent': p_name, 'similarity': p["similarity"],
                            'rrf_score': None,
                            'num_subclasses': len(self.hierarchy[p_name]['subclasses']),
                        },
                        'top_subclasses': subclasses,
                    })
                return results

            if use_hybrid:
                subclasses = self.find_subclass_hybrid(
                    query, parent_name, top_k=top_subclasses, query_vector=query_vector)
            else:
                subclasses = self.find_subclass(
                    query, parent_name, top_k=top_subclasses, query_vector=query_vector)
            return [{
                'parent': {
                    'parent': parent_name,
                    'similarity': 1.0 if not use_hybrid else None,
                    'rrf_score': None,
                    'num_subclasses': len(self.hierarchy[parent_name]['subclasses']),
                },
                'top_subclasses': subclasses,
            }]

        # parent_name не задан: роутим жалобу к родителю (имя + подклассы), затем подкласс
        parents = self.find_parent_for_query(query, top_k=top_parents, query_vector=query_vector)
        results = []
        for parent_info in parents:
            p_name = parent_info['parent']
            if use_hybrid:
                subclasses = self.find_subclass_hybrid(
                    query, p_name, top_k=top_subclasses, query_vector=query_vector)
            else:
                subclasses = self.find_subclass(
                    query, parent_name=p_name, top_k=top_subclasses, query_vector=query_vector)
            results.append({'parent': parent_info, 'top_subclasses': subclasses})
        return results

    # ------------------------------------------------------------------ #
    # Геттеры
    # ------------------------------------------------------------------ #
    def get_subclasses(self, parent_name: str) -> List[str]:
        return self.hierarchy.get(parent_name, {}).get('subclasses', [])

    def get_all_parents(self) -> List[str]:
        return self.parent_names

    def search_in_parent(self, query: str, parent_name: str, top_k: int = 5) -> List[Dict]:
        return self.find_subclass(query, parent_name=parent_name, top_k=top_k)