"""
overview_reader.py — собирает плотный слайд-обзор (OverviewSlide) ИЗ xlsx
с листами «исх» (исходные сигналы) и «динамика» (проценты к прошлой неделе).

Особенность: точные имена колонок заранее неизвестны, поэтому колонки
определяются ПО КЛЮЧЕВЫМ СЛОВАМ в заголовках. Функция возвращает не только
данные, но и ОТЧЁТ по маппингу (какая колонка на какую роль легла), чтобы
можно было проверить разбор без отправки файла. Любую привязку можно
зафиксировать вручную через mapping=.

Агрегация детерминированная (без LLM): группируем по кластеру 1 уровня
(продукт) → кластеру 2 уровня (тема), считаем упоминания, источники,
тянем цитату и статус, джойним проценты из «динамика».
"""
from __future__ import annotations

import logging
import re
from difflib import SequenceMatcher
from typing import Dict, List, Optional, Tuple

import pandas as pd

logger = logging.getLogger(__name__)

# Ключевые слова заголовков для каждой роли (нижний регистр, подстрока).
# Порядок важен: первое совпадение выигрывает. Можно расширять.
HINTS = {
    "cluster_l1": ["кластер 1", "кластер1", "кластер_1", "1 уровн", "первого уровн",
                   "объект сигнал", "продукт", "группа", "родител", "parent", "l1", "объект"],
    "cluster_l2": ["кластер сигналов 2", "кластер 2", "кластер2", "кластер_2",
                   "2 уровн", "второго уровн", "тема", "подкластер", "child", "l2"],
    "source": ["источник сигнал", "канал", "чат", "source", "площадк", "источник"],
    "source_block": ["блок", "тип источ", "категория источ", "группа источ",
                     "вид источ", "сегмент", "раздел источ"],
    "text": ["текст из источника обратной связи", "текст", "сообщени", "цитат",
             "обращени", "feedback", "комментар", "отзыв", "реплик"],
    "status": ["статус", "в списке бол", "на анализе", "состояни", "метк болей"],
    "count": ["упомин", "кол-во сигнал", "количеств сигнал", "count", "частот"],
    "is_new": ["призн нов", "флаг нов", "is_new", "новая тема"],
}
DYN_HINTS = {
    "key": ["тема", "кластер 2", "кластер2", "2 уровн", "название", "l2", "name"],
    "product": ["продукт", "объект", "группа", "product"],
    "current": ["текущ", "current", "этой недел", "за неделю"],
    "previous": ["прошл", "previous", "предыдущ"],
    "pct": ["динамик", "%", "процент", "изменени", "к прошл", "delta", "дельта"],
}


def _norm(s) -> str:
    return str(s).strip().lower().replace("ё", "е")


def _parse_num(raw) -> Optional[float]:
    """Число из ячейки: поддержка '22', '0,22', '-45%', ' 12 '. Иначе None."""
    if raw is None:
        return None
    if isinstance(raw, bool):
        return None
    if isinstance(raw, (int, float)):
        return None if pd.isna(raw) else float(raw)
    s = str(raw).strip().replace("%", "").replace("\u00a0", "").replace(" ", "")
    if not s:
        return None
    s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def _match_key(s) -> str:
    """Жёсткая нормализация для джойна тем: регистр, ё/е, пунктуация, пробелы."""
    s = str(s).lower().replace("ё", "е")
    s = re.sub(r"[^\w\s]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def _dyn_lookup(table: dict, mk: str):
    """Ищем значение по ключу темы: точное совпадение, иначе по вхождению."""
    if mk in table:
        return table[mk]
    if len(mk) >= 10:
        for k, v in table.items():
            if k and len(k) >= 10 and (mk in k or k in mk):
                return v
    return None


_STOP = {"и", "в", "во", "на", "по", "с", "со", "для", "к", "из", "о", "об", "при",
         "не", "от", "до", "за", "the", "a", "of", "in", "to", "при", "что", "как"}


def _tokens(s) -> set:
    """Значимые слова темы (без пунктуации, стоп-слов и коротких токенов)."""
    return {w for w in _match_key(s).split() if len(w) > 2 and w not in _STOP}


def _similarity(a, b) -> float:
    """Схожесть двух названий тем: пересечение значимых слов + посимвольно."""
    ta, tb = _tokens(a), _tokens(b)
    inter = len(ta & tb)
    if ta and tb:
        jac = inter / len(ta | tb)
        overlap = inter / min(len(ta), len(tb))
    else:
        jac = overlap = 0.0
    seq = SequenceMatcher(None, _match_key(a), _match_key(b)).ratio()
    base = max(jac, 0.85 * overlap, 0.9 * seq)
    if inter < 2 and seq < 0.7:   # слабое пересечение слов — занижаем
        base *= 0.5
    return base


def _find_col(df: pd.DataFrame, keywords: List[str]) -> Optional[str]:
    cols = {c: _norm(c) for c in df.columns}
    for kw in keywords:
        for c, n in cols.items():
            if kw in n:
                return c
    return None


def _pick_sheet(xls: pd.ExcelFile, keywords: List[str]) -> Optional[str]:
    for name in xls.sheet_names:
        n = _norm(name)
        if any(kw in n for kw in keywords):
            return name
    return None


def read_overview(
    xlsx_path: str,
    title: str = "Голос IT: дайджест для программы AI PDLC",
    subtitle: Optional[str] = "Темы, волнующие сотрудников по продуктам программы PDLC",
    mapping: Optional[Dict[str, str]] = None,
    max_groups: int = 12,
    max_topics_per_group: int = 12,
):
    """
    Возвращает (OverviewSlide, report).

    mapping — необязательная ручная привязка ролей к именам колонок, напр.
    {"cluster_l1": "Кластер 1 ур.", "text": "Текст обращения", ...}. Переопределяет
    авто-детект. report — словарь {роль: колонка|None, "_sheets": (...)} для проверки.
    """
    from .schemas import (
        OverviewSlide, OverviewKPI, OverviewTopic, ProductGroup, SourceBlock,
    )

    mapping = mapping or {}
    xls = pd.ExcelFile(xlsx_path)

    src_sheet = mapping.get("_src_sheet") or _pick_sheet(xls, ["исх", "source", "сигнал", "данны"])
    dyn_sheet = mapping.get("_dyn_sheet") or _pick_sheet(xls, ["динамик", "dynamic", "процент", "%"])
    if src_sheet is None:
        src_sheet = xls.sheet_names[0]

    df = xls.parse(src_sheet)
    df.columns = [str(c) for c in df.columns]

    # привязка колонок
    cols: Dict[str, Optional[str]] = {}
    for role, kws in HINTS.items():
        cols[role] = mapping.get(role) or _find_col(df, kws)

    report = {**cols, "_src_sheet": src_sheet, "_dyn_sheet": dyn_sheet,
              "_src_columns": list(df.columns)}

    # обязательные роли
    c_l1, c_l2 = cols["cluster_l1"], cols["cluster_l2"]
    if c_l2 is None:
        raise ValueError(
            "Не нашёл колонку темы (кластер 2 уровня). Колонки листа: "
            f"{list(df.columns)}. Задай вручную mapping={{'cluster_l2': '<имя>'}}."
        )
    if c_l1 is None:
        c_l1 = c_l2  # нет группировки — каждая тема сама себе группа

    c_src, c_block = cols["source"], cols["source_block"]
    c_text, c_status = cols["text"], cols["status"]
    c_count, c_new = cols["count"], cols["is_new"]

    df = df.dropna(subset=[c_l2])
    l1_fallback = c_l1 if c_l1 != c_l2 else None  # напр. «Объект сигнала»

    # ---- агрегаты по теме (кластер 2 уровня) ----
    def _mentions(sub) -> int:
        if c_count and c_count in sub:
            try:
                return int(pd.to_numeric(sub[c_count], errors="coerce").fillna(0).sum())
            except Exception:
                pass
        return len(sub)

    def _quote(sub) -> Optional[str]:
        if not c_text:
            return None
        # схлопываем переносы/пробелы (в данных — целые сообщения с авторами/датами)
        cands = []
        for v in sub[c_text].dropna():
            s = " ".join(str(v).split())
            if s:
                cands.append(s)
        if not cands:
            return None
        # берём самое содержательное сообщение и обрезаем до сниппета
        cands.sort(key=len, reverse=True)
        q = cands[0]
        return (q[:150].rstrip() + "…") if len(q) > 150 else q

    def _status(sub) -> Optional[str]:
        if not c_status:
            return None
        vals = [str(v).strip() for v in sub[c_status].dropna().tolist() if str(v).strip()]
        return vals[0] if vals else None

    def _is_new(sub) -> bool:
        if c_new and c_new in sub:
            return bool(pd.to_numeric(sub[c_new], errors="coerce").fillna(0).sum() > 0)
        st = (_status(sub) or "").lower()
        return "нов" in st or "new" in st

    topics_by_l2: Dict[str, dict] = {}
    l1_by_l2: Dict[str, Optional[str]] = {}
    for l2, sub in df.groupby(c_l2, sort=False):
        topics_by_l2[str(l2)] = dict(
            mentions=_mentions(sub), quote=_quote(sub),
            status=_status(sub), is_new=_is_new(sub),
        )
        if l1_fallback:
            v = sub[l1_fallback].dropna()
            l1_by_l2[str(l2)] = str(v.iloc[0]) if len(v) else None

    # ---- лист «динамика»: тема -> продукт(=группа) и % ----
    dyn_pct: Dict[str, float] = {}
    dyn_product: Dict[str, str] = {}
    dyn_new: Dict[str, bool] = {}
    dyn_entries: List[Tuple[str, str, Optional[str]]] = []  # (тема_ориг, ключ, продукт)
    if dyn_sheet:
        dd = xls.parse(dyn_sheet)
        dd.columns = [str(c) for c in dd.columns]
        dk = mapping.get("dyn_key") or _find_col(dd, DYN_HINTS["key"])
        dprod = mapping.get("dyn_product") or _find_col(dd, DYN_HINTS["product"])
        dcur = mapping.get("dyn_current") or _find_col(dd, DYN_HINTS["current"])
        dprev = mapping.get("dyn_previous") or _find_col(dd, DYN_HINTS["previous"])
        dp = mapping.get("dyn_pct") or _find_col(dd, DYN_HINTS["pct"])
        report.update(dyn_key=dk, dyn_product=dprod, dyn_current=dcur,
                      dyn_previous=dprev, dyn_pct=dp, _dyn_columns=list(dd.columns))
        if dk:
            raw_pcts: Dict[str, float] = {}   # колонка «динамика» (в этих данных — модуль)
            counts: Dict[str, Tuple] = {}     # ключ -> (текущая, прошлая)
            for _, row in dd.iterrows():
                key = _match_key(row[dk])
                if not key:
                    continue
                prod_val = str(row[dprod]).strip() if (dprod and pd.notna(row[dprod])) else None
                if prod_val:
                    dyn_product[key] = prod_val
                dyn_entries.append((str(row[dk]).strip(), key, prod_val))
                if dp is not None:
                    v = _parse_num(row[dp])
                    if v is not None:
                        raw_pcts[key] = v
                cur = pd.to_numeric(row[dcur], errors="coerce") if dcur else None
                prev = pd.to_numeric(row[dprev], errors="coerce") if dprev else None
                counts[key] = (cur, prev)
            # доли (0.22) или проценты (22)? — только для запасной колонки
            vals = list(raw_pcts.values())
            as_fraction = bool(vals) and max(abs(v) for v in vals) <= 1.5
            for key in set(list(raw_pcts) + list(counts)):
                cur, prev = counts.get(key, (None, None))
                has_cur = cur is not None and pd.notna(cur)
                has_prev = prev is not None and pd.notna(prev)
                pct = None
                if has_cur and has_prev and prev != 0:
                    # ЗНАКОВЫЙ % из счётчиков — даёт направление (рост/падение)
                    pct = (float(cur) - float(prev)) / float(prev) * 100.0
                elif key in raw_pcts:
                    # запас: готовая колонка (в этих данных — модуль, без знака)
                    pct = raw_pcts[key] * 100.0 if as_fraction else raw_pcts[key]
                if pct is not None:
                    dyn_pct[key] = float(pct)
                elif has_cur and not has_prev:
                    # есть сигналы на этой неделе, но нет прошлой -> новая тема
                    dyn_new[key] = True

    # ---- источники по блокам ----
    source_blocks: List = []
    if c_src:
        if c_block:
            for block, sub in df.groupby(c_block, sort=False):
                tags = sorted({str(s).strip() for s in sub[c_src].dropna() if str(s).strip()})
                if tags:
                    source_blocks.append(SourceBlock(title=str(block)[:40], tags=tags[:40]))
            source_blocks.sort(key=lambda b: -len(b.tags))
            source_blocks = source_blocks[:2]
        else:
            tags = sorted({str(s).strip() for s in df[c_src].dropna() if str(s).strip()})
            source_blocks = [SourceBlock(title="Источники", tags=tags[:40])]

    # ---- сопоставление темы из «исх» с темой из «динамика» ----
    dyn_keys = set(dyn_pct) | set(dyn_new) | set(dyn_product)

    def _lookup_key(l2: str, l1: Optional[str]):
        """Возвращает (ключ_в_динамике | None, был_ли_нечёткий_матч)."""
        mk = _match_key(l2)
        if mk in dyn_keys:
            return mk, False
        if len(mk) >= 10:                                  # по вхождению
            for k in dyn_keys:
                if len(k) >= 10 and (mk in k or k in mk):
                    return k, False
        best_key, best_sc = None, 0.0                      # по схожести
        for theme, ekey, prod in dyn_entries:
            sc = _similarity(l2, theme)
            if l1 and prod:                                # бонус за совпадение продукта
                pl, pp = _match_key(l1), _match_key(prod)
                if pl and pp and (pl in pp or pp in pl):
                    sc += 0.12
            if sc > best_sc:
                best_key, best_sc = ekey, sc
        if best_key is not None and best_sc >= 0.5:
            return best_key, True
        return None, False

    resolved: Dict[str, Tuple[Optional[str], bool]] = {
        l2: _lookup_key(l2, l1_by_l2.get(l2)) for l2 in topics_by_l2
    }

    # ---- группировка тем по продукту (из «динамика», иначе L1, иначе «Прочее») ----
    group_topics: Dict[str, List[Tuple[str, dict]]] = {}
    for l2, facts in topics_by_l2.items():
        mkey = resolved[l2][0]
        prod = (dyn_product.get(mkey) if mkey else None) or l1_by_l2.get(l2) or "Прочее"
        group_topics.setdefault(prod, []).append((l2, facts))

    groups: List = []
    new_count = 0
    unmatched_dyn: List[str] = []
    fuzzy_count = 0
    for prod, items in group_topics.items():
        items.sort(key=lambda x: -x[1]["mentions"])
        ov_topics = []
        for l2, facts in items[:max_topics_per_group]:
            mkey, fuzzy = resolved[l2]
            pct = dyn_pct.get(mkey) if mkey else None
            is_new = facts["is_new"] or (bool(dyn_new.get(mkey)) if mkey else False)
            if fuzzy and pct is not None:
                fuzzy_count += 1
            if is_new:
                new_count += 1
            if pct is None and not is_new:
                unmatched_dyn.append(str(l2))
            ov_topics.append(OverviewTopic(
                title=str(l2)[:160], quote=facts["quote"], mentions=facts["mentions"],
                dynamics_pct=pct,
                status=("new" if is_new and not facts["status"] else facts["status"]),
            ))
        gtotal = sum(t.mentions for t in ov_topics)
        groups.append((gtotal, ProductGroup(name=str(prod)[:80], topics=ov_topics)))
    groups.sort(key=lambda x: -x[0])
    groups = [g for _, g in groups[:max_groups]]

    # ---- KPI ----
    total_signals = sum(f["mentions"] for f in topics_by_l2.values())
    n_sources = df[c_src].nunique() if c_src else 0
    n_topics = len(topics_by_l2)
    kpis = [
        OverviewKPI(value=str(int(total_signals)), label="Сигналов проанализировано", icon_hint="signal"),
        OverviewKPI(value=str(int(n_sources)), label="Источников обратной связи", icon_hint="search"),
        OverviewKPI(value=str(int(n_topics)), label="Активных тем", icon_hint="growth"),
        OverviewKPI(value=str(int(new_count)), label="Новых тем", icon_hint="info"),
    ]

    overview = OverviewSlide(
        title=title, subtitle=subtitle, kpis=kpis,
        source_blocks=source_blocks, groups=groups,
    )
    # Прозрачность: что РЕАЛЬНО прочитано из данных (для проверки без файла)
    report["_rows_read"] = int(len(df))
    report["_sources_extracted"] = sorted(
        {str(s).strip() for s in df[c_src].dropna()} if c_src else set()
    )
    report["_groups_extracted"] = [
        (g.name, [(t.title, t.mentions, t.dynamics_pct) for t in g.topics])
        for g in groups
    ]
    report["_dyn_matched"] = sum(
        1 for g in groups for t in g.topics if t.dynamics_pct is not None
    )
    report["_dyn_unmatched"] = unmatched_dyn
    report["_dyn_fuzzy"] = fuzzy_count
    logger.info("Overview-маппинг: %s", {k: v for k, v in report.items()
                                         if not k.startswith("_")})
    return overview, report


def format_report(report: Dict) -> str:
    """Человекочитаемый отчёт по маппингу — вывести пользователю для проверки."""
    lines = ["Разбор xlsx для слайда-обзора:",
             f"  лист исходных: {report.get('_src_sheet')}",
             f"  лист динамики: {report.get('_dyn_sheet')}",
             "  колонки -> роли:"]
    roles = ["cluster_l1", "cluster_l2", "source", "source_block",
             "text", "status", "count", "is_new",
             "dyn_key", "dyn_product", "dyn_current", "dyn_previous", "dyn_pct"]
    human = {
        "cluster_l1": "группа/продукт (кластер 1)",
        "cluster_l2": "тема (кластер 2)",
        "source": "источник", "source_block": "блок источника",
        "text": "текст/цитата", "status": "статус",
        "count": "кол-во упоминаний", "is_new": "признак новизны",
        "dyn_key": "'динамика': ключ-тема", "dyn_product": "'динамика': продукт=группа",
        "dyn_current": "'динамика': кол-во тек. неделя", "dyn_previous": "'динамика': кол-во прош. неделя",
        "dyn_pct": "'динамика': готовый %",
    }
    for r in roles:
        col = report.get(r)
        mark = "✓" if col else "— НЕ НАЙДЕНО"
        lines.append(f"    {human[r]:32s}: {col or ''} {mark}")

    # что РЕАЛЬНО прочитано из данных — для проверки, что ничего не выдумано
    if "_rows_read" in report:
        lines.append(f"  прочитано строк (лист исх): {report['_rows_read']}")
    srcs = report.get("_sources_extracted")
    if srcs is not None:
        lines.append(f"  источники из данных ({len(srcs)}):")
        for s in srcs:
            lines.append(f"    • {s}")
    groups = report.get("_groups_extracted")
    if groups:
        matched = report.get("_dyn_matched")
        fuzzy = report.get("_dyn_fuzzy") or 0
        suffix = ""
        if matched is not None:
            suffix = f" (динамика подтянулась к {matched} темам"
            suffix += f", из них по схожести: {fuzzy})" if fuzzy else ")"
        lines.append("  группы и темы из данных" + suffix + ":")
        for gname, topics in groups:
            lines.append(f"    ▸ {gname}")
            for item in topics:
                if isinstance(item, tuple):
                    title, ment, pct = item
                    p = "—" if pct is None else f"{pct:+.0f}%"
                    lines.append(f"        – {title}  [{ment} упом., динамика {p}]")
                else:
                    lines.append(f"        – {item}")
    unmatched = report.get("_dyn_unmatched")
    if unmatched:
        lines.append(f"  БЕЗ динамики (не нашёл % по названию темы) — {len(unmatched)}:")
        for t in unmatched:
            lines.append(f"    ✗ {t}")
        lines.append("    (проверь, совпадает ли 'тема' в листе 'динамика' с "
                     "'Кластер сигналов 2 уровня' в 'исх')")
    return "\n".join(lines)
