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
from typing import Callable, Dict, List, Optional, Tuple

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
    "current": ["текущ", "current", "этой недел", "за неделю", "тек недел", "тек."],
    "previous": ["прошл", "previous", "предыдущ", "пред недел", "пред."],
    "pct": ["динамик", "%", "процент", "изменени", "к прошл", "delta", "дельта"],
    "status": ["статус", "status", "состояни"],
}

# Фиксированные блоки источников — выводятся ВСЕГДА, независимо от данных.
FIXED_SOURCE_BLOCKS = [
    ("Чаты в Сберчате", [
        "RewAi: Code Review Agent", "делай вместе с Чемоданом", "AI Коктейль",
        "SberOS+P7+Почта", "SberWorks APIStudio", "GIGA IDE support",
        "Поддержка Sbermock", "Сбер Id. AI Workspace", "Atomic Code",
        "AI in Dev Community", "VibeCoding Community", "GigaChat API",
        "Первопроходцы Сбертрек",
    ]),
    ("Другие источники", [
        "Обращения в SberF1", "Виджет SW", "Открытые диалоги", "Help Desk SW",
    ]),
]


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


def _read_sheet(xls, sheet: str, header_keywords: List[str]) -> "pd.DataFrame":
    """Читает лист, САМ находя строку заголовков (над ней могут быть пустые
    строки), и нормализует имена колонок (схлопывает переносы/пробелы)."""
    raw = xls.parse(sheet, header=None)
    header_row = 0
    for i in range(min(10, len(raw))):
        joined = " ".join(_norm(v) for v in raw.iloc[i].tolist() if pd.notna(v))
        if sum(1 for kw in header_keywords if kw in joined) >= 2:
            header_row = i
            break
    df = xls.parse(sheet, header=header_row)
    df.columns = [" ".join(str(c).split()) for c in df.columns]
    # выкидываем полностью пустые колонки-«Unnamed»
    df = df.loc[:, [c for c in df.columns if not str(c).lower().startswith("unnamed")]]
    return df


_MONTHS = ("января|февраля|марта|апреля|мая|июня|июля|августа|сентября|"
           "октября|ноября|декабря|янв|фев|мар|апр|июн|июл|авг|сен|окт|ноя|дек")

# Частые русские имена — якорь для вырезания ФИО.
_FIRST_NAMES = {
    "Александр", "Александра", "Алексей", "Анастасия", "Анатолий", "Андрей",
    "Анна", "Антон", "Артём", "Артем", "Борис", "Вадим", "Валентина",
    "Валерий", "Василий", "Вера", "Виктор", "Виктория", "Виталий", "Владимир",
    "Владислав", "Вячеслав", "Галина", "Геннадий", "Георгий", "Григорий",
    "Дарья", "Денис", "Дмитрий", "Евгений", "Евгения", "Егор", "Екатерина",
    "Елена", "Елизавета", "Иван", "Игорь", "Илья", "Ирина", "Кирилл",
    "Константин", "Ксения", "Лариса", "Леонид", "Лидия", "Любовь", "Людмила",
    "Максим", "Марат", "Маргарита", "Марина", "Мария", "Михаил", "Надежда",
    "Наталья", "Наталия", "Никита", "Николай", "Олег", "Ольга", "Оксана",
    "Павел", "Пётр", "Петр", "Полина", "Роман", "Руслан", "Светлана", "Семён",
    "Семен", "Сергей", "Софья", "София", "Станислав", "Степан", "Тамара",
    "Татьяна", "Тимофей", "Тимур", "Фёдор", "Федор", "Эдуард", "Юлия", "Юрий",
    "Яков", "Ян", "Яна", "Арина", "Вероника", "Глеб", "Диана", "Жанна",
    "Зоя", "Инна", "Карина", "Лев", "Матвей", "Мирослава", "Олеся", "Раиса",
    "Савелий", "Таисия", "Услада", "Феликс", "Христина", "Эмма", "Ярослав",
}
_NAME_ALT = "|".join(sorted(_FIRST_NAMES, key=len, reverse=True))
_SURNAME_SUFFIX = (r"(?:ов|ова|ев|ева|ёв|ёва|ин|ина|ын|ына|ский|ская|"
                   r"цкий|цкая|их|ко|ук|юк|ян|ко|енко)")

_RE_DATE = [
    re.compile(r"\b\d{1,2}[.\-/]\d{1,2}[.\-/]\d{2,4}\b"),
    re.compile(r"\b\d{4}[.\-/]\d{1,2}[.\-/]\d{1,2}\b"),
    re.compile(r"\b\d{1,2}[.\-/]\d{1,2}\b(?!\d)"),
    re.compile(r"\b\d{1,2}\s+(?:%s)\b(?:\s+\d{2,4})?" % _MONTHS, re.IGNORECASE),
    re.compile(r"\b\d{1,2}:\d{2}(?::\d{2})?\b"),
]
_RE_NUM = [
    re.compile(r"\+?\d[\d\-\s()]{8,}\d"),   # телефоны с разделителями
    re.compile(r"\b\d{5,}\b"),               # длинные id/номера
]
_RE_PHONE = re.compile(
    r"(?:\+7|\+|\b8)[\s\-]?\(?\d{3}\)?[\s\-]?\d{3}[\s\-]?\d{2}[\s\-]?\d{2}\b"
)
_RE_IDLABEL = re.compile(r"\b(?:id|ид|№|тел\.?|телефон|номер|user|логин)\b[:#№\s]*", re.IGNORECASE)
_RE_NAME = [
    re.compile(r"\b[А-ЯЁ][а-яё]+%s\b\s+(?:%s)\b" % (_SURNAME_SUFFIX, _NAME_ALT)),  # Фамилия Имя
    re.compile(r"\b(?:%s)\b(?:\s+[А-ЯЁ][а-яё]+){0,2}" % _NAME_ALT),                # Имя [Отч] [Фам]
    re.compile(r"\b[А-ЯЁ][а-яё]+%s\b(?:\s*[А-ЯЁ]\.){1,2}" % _SURNAME_SUFFIX),       # Фамилия И.О.
]
# Общий шаблон ФИО: 2-3 заглавных кириллических слова подряд (любые имена,
# даже не из словаря). Бренды на кириллице исключаем стоп-листом.
_CYR_CAP = r"[А-ЯЁ][а-яё]{2,}"
_RE_NAME_GENERIC = re.compile(r"\b%s(?:\s+%s){1,2}\b" % (_CYR_CAP, _CYR_CAP))
# Отчества (в т.ч. осиротевшие после удаления имени/фамилии)
_RE_PATRONYMIC = re.compile(
    r"\b[А-ЯЁ][а-яё]+(?:евич|ович|евна|овна|ична|инична|ьевич|ьевна|ич|вна)\b"
)
# Метки автора в начале строки: «от: …», «автор: …», «пользователь — …»
_RE_AUTHOR_LABEL = re.compile(
    r"^\s*(?:от|автор|пользовател\w*|клиент|сотрудник|user|from)\s*[:>—–\-]\s*",
    re.IGNORECASE,
)
# Бренды/слова, которые НЕ являются ФИО, хотя выглядят как пара заглавных слов
_NAME_STOP = {"сбер", "сбербанк", "сбермаркет", "онлайн", "почта", "коктейль",
              "чемоданом", "сбертрек", "сберчате", "сберчат"}


def _strip_names_generic(s: str) -> str:
    def repl(m):
        words = m.group(0).split()
        if any(w.lower() in _NAME_STOP for w in words):
            return m.group(0)        # это бренд, не трогаем
        return " "
    return _RE_NAME_GENERIC.sub(repl, s)


def _anonymize(text: str) -> str:
    """Убирает из текста ФИО, телефоны/id/длинные номера, даты и метки автора.

    Цель — оставить только сам отзыв/запрос, без данных о пользователе.
    """
    s = text
    s = _RE_AUTHOR_LABEL.sub("", s)          # «от: …» в начале
    s = _RE_IDLABEL.sub(" ", s)              # метки id/тел/№
    s = _RE_PHONE.sub(" ", s)                # телефоны (якорные) — до дат
    for r in _RE_DATE:                        # даты/время
        s = r.sub(" ", s)
    for r in _RE_NUM:                        # длинные номера/id
        s = r.sub(" ", s)
    s = _strip_names_generic(s)              # ФИО: 2-3 заглавных слова (любые)
    for r in _RE_NAME:                       # ФИО из словаря (в т.ч. одиночное имя)
        s = r.sub(" ", s)
    s = _RE_PATRONYMIC.sub(" ", s)           # осиротевшие отчества
    s = re.sub(r"[А-ЯЁ]\.\s?[А-ЯЁ]\.", " ", s)  # инициалы И.О.
    # аккуратная чистка: НЕ трогаем дефисы внутри слов (mcp-sourcecontrol, где-то)
    s = re.sub(r"\(\s*\)|\[\s*\]", " ", s)              # пустые скобки
    s = re.sub(r"\s[—–\-]\s", " ", s)                   # осиротевшее тире в пробелах
    s = re.sub(r"(?<=\s)[—–:;,|]+(?=\s)", " ", s)       # осиротевшие разделители
    s = re.sub(r"\s+([,.:;!?])", r"\1", s)              # пробел перед пунктуацией
    s = " ".join(s.split())
    return s.strip(" .,:;—–-|«»\"'()")


def _pick_quote(series) -> Optional[str]:
    """Самый содержательный комментарий: обезличенный, схлопнутый, обрезанный."""
    cands = _quote_candidates(series)
    return cands[0] if cands else None


def _quote_candidates(series, limit: int = 6) -> List[str]:
    """Обезличенные кандидаты-комментарии темы, по убыванию длины, без дублей.

    Возвращает список (а не один), чтобы выбор мог делать LLM. Фолбэк —
    первый (самый длинный), что повторяет прежнее поведение _pick_quote.
    """
    seen = set()
    cands: List[str] = []
    for v in series.dropna():
        s = _anonymize(" ".join(str(v).split()))
        if not s or len(s) < 8:
            continue
        s = (s[:150].rstrip() + "…") if len(s) > 150 else s
        key = s.lower()
        if key in seen:
            continue
        seen.add(key)
        cands.append(s)
    cands.sort(key=len, reverse=True)
    return cands[:limit]


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


_LABEL_WORDS = ("период", "номер", "дайдж", "выпуск", "продукт", "тема",
                "статус", "динамик", "сигнал", "недел", "кол-во")


def _scan_labeled_value(xls, sheet: str, keywords: List[str]) -> Optional[str]:
    """Ищет в листе ячейку-метку (содержит одно из keywords) и возвращает
    значение рядом. Порядок: хвост после «:» в самой ячейке → СНИЗУ → справа.

    «Снизу» приоритетнее, потому что «Период»/«Номер дайджеста» оформлены как
    заголовки столбцов, а значение стоит в строке под ними. Кандидаты, которые
    сами являются метками (другой заголовок), отбрасываются — иначе «Период»
    цеплял соседний «Номер дайджеста».
    """
    def _is_label(v) -> bool:
        s = str(v).strip().lower()
        return any(w in s for w in _LABEL_WORDS)

    try:
        raw = pd.read_excel(xls, sheet_name=sheet, header=None, dtype=object)
    except Exception:
        return None
    nrows, ncols = raw.shape
    for i in range(nrows):
        for j in range(ncols):
            cell = raw.iat[i, j]
            if cell is None or (isinstance(cell, float) and pd.isna(cell)):
                continue
            s = str(cell).strip().lower()
            if not any(kw in s for kw in keywords):
                continue
            # «Период: 21–27 мая» в одной ячейке
            if ":" in str(cell):
                tail = str(cell).split(":", 1)[1].strip()
                if tail:
                    return tail
            # значение снизу, затем справа; пропускаем ячейки-метки
            for ri, rj in ((i + 1, j), (i, j + 1)):
                if 0 <= ri < nrows and 0 <= rj < ncols:
                    v = raw.iat[ri, rj]
                    if v is None or (isinstance(v, float) and pd.isna(v)):
                        continue
                    if not str(v).strip() or _is_label(v):
                        continue
                    return str(v).strip()
    return None


def read_overview(
    xlsx_path: str,
    title: str = "Голос IT: дайджест для программы AI PDLC",
    subtitle: Optional[str] = "Темы, волнующие сотрудников по продуктам программы PDLC",
    mapping: Optional[Dict[str, str]] = None,
    max_groups: int = 12,
    max_topics_per_group: int = 12,
    comment_picker: Optional[Callable[[List[dict]], Dict[str, str]]] = None,
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

    df = _read_sheet(xls, src_sheet,
                     ["источник", "кластер", "текст", "объект", "сигнал", "дата"])

    cols: Dict[str, Optional[str]] = {}
    for role, kws in HINTS.items():
        cols[role] = mapping.get(role) or _find_col(df, kws)
    report = {**cols, "_src_sheet": src_sheet, "_dyn_sheet": dyn_sheet,
              "_src_columns": list(df.columns)}

    c_l2 = cols["cluster_l2"]               # тема в «исх» — нужна ТОЛЬКО для комментариев
    c_src, c_block, c_text = cols["source"], cols["source_block"], cols["text"]

    # ---- индекс комментариев из «исх»: тема -> представительный комментарий ----
    comment_index: Dict[str, Tuple[str, str]] = {}   # ключ -> (тема_ориг, цитата)
    if c_l2 and c_text:
        # все кандидаты по каждой теме (обезличенные, по убыванию длины)
        cand_map: Dict[str, Tuple[str, List[str]]] = {}
        for l2, sub in df.dropna(subset=[c_l2]).groupby(c_l2, sort=False):
            cands = _quote_candidates(sub[c_text])
            if cands:
                cand_map[_match_key(l2)] = (str(l2), cands)

        # выбор одного комментария на тему: LLM (если дан) → иначе самый длинный
        chosen: Dict[str, str] = {}
        if comment_picker and cand_map:
            try:
                picked = comment_picker([
                    {"theme": o, "candidates": c} for (o, c) in cand_map.values()
                ]) or {}
                for k, (o, c) in cand_map.items():
                    ch = picked.get(o)
                    # валидируем: должно быть одним из кандидатов (без галлюцинаций)
                    chosen[k] = ch if (isinstance(ch, str) and ch in c) else c[0]
            except Exception as e:  # noqa: BLE001
                logger.warning("LLM-выбор комментариев не удался (%s) — беру самый длинный", e)
                chosen = {}
        if not chosen:
            chosen = {k: c[0] for k, (o, c) in cand_map.items()}
        comment_index = {k: (cand_map[k][0], chosen[k]) for k in chosen}

    def _comment_for(theme: str) -> Optional[str]:
        if not comment_index:
            return None
        mk = _match_key(theme)
        if mk in comment_index:
            return comment_index[mk][1]
        if len(mk) >= 10:
            for k, (o, q) in comment_index.items():
                if len(k) >= 10 and (mk in k or k in mk):
                    return q
        best_sc, best_q = 0.0, None
        for k, (o, q) in comment_index.items():
            sc = _similarity(theme, o)
            if sc > best_sc:
                best_sc, best_q = sc, q
        return best_q if best_sc >= 0.5 else None

    # ---- «динамика» = таблица проблем: продукт, тема, текущая, прошлая, % ----
    if not dyn_sheet:
        raise ValueError("Для обзор-слайда нужен лист «динамика» — не нашёл его.")
    dd = _read_sheet(xls, dyn_sheet,
                     ["продукт", "тема", "динамик", "недел", "сигнал", "кол-во"])
    dprod = mapping.get("dyn_product") or _find_col(dd, DYN_HINTS["product"])
    dk = mapping.get("dyn_key") or _find_col(dd, DYN_HINTS["key"])
    dcur = mapping.get("dyn_current") or _find_col(dd, DYN_HINTS["current"])
    dprev = mapping.get("dyn_previous") or _find_col(dd, DYN_HINTS["previous"])
    dp = mapping.get("dyn_pct") or _find_col(dd, DYN_HINTS["pct"])
    dstatus = mapping.get("dyn_status") or _find_col(dd, DYN_HINTS["status"])
    report.update(dyn_product=dprod, dyn_key=dk, dyn_current=dcur,
                  dyn_previous=dprev, dyn_pct=dp, dyn_status=dstatus,
                  _dyn_columns=list(dd.columns))
    if dk is None:
        raise ValueError(
            f"В листе «{dyn_sheet}» не нашёл колонку темы. Колонки: {list(dd.columns)}. "
            "Задай mapping={'dyn_key': '<имя>'}."
        )
    # доли (0.22) или проценты (22) в колонке «динамика» — решаем по всей колонке
    rawvals = []
    if dp is not None:
        for _, row in dd.iterrows():
            v = _parse_num(row[dp])
            if v is not None:
                rawvals.append(v)
    as_fraction = bool(rawvals) and max(abs(v) for v in rawvals) <= 1.5

    groups_map: Dict[str, List[dict]] = {}
    n_topics = 0
    new_count = 0
    no_comment: List[str] = []
    for _, row in dd.iterrows():
        theme = str(row[dk]).strip() if pd.notna(row[dk]) else ""
        if not theme:
            continue
        prod = str(row[dprod]).strip() if (dprod and pd.notna(row[dprod])) else "Прочее"
        cur = _parse_num(row[dcur]) if dcur else None
        prev = _parse_num(row[dprev]) if dprev else None
        has_cur = cur is not None
        has_prev = prev is not None
        pct = None
        if has_cur and has_prev and prev != 0:
            pct = (float(cur) - float(prev)) / float(prev) * 100.0   # знаковый %
        elif dp is not None:
            v = _parse_num(row[dp])
            if v is not None:
                pct = v * 100.0 if as_fraction else v
        is_new = pct is None and has_cur and not has_prev
        status_txt = (str(row[dstatus]).strip()[:40]
                      if (dstatus and pd.notna(row[dstatus]) and str(row[dstatus]).strip())
                      else None)
        if status_txt and "нов" in status_txt.lower():
            is_new = True
        if is_new:
            new_count += 1
        quote = _comment_for(theme)
        if not quote:
            no_comment.append(theme[:60])
        groups_map.setdefault(prod, []).append(dict(
            title=theme[:160], mentions=int(round(cur)) if has_cur else 0,
            prev=int(round(prev)) if has_prev else None,
            pct=pct, is_new=is_new, quote=quote,
            status=status_txt or ("new" if is_new else None),
        ))
        n_topics += 1

    groups: List = []
    for prod, items in groups_map.items():
        items.sort(key=lambda d: -d["mentions"])
        ov_topics = [
            OverviewTopic(title=d["title"], quote=d["quote"], mentions=d["mentions"],
                          dynamics_pct=d["pct"],
                          status=d.get("status") or ("new" if d["is_new"] else None))
            for d in items[:max_topics_per_group]
        ]
        gmax = max((t.mentions for t in ov_topics), default=0)
        gtotal = sum(t.mentions for t in ov_topics)
        groups.append((gmax, gtotal, ProductGroup(name=str(prod)[:80], topics=ov_topics)))
    # сначала группы с самой «горячей» темой (макс. упоминаний), затем по сумме
    groups.sort(key=lambda x: (-x[0], -x[1]))
    groups = [g for _, _, g in groups[:max_groups]]

    # ---- источники: ВСЕГДА фиксированные блоки (по требованию) ----
    source_blocks = [
        SourceBlock(title=t, tags=list(tags)) for t, tags in FIXED_SOURCE_BLOCKS
    ]

    # ---- KPI: 3 из «исх», «новые» — из «динамика» ----
    total_signals = len(df)
    n_sources = df[c_src].nunique() if c_src else 0
    active_themes = df[c_l2].nunique() if c_l2 else n_topics
    kpis = [
        OverviewKPI(value=str(int(total_signals)), label="Сигналов проанализировано", icon_hint="signal"),
        OverviewKPI(value=str(int(n_sources)), label="Источников обратной связи", icon_hint="search"),
        OverviewKPI(value=str(int(active_themes)), label="Активных тем", icon_hint="growth"),
        OverviewKPI(value=str(int(new_count)), label="Новых тем", icon_hint="info"),
    ]

    # ---- служебные поля из «динамика»: период и номер дайджеста ----
    period_val = _scan_labeled_value(xls, dyn_sheet, ["период"])
    issue_val = _scan_labeled_value(
        xls, dyn_sheet, ["номер дайдж", "№ дайдж", "номер выпуск", "номер дайджеста"])
    report["_period"] = period_val
    report["_issue_number"] = issue_val

    overview = OverviewSlide(
        title=title, subtitle=subtitle, kpis=kpis,
        source_blocks=source_blocks, groups=groups,
    )
    # Прозрачность для проверки без файла
    report["_rows_read"] = int(len(df))
    report["_sources_extracted"] = sorted(
        {str(s).strip() for s in df[c_src].dropna()} if c_src else set()
    )
    report["_groups_extracted"] = [
        (g.name, [(t.title, t.mentions, t.dynamics_pct) for t in g.topics]) for g in groups
    ]
    report["_topics_total"] = n_topics
    report["_dyn_matched"] = sum(1 for g in groups for t in g.topics if t.dynamics_pct is not None)
    report["_no_comment"] = no_comment
    report["_dyn_sample"] = [
        (g.name, t.title[:40], t.mentions, t.dynamics_pct)
        for g in groups[:3] for t in g.topics[:2]
    ]
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
             "dyn_key", "dyn_product", "dyn_current", "dyn_previous", "dyn_pct",
             "dyn_status"]
    human = {
        "cluster_l1": "группа/продукт (кластер 1)",
        "cluster_l2": "тема (кластер 2)",
        "source": "источник", "source_block": "блок источника",
        "text": "текст/цитата", "status": "статус",
        "count": "кол-во упоминаний", "is_new": "признак новизны",
        "dyn_key": "'динамика': ключ-тема", "dyn_product": "'динамика': продукт=группа",
        "dyn_current": "'динамика': кол-во тек. неделя", "dyn_previous": "'динамика': кол-во прош. неделя",
        "dyn_pct": "'динамика': готовый %",
        "dyn_status": "'динамика': СТАТУС (для значков)",
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
        total = report.get("_topics_total")
        suffix = ""
        if matched is not None and total is not None:
            suffix = f" (проценты у {matched} из {total} тем; остальные — новые)"
        lines.append("  темы по продуктам из листа «динамика»" + suffix + ":")
        for gname, topics in groups:
            lines.append(f"    ▸ {gname}")
            for item in topics:
                if isinstance(item, tuple):
                    title, ment, pct = item
                    p = "new" if pct is None else f"{pct:+.0f}%"
                    lines.append(f"        – {title}  [{ment} упом., {p}]")
                else:
                    lines.append(f"        – {item}")
    no_comment = report.get("_no_comment")
    if no_comment:
        lines.append(f"  БЕЗ комментария из «исх» — {len(no_comment)} "
                     "(не нашёл похожую тему в «исх» для цитаты):")
        for t in no_comment:
            lines.append(f"    ✗ {t}")
    return "\n".join(lines)
