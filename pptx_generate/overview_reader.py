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


def _pick_quote(series) -> Optional[str]:
    """Самый содержательный комментарий из набора, схлопнутый и обрезанный."""
    cands = []
    for v in series.dropna():
        s = " ".join(str(v).split())
        if s:
            cands.append(s)
    if not cands:
        return None
    cands.sort(key=len, reverse=True)
    q = cands[0]
    return (q[:150].rstrip() + "…") if len(q) > 150 else q


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

    df = xls.parse(src_sheet)               # лист «исх»
    df.columns = [str(c) for c in df.columns]

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
        for l2, sub in df.dropna(subset=[c_l2]).groupby(c_l2, sort=False):
            q = _pick_quote(sub[c_text])
            if q:
                comment_index[_match_key(l2)] = (str(l2), q)

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
    dd = xls.parse(dyn_sheet)
    dd.columns = [str(c) for c in dd.columns]
    dprod = mapping.get("dyn_product") or _find_col(dd, DYN_HINTS["product"])
    dk = mapping.get("dyn_key") or _find_col(dd, DYN_HINTS["key"])
    dcur = mapping.get("dyn_current") or _find_col(dd, DYN_HINTS["current"])
    dprev = mapping.get("dyn_previous") or _find_col(dd, DYN_HINTS["previous"])
    dp = mapping.get("dyn_pct") or _find_col(dd, DYN_HINTS["pct"])
    report.update(dyn_product=dprod, dyn_key=dk, dyn_current=dcur,
                  dyn_previous=dprev, dyn_pct=dp, _dyn_columns=list(dd.columns))
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
        cur = pd.to_numeric(row[dcur], errors="coerce") if dcur else None
        prev = pd.to_numeric(row[dprev], errors="coerce") if dprev else None
        has_cur = cur is not None and pd.notna(cur)
        has_prev = prev is not None and pd.notna(prev)
        pct = None
        if has_cur and has_prev and prev != 0:
            pct = (float(cur) - float(prev)) / float(prev) * 100.0   # знаковый %
        elif dp is not None:
            v = _parse_num(row[dp])
            if v is not None:
                pct = v * 100.0 if as_fraction else v
        is_new = pct is None and has_cur and not has_prev
        if is_new:
            new_count += 1
        quote = _comment_for(theme)
        if not quote:
            no_comment.append(theme[:60])
        groups_map.setdefault(prod, []).append(dict(
            title=theme[:160], mentions=int(cur) if has_cur else 0,
            pct=pct, is_new=is_new, quote=quote,
        ))
        n_topics += 1

    groups: List = []
    for prod, items in groups_map.items():
        items.sort(key=lambda d: -d["mentions"])
        ov_topics = [
            OverviewTopic(title=d["title"], quote=d["quote"], mentions=d["mentions"],
                          dynamics_pct=d["pct"],
                          status=("new" if d["is_new"] else None))
            for d in items[:max_topics_per_group]
        ]
        gtotal = sum(t.mentions for t in ov_topics)
        groups.append((gtotal, ProductGroup(name=str(prod)[:80], topics=ov_topics)))
    groups.sort(key=lambda x: -x[0])
    groups = [g for _, g in groups[:max_groups]]

    # ---- источники из «исх» ----
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
