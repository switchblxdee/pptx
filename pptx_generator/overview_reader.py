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
from typing import Dict, List, Optional, Tuple

import pandas as pd

logger = logging.getLogger(__name__)

# Ключевые слова заголовков для каждой роли (нижний регистр, подстрока).
# Порядок важен: первое совпадение выигрывает. Можно расширять.
HINTS = {
    "cluster_l1": ["кластер 1", "кластер1", "кластер_1", "1 уровн", "первого уровн",
                   "продукт", "группа", "родител", "parent", "l1"],
    "cluster_l2": ["кластер 2", "кластер2", "кластер_2", "2 уровн", "второго уровн",
                   "тема", "подкластер", "child", "l2"],
    "source": ["источник", "канал", "чат", "source", "площадк"],
    "source_block": ["блок", "тип источ", "категория источ", "группа источ",
                     "вид источ", "сегмент", "раздел источ"],
    "text": ["текст", "сообщени", "цитат", "обращени", "feedback", "комментар",
             "отзыв", "реплик"],
    "status": ["статус", "боль", "анализ", "состояни", "метк"],
    "count": ["упомин", "кол-во", "количеств", "count", "частот", "число сигнал"],
    "is_new": ["нов", "new"],
}
DYN_HINTS = {
    "key": ["кластер 2", "кластер2", "2 уровн", "тема", "кластер", "название", "l2", "name"],
    "pct": ["%", "процент", "динамик", "изменени", "к прошл", "рост", "delta", "дельта"],
}


def _norm(s) -> str:
    return str(s).strip().lower().replace("ё", "е")


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

    # ---- агрегация тем по группам ----
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
        vals = [str(v).strip() for v in sub[c_text].dropna().tolist() if str(v).strip()]
        if not vals:
            return None
        # самая содержательная, но не слишком длинная
        vals.sort(key=lambda s: (-(len(s) <= 300), -len(s)))
        return vals[0][:380]

    def _status(sub) -> Optional[str]:
        if not c_status:
            return None
        vals = [str(v).strip() for v in sub[c_status].dropna().tolist() if str(v).strip()]
        return vals[0] if vals else None

    def _is_new(sub) -> bool:
        if c_new and c_new in sub:
            return bool(pd.to_numeric(sub[c_new], errors="coerce").fillna(0).sum() > 0)
        st = _status(sub) or ""
        return "нов" in st.lower() or "new" in st.lower()

    groups_map: Dict[str, List[Tuple[str, int, Optional[str], Optional[str], bool]]] = {}
    for (l1, l2), sub in df.groupby([c_l1, c_l2], sort=False):
        groups_map.setdefault(str(l1), []).append(
            (str(l2), _mentions(sub), _quote(sub), _status(sub), _is_new(sub))
        )

    # ---- динамика: тема -> % ----
    dyn: Dict[str, float] = {}
    if dyn_sheet:
        dd = xls.parse(dyn_sheet)
        dd.columns = [str(c) for c in dd.columns]
        dk = mapping.get("dyn_key") or _find_col(dd, DYN_HINTS["key"])
        dp = mapping.get("dyn_pct") or _find_col(dd, DYN_HINTS["pct"])
        report["dyn_key"] = dk
        report["dyn_pct"] = dp
        report["_dyn_columns"] = list(dd.columns)
        if dk and dp:
            for _, row in dd.iterrows():
                key = _norm(row[dk])
                val = pd.to_numeric(row[dp], errors="coerce")
                if key and pd.notna(val):
                    dyn[key] = float(val)

    # ---- источники по блокам ----
    source_blocks: List = []
    if c_src:
        if c_block:
            for block, sub in df.groupby(c_block, sort=False):
                tags = sorted({str(s).strip() for s in sub[c_src].dropna() if str(s).strip()})
                if tags:
                    source_blocks.append(SourceBlock(title=str(block), tags=tags))
            source_blocks.sort(key=lambda b: -len(b.tags))
            source_blocks = source_blocks[:2]
        else:
            tags = sorted({str(s).strip() for s in df[c_src].dropna() if str(s).strip()})
            source_blocks = [SourceBlock(title="Источники", tags=tags)]

    # ---- сборка групп/тем ----
    groups: List = []
    new_count = 0
    for gname, topics in groups_map.items():
        topics.sort(key=lambda t: -t[1])  # по упоминаниям
        ov_topics = []
        for (l2, ment, quote, status, isnew) in topics[:max_topics_per_group]:
            if isnew:
                new_count += 1
            ov_topics.append(OverviewTopic(
                title=l2, quote=quote, mentions=ment,
                dynamics_pct=dyn.get(_norm(l2)),
                status=("new" if isnew and not status else status),
            ))
        gtotal = sum(t.mentions for t in ov_topics)
        groups.append((gtotal, ProductGroup(name=gname, topics=ov_topics)))
    groups.sort(key=lambda x: -x[0])
    groups = [g for _, g in groups[:max_groups]]

    # ---- KPI ----
    total_signals = sum(_mentions(sub) for _, sub in df.groupby([c_l1, c_l2], sort=False))
    n_sources = df[c_src].nunique() if c_src else 0
    n_topics = df[c_l2].nunique()
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
             "text", "status", "count", "is_new", "dyn_key", "dyn_pct"]
    human = {
        "cluster_l1": "группа/продукт (кластер 1)",
        "cluster_l2": "тема (кластер 2)",
        "source": "источник", "source_block": "блок источника",
        "text": "текст/цитата", "status": "статус",
        "count": "кол-во упоминаний", "is_new": "признак новизны",
        "dyn_key": "ключ в 'динамика'", "dyn_pct": "% в 'динамика'",
    }
    for r in roles:
        col = report.get(r)
        mark = "✓" if col else "— НЕ НАЙДЕНО"
        lines.append(f"    {human[r]:32s}: {col or ''} {mark}")
    return "\n".join(lines)
