"""Распознавание цвета по названию (рус/англ) и из промпта."""
from __future__ import annotations

import re
from typing import Optional

# Название цвета -> HEX (подобраны достаточно насыщенные/читаемые тона)
COLOR_NAMES = {
    # фиолетовый
    "фиолетов": "7B61FF", "пурпурн": "7B61FF", "purple": "7B61FF", "violet": "7B61FF",
    "лавандов": "A78BFA", "lavender": "A78BFA", "сиренев": "9D7BEA",
    "индиго": "6366F1", "indigo": "6366F1",
    # синий/голубой
    "син": "0669E0", "blue": "0669E0", "голуб": "38BDF8", "cyan": "22B8CF",
    "navy": "1E3A8A", "тёмно-син": "1E3A8A", "темно-син": "1E3A8A",
    # бирюзовый/зелёный
    "бирюзов": "0B9B98", "teal": "0B9B98", "циан": "0B9B98",
    "зелён": "21A038", "зелен": "21A038", "green": "21A038",
    "салатов": "84CC16", "lime": "84CC16", "изумруд": "059669", "emerald": "059669",
    "мятн": "2DD4BF", "mint": "2DD4BF",
    # тёплые
    "красн": "E53935", "red": "E53935", "алый": "EF4444",
    "оранжев": "F59E0B", "orange": "F59E0B", "оранж": "F59E0B",
    "жёлт": "EAB308", "желт": "EAB308", "yellow": "EAB308", "золот": "D4A017", "gold": "D4A017",
    "розов": "EC4899", "pink": "EC4899", "малинов": "DB2777",
    "коралл": "FB7185", "coral": "FB7185",
    "персик": "F4C99A", "peach": "F4C99A",
    "бордов": "7B1E3B", "burgundy": "7B1E3B", "винн": "7B1E3B",
    "коричнев": "8B5E3C", "brown": "8B5E3C",
    # нейтральные
    "сер": "94A3B8", "gray": "94A3B8", "grey": "94A3B8",
    "чёрн": "111827", "черн": "111827", "black": "111827",
    "бел": "FFFFFF", "white": "FFFFFF",
}

_HEX_RE = re.compile(r"#?([0-9a-fA-F]{6})\b")

# слова-маркеры «фоновых объектов для текста»
_OBJ_MARKERS = (
    "фонов", "подложк", "плашк", "объекты для текст", "объекты заднего",
    "задний фон", "заднего фона", "фон объект", "background object",
)


def resolve_color_name(value: Optional[str]) -> Optional[str]:
    """Принимает HEX или название цвета (рус/англ) -> HEX без '#'. None если не цвет."""
    if not value:
        return None
    s = str(value).strip().lower()
    m = _HEX_RE.fullmatch(s) or _HEX_RE.match(s)
    if m and len(s.lstrip("#")) >= 6:
        return m.group(1).upper()
    for key, hexc in COLOR_NAMES.items():
        if key in s:
            return hexc
    return None


def extract_object_color(prompt: Optional[str]) -> Optional[str]:
    """Из промпта вида «сделай фоновые объекты фиолетовыми» достаёт HEX цвета.

    Срабатывает только если в тексте есть маркер «фоновых объектов» И название цвета.
    """
    if not prompt:
        return None
    low = prompt.lower()
    if not any(mk in low for mk in _OBJ_MARKERS):
        return None
    for key, hexc in COLOR_NAMES.items():
        if key in low:
            return hexc
    m = _HEX_RE.search(prompt)
    return m.group(1).upper() if m else None


# Нейтральные ключи: они НЕ годятся как акцентный цвет карточек
# (чёрный/белый/серый сделали бы карточки чёрными/белыми, а это не «тон»).
_NEUTRAL_KEYS = {"чёрн", "черн", "black", "бел", "white", "сер", "gray", "grey"}


def extract_accent_color(prompt: Optional[str]) -> Optional[str]:
    """Достаёт ЛЮБОЙ акцентный цвет из свободной фразы.

    В отличие от extract_object_color, не требует слова-маркера: ловит
    «в фиолетовых тонах», «фиолетовая презентация», «#7B61FF» и т.п.
    Нужен как фолбэк для overview, где единственный осмысленный цветовой
    рычаг — цвет фоновых объектов (фон обзора зафиксирован отдельно).
    Нейтральные (чёрный/белый/серый) игнорируются.
    """
    if not prompt:
        return None
    m = _HEX_RE.search(prompt)
    if m:
        return m.group(1).upper()
    low = prompt.lower()
    for key, hexc in COLOR_NAMES.items():
        if key in _NEUTRAL_KEYS:
            continue
        if key in low:
            return hexc
    return None
