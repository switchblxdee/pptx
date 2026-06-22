"""
style_resolver.py — разбор пожеланий по стилю ПРЯМО из промпта.

Зачем: старый detect_theme искал подстроки ("чёрн" → чёрная тема) и
перетирал палитру целиком. Из-за этого "тёмный фон и ТЁМНЫЙ ТЕКСТ" ломался:
слово "чёрн" в "чёрный текст" включало тему с чёрным ФОНОМ и белым текстом.

Здесь мы разбираем директивы по РОЛЯМ отдельно: что относится к фону, что —
к тексту, акценту, карточкам, бейджу. Поддерживаем:
  - имена цветов (рус/англ): "чёрный", "белый", "оранжевый", "navy", ...
  - hex: "#101010", "1A1A2E"
  - модификаторы: "тёмный", "светлый", "яркий", "приглушённый"
  - роль без цвета: "тёмный фон" (нет имени цвета) → нейтральный тёмный.

Возвращаем:
  overrides — {slot_hex: ...} для слотов ColorPalette,
  locked    — множество РОЛЕЙ, заданных пользователем явно. Эти роли
              билдер не трогает авто-контрастом: "делай как просил".

Это не заменяет detect_theme как БАЗУ (она задаёт согласованную палитру,
когда явных директив нет), но имеет приоритет над ней.
"""
from __future__ import annotations

import re
from typing import Dict, Set, Tuple

# --------------------------------------------------------------------------- #
# Словарь цветов (рус + англ) -> hex
# --------------------------------------------------------------------------- #

COLOR_NAMES: Dict[str, str] = {
    # ключи — ОСНОВЫ (стемы), чтобы ловить склонения:
    # "чёрный/чёрным/чёрная/чёрное" → "чёрн"
    # базовые
    "чёрн": "0A0A0C", "черн": "0A0A0C", "black": "0A0A0C",
    "бел": "FFFFFF", "white": "FFFFFF",
    "сер": "808080", "gray": "808080", "grey": "808080",
    "графит": "2C3540", "charcoal": "2C3540",
    # тёплые
    "красн": "E2362D", "red": "E2362D",
    "оранж": "F97316", "orange": "F97316",
    "жёлт": "F5C518", "желт": "F5C518", "yellow": "F5C518",
    "золот": "C9A227", "gold": "C9A227",
    "коралл": "FF7E5F", "coral": "FF7E5F",
    "бордов": "5C1A35", "burgundy": "5C1A35", "винн": "5C1A35",
    "терракот": "C96E4A", "terracotta": "C96E4A",
    "бежев": "EDE0D0", "беж": "EDE0D0", "beige": "EDE0D0", "крем": "F8F1E7",
    "коричнев": "5C3A21", "brown": "5C3A21",
    # холодные
    "син": "1E5BFF", "blue": "1E5BFF", "navy": "0B1F3A", "indigo": "16203C",
    "голуб": "6DD5FA", "lightblue": "6DD5FA",
    "бирюз": "16A085", "teal": "16A085",
    "мятн": "76E3CE", "мята": "76E3CE", "mint": "76E3CE",
    "зелён": "27AE60", "зелен": "27AE60", "green": "27AE60",
    "изумруд": "2ECC71", "emerald": "2ECC71",
    "лайм": "A8E063", "lime": "A8E063",
    # фиолетовые/розовые
    "фиолет": "8E44AD", "purple": "8E44AD", "violet": "8E44AD",
    "сирен": "B79CED", "лаванд": "B79CED", "lavender": "B79CED",
    "розов": "FF6B9D", "pink": "FF6B9D", "фукси": "E84393",
}

# --------------------------------------------------------------------------- #
# Роли и их триггеры в тексте
# --------------------------------------------------------------------------- #

ROLE_KEYWORDS: Dict[str, Tuple[str, ...]] = {
    "background": ("фон", "подложк", "задник", "background"),
    "text": ("текст", "шрифт", "буквы", "надпис", "text", "font"),
    "text_muted": ("приглушён", "приглушен", "вторичн", "подзаголов", "muted"),
    "accent": ("акцент", "выделен", "accent"),
    "card": ("карточк", "плашк", "card"),
    "kpi": ("kpi", "кпи"),
    "badge": ("бейдж", "значок", "метк", "badge"),
}

# роль -> слоты ColorPalette, которые она задаёт
ROLE_TO_SLOTS: Dict[str, Tuple[str, ...]] = {
    "background": ("gradient_start", "gradient_end"),
    "text": ("text_dark",),
    "text_muted": ("text_muted",),
    "accent": ("accent",),
    "card": ("card_bg",),
    "kpi": ("kpi_bg",),
    "badge": ("badge",),
}

# модификаторы светлоты (нет имени цвета — берём нейтраль нужной светлоты)
DARK_WORDS = ("тёмн", "темн", "dark", "глубок", "ночн")
LIGHT_WORDS = ("светл", "light", "бледн", "пастельн", "воздушн")
VIVID_WORDS = ("ярк", "сочн", "насыщен", "vivid", "bright")
MUTED_WORDS = ("приглушён", "приглушен", "спокойн", "мягк", "muted", "pale")

NEUTRAL_DARK = "121420"
NEUTRAL_LIGHT = "F4F7FA"

_HEX_RE = re.compile(r"#?\b([0-9a-fA-F]{6})\b")
_WORD_RE = re.compile(r"[#0-9a-zA-Zа-яёА-ЯЁ]+", re.UNICODE)


# --------------------------------------------------------------------------- #
# Вспомогательные операции над цветом (без зависимостей)
# --------------------------------------------------------------------------- #

def _rgb(hexs: str) -> Tuple[int, int, int]:
    h = hexs.lstrip("#")
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


def _hex(rgb: Tuple[int, int, int]) -> str:
    return "".join(f"{max(0, min(255, int(round(c)))):02X}" for c in rgb)


def _mix(a: str, b: str, t: float) -> str:
    ra, ga, ba = _rgb(a)
    rb, gb, bb = _rgb(b)
    return _hex((ra + (rb - ra) * t, ga + (gb - ga) * t, ba + (bb - ba) * t))


def _apply_modifier(color: str, mod: str | None) -> str:
    """Двигает светлоту/насыщенность цвета по словесному модификатору."""
    if mod is None:
        return color
    if mod == "dark":
        return _mix(color, "000000", 0.45)
    if mod == "light":
        return _mix(color, "FFFFFF", 0.55)
    if mod == "vivid":
        # к насыщенному: чуть к чистому цвету (упрощённо — лёгкое затемнение
        # к среднему тону, чтобы не белёсо)
        return _mix(color, "000000", 0.1)
    if mod == "muted":
        return _mix(color, "808080", 0.4)
    return color


def _detect_modifier(tokens: list[str]) -> str | None:
    blob = " ".join(tokens)
    if any(w in blob for w in DARK_WORDS):
        return "dark"
    if any(w in blob for w in LIGHT_WORDS):
        return "light"
    if any(w in blob for w in VIVID_WORDS):
        return "vivid"
    if any(w in blob for w in MUTED_WORDS):
        return "muted"
    return None


def _find_color(tokens: list[str]) -> str | None:
    """Ищет hex или имя цвета в токенах клаузы. hex имеет приоритет,
    затем самое длинное совпавшее имя цвета (подстрокой)."""
    for tok in tokens:
        m = _HEX_RE.fullmatch(tok) or _HEX_RE.match(tok)
        if m:
            return m.group(1).upper()
    blob = " ".join(tokens).lower()
    best, best_len = None, 0
    for name, hx in COLOR_NAMES.items():
        if name in blob and len(name) > best_len:
            best, best_len = hx, len(name)
    return best


def _role_at(token: str) -> str | None:
    t = token.lower()
    for role, kws in ROLE_KEYWORDS.items():
        if any(kw in t for kw in kws):
            return role
    return None


# --------------------------------------------------------------------------- #
# Главный разбор
# --------------------------------------------------------------------------- #

_CLAUSE_SPLIT_RE = re.compile(r"[,.;:!?]|\bи\b|\bс\b|\bа\s+также\b|\+", re.UNICODE)


def parse_style_directives(prompt: str) -> Tuple[Dict[str, str], Set[str]]:
    """
    Разбирает промпт на (overrides слотов ColorPalette, множество ролей-замков).

    Промпт режется на КЛАУЗЫ по запятым/«и»/«с»/«+». В каждой клаузе ищем
    роль ("фон","текст","акцент"...), имя цвета/hex и модификатор светлоты —
    и назначаем строго в пределах клаузы. Так "белый фон, чёрный текст" даёт
    фон белый + текст чёрный без перетекания цветов между ролями.
    """
    if not prompt:
        return {}, set()

    overrides: Dict[str, str] = {}
    locked: Set[str] = set()
    pending_bg_mod: str | None = None  # модификатор светлоты без явной роли

    for clause in _CLAUSE_SPLIT_RE.split(prompt):
        clause = clause.strip()
        if not clause:
            continue
        tokens = _WORD_RE.findall(clause)
        if not tokens:
            continue

        role = _role_in_clause(tokens)
        color = _find_color(tokens)
        mod = _detect_modifier(tokens)

        if role is None:
            # роли нет: запомним модификатор светлоты для фона
            # ("светлая презентация" → светлый фон, если фон явно не задан)
            if mod in ("dark", "light"):
                pending_bg_mod = mod
            continue

        if color is None and mod is None:
            continue

        if color is None:
            if mod == "dark":
                color = NEUTRAL_DARK
            elif mod == "light":
                color = NEUTRAL_LIGHT
            else:
                continue
        else:
            color = _apply_modifier(color, mod)

        _assign_role(overrides, role, color)
        locked.add(role)

    # Применяем отложенный модификатор фона, если фон не задан явно
    if pending_bg_mod and "background" not in locked:
        color = NEUTRAL_DARK if pending_bg_mod == "dark" else NEUTRAL_LIGHT
        _assign_role(overrides, "background", color)
        locked.add("background")

    return overrides, locked


def _role_in_clause(tokens: list[str]) -> str | None:
    """Роль клаузы — по самому длинному совпавшему ключевому слову
    (специфичные роли важнее общих, напр. text_muted важнее text)."""
    best_role = None
    best_len = 0
    blob = " ".join(tokens).lower()
    for role, kws in ROLE_KEYWORDS.items():
        for kw in kws:
            if kw in blob and len(kw) > best_len:
                best_role, best_len = role, len(kw)
    return best_role


def _assign_role(overrides: Dict[str, str], role: str, color: str) -> None:
    slots = ROLE_TO_SLOTS.get(role, ())
    if role == "background":
        # лёгкий градиент от заданного цвета: второй стоп чуть смещён
        start = color
        # светлый фон -> второй стоп чуть темнее, тёмный -> чуть светлее
        end = _mix(color, "000000", 0.10) if _is_light(color) else _mix(color, "FFFFFF", 0.10)
        overrides["gradient_start"] = start
        overrides["gradient_end"] = end
        return
    for slot in slots:
        overrides[slot] = color


def _is_light(color: str) -> bool:
    r, g, b = _rgb(color)
    # быстрая оценка яркости
    return (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255.0 > 0.5


def merge_palette(base: Dict[str, str], overrides: Dict[str, str]) -> Dict[str, str]:
    """Накладывает overrides поверх базовой палитры (копия)."""
    merged = dict(base)
    merged.update(overrides)
    return merged
