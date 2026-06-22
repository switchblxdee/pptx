"""
Готовые палитры-пресеты и детекция темы из промпта пользователя.

Проблема: LLM плохо слушает указания про цвета («сделай тёмное оформление»)
и копирует палитру из few-shot примера. Решение: детектируем намерение
пользователя по ключевым словам и подставляем готовую палитру программно.

Так стиль из запроса применяется ГАРАНТИРОВАННО, а LLM отвечает только
за контент.

Каждый пресет — полный набор из 8 цветов под схему ColorPalette:
  gradient_start, gradient_end — градиент фона
  card_bg   — фон карточек с темами
  kpi_bg    — фон KPI-карточек
  text_dark — основной текст (на тёмных темах это СВЕТЛЫЙ цвет!)
  text_muted — приглушённый текст
  accent    — акцент (заголовки, полосы)
  badge     — бейдж «new» / высокий приоритет

ВАЖНО про контраст: text_dark — это «цвет основного текста», не обязательно
тёмный. На тёмном фоне сюда кладётся светлый цвет, иначе текст не виден.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional


@dataclass
class ThemePreset:
    """Готовая палитра (8 цветов) + ключевые слова для детекции.

    Опционально несёт брендовую типографику и фоновые картинки-ассеты
    (например для фирменного шаблона SberF1).
    """
    name: str
    keywords: List[str]
    palette: Dict[str, str]
    heading_font: Optional[str] = None
    body_font: Optional[str] = None
    bg_cover: Optional[str] = None
    bg_content: Optional[str] = None


# --------------------------------------------------------------------------- #
# Пресеты палитр. Порядок не критичен — детект выбирает по длине совпавшего
# ключа (более специфичные ключи побеждают общие).
# --------------------------------------------------------------------------- #

THEME_PRESETS: List[ThemePreset] = [
    # ---- ТЁМНЫЕ ----
    ThemePreset(
        name="black",
        keywords=["чёрн", "черн", "black", "максимально тёмн", "очень тёмн",
                  "очень темн", "амолед", "amoled", "глубокий чёрный"],
        palette={
            "gradient_start": "0A0A0C", "gradient_end": "141417",
            "card_bg": "1C1C20", "kpi_bg": "242428",
            "text_dark": "FFFFFF", "text_muted": "A0A0A8",
            "accent": "6C9EFF", "badge": "FF6B6B",
        },
    ),
    ThemePreset(
        name="dark",
        keywords=["тёмн", "темн", "dark", "ночн", "мрачн", "dark mode",
                  "тёмная тема", "темная тема", "графит", "тёмное оформление",
                  "тёмные тона", "в тёмных", "в темных"],
        palette={
            "gradient_start": "121420", "gradient_end": "1C2030",
            "card_bg": "1F2230", "kpi_bg": "282C3C",
            "text_dark": "F5F6FA", "text_muted": "9BA1B0",
            "accent": "5B8DEF", "badge": "FF6B6B",
        },
    ),
    ThemePreset(
        name="dark_emerald",
        keywords=["тёмно-зелён", "темно-зелен", "изумруд", "emerald",
                  "тёмный зелёный", "малахит"],
        palette={
            "gradient_start": "0F1B14", "gradient_end": "1A2E22",
            "card_bg": "16271D", "kpi_bg": "1E3329",
            "text_dark": "EAF5EE", "text_muted": "8FA89A",
            "accent": "2ECC71", "badge": "F39C12",
        },
    ),
    ThemePreset(
        name="dark_purple",
        keywords=["тёмно-фиолет", "темно-фиолет", "тёмно-сиреневый",
                  "dark purple", "неон", "neon", "киберпанк", "cyberpunk"],
        palette={
            "gradient_start": "13101F", "gradient_end": "231A38",
            "card_bg": "1D1730", "kpi_bg": "281E40",
            "text_dark": "F2ECFF", "text_muted": "A99BC4",
            "accent": "B388FF", "badge": "FF5C8A",
        },
    ),
    ThemePreset(
        name="midnight_blue",
        keywords=["полуночн", "midnight", "тёмно-син", "темно-син",
                  "глубокий синий", "indigo", "индиго"],
        palette={
            "gradient_start": "0B1026", "gradient_end": "16203C",
            "card_bg": "141B33", "kpi_bg": "1D2746",
            "text_dark": "EAF0FF", "text_muted": "8A97B8",
            "accent": "4D8DFF", "badge": "FF7A59",
        },
    ),

    # ---- КОРПОРАТИВНЫЕ / ДЕЛОВЫЕ ----
    ThemePreset(
        name="navy_corporate",
        keywords=["корпоратив", "corporate", "деловой", "официальн",
                  "совет директоров", "navy", "премиум", "premium",
                  "представительск", "солидн"],
        palette={
            "gradient_start": "0B1F3A", "gradient_end": "1B3A5C",
            "card_bg": "F0F4F8", "kpi_bg": "FFFFFF",
            "text_dark": "0B1F3A", "text_muted": "5A6B7D",
            "accent": "C9A227", "badge": "E07A5F",
        },
    ),
    ThemePreset(
        name="slate_pro",
        keywords=["сланец", "slate", "сталь", "steel", "технологичн",
                  "tech", "айти", "it-стиль", "инженерн"],
        palette={
            "gradient_start": "2C3540", "gradient_end": "4A5560",
            "card_bg": "EEF1F4", "kpi_bg": "FFFFFF",
            "text_dark": "1E262E", "text_muted": "667380",
            "accent": "00A8A8", "badge": "FF6B5B",
        },
    ),
    ThemePreset(
        name="burgundy_lux",
        keywords=["бордов", "burgundy", "винн", "марсала", "люкс", "luxury",
                  "роскошн", "благородн"],
        palette={
            "gradient_start": "3B0D1F", "gradient_end": "5C1A35",
            "card_bg": "FBEEF2", "kpi_bg": "FFFFFF",
            "text_dark": "3B0D1F", "text_muted": "8A5A6B",
            "accent": "C0392B", "badge": "D4AF37",
        },
    ),

    # ---- ЯРКИЕ / ЭНЕРГИЧНЫЕ ----
    ThemePreset(
        name="vibrant",
        keywords=["яркий", "яркое", "яркая", "энергичн", "vibrant", "сочн",
                  "стартап", "startup", "молодёжн", "молодежн", "креатив"],
        palette={
            "gradient_start": "6C5CE7", "gradient_end": "FD79A8",
            "card_bg": "FFF0F6", "kpi_bg": "FFFFFF",
            "text_dark": "2D3436", "text_muted": "636E72",
            "accent": "00B894", "badge": "FDCB6E",
        },
    ),
    ThemePreset(
        name="coral_pop",
        keywords=["коралл", "coral", "розов", "pink", "фуксия", "малинов"],
        palette={
            "gradient_start": "FF7E5F", "gradient_end": "FEB47B",
            "card_bg": "FFF1EC", "kpi_bg": "FFFFFF",
            "text_dark": "5A2A27", "text_muted": "9C6F66",
            "accent": "E84393", "badge": "6C5CE7",
        },
    ),
    ThemePreset(
        name="lime_fresh",
        keywords=["лайм", "lime", "салатов", "ярко-зелён", "кислотн"],
        palette={
            "gradient_start": "A8E063", "gradient_end": "56AB2F",
            "card_bg": "F2FBE8", "kpi_bg": "FFFFFF",
            "text_dark": "1E3A0F", "text_muted": "5E7A4A",
            "accent": "27AE60", "badge": "FF7043",
        },
    ),

    # ---- ХОЛОДНЫЕ ----
    ThemePreset(
        name="ocean",
        keywords=["океан", "ocean", "морск", "синий", "blue", "голуб",
                  "вода", "аква", "aqua", "лазурн"],
        palette={
            "gradient_start": "2980B9", "gradient_end": "6DD5FA",
            "card_bg": "EAF6FB", "kpi_bg": "FFFFFF",
            "text_dark": "0A3D62", "text_muted": "5B7A8C",
            "accent": "0984E3", "badge": "FF7675",
        },
    ),
    ThemePreset(
        name="mint",
        keywords=["мят", "mint", "бирюзов", "teal", "тиффани", "tiffany",
                  "ментол"],
        palette={
            "gradient_start": "76E3CE", "gradient_end": "C2F0E3",
            "card_bg": "EDFBF6", "kpi_bg": "FFFFFF",
            "text_dark": "0E4D43", "text_muted": "5A8077",
            "accent": "16A085", "badge": "FF8A65",
        },
    ),
    ThemePreset(
        name="lavender",
        keywords=["лаванд", "lavender", "сиренев", "фиолетов", "purple",
                  "violet", "пудров"],
        palette={
            "gradient_start": "B79CED", "gradient_end": "E0C3FC",
            "card_bg": "F6F0FF", "kpi_bg": "FFFFFF",
            "text_dark": "3D2A5C", "text_muted": "7A6A95",
            "accent": "8E44AD", "badge": "FF6B9D",
        },
    ),

    # ---- ТЁПЛЫЕ ----
    ThemePreset(
        name="warm",
        keywords=["тёпл", "тепл", "warm", "оранж", "orange", "осен",
                  "закат", "sunset", "янтар"],
        palette={
            "gradient_start": "F8B500", "gradient_end": "FCEEB5",
            "card_bg": "FFF6E5", "kpi_bg": "FFFFFF",
            "text_dark": "5C3A21", "text_muted": "977C5E",
            "accent": "E17055", "badge": "D63031",
        },
    ),
    ThemePreset(
        name="terracotta",
        keywords=["терракот", "terracotta", "глин", "кирпич", "охра",
                  "земл", "earthy", "тёпло-коричнев"],
        palette={
            "gradient_start": "C96E4A", "gradient_end": "E8B796",
            "card_bg": "FBEDE3", "kpi_bg": "FFFFFF",
            "text_dark": "4A2618", "text_muted": "8A6453",
            "accent": "B0492B", "badge": "5B8C5A",
        },
    ),
    ThemePreset(
        name="cream_beige",
        keywords=["беж", "beige", "крем", "cream", "молочн", "айвори",
                  "ivory", "пастельн тёпл", "нюд", "nude"],
        palette={
            "gradient_start": "EDE0D0", "gradient_end": "F8F1E7",
            "card_bg": "FFFBF5", "kpi_bg": "FFFFFF",
            "text_dark": "4A3F35", "text_muted": "9A8A7A",
            "accent": "B08968", "badge": "C77966",
        },
    ),

    # ---- МИНИМАЛИЗМ / НЕЙТРАЛЬНЫЕ ----
    ThemePreset(
        name="minimal_mono",
        keywords=["минимал", "minimal", "монохром", "mono", "ч/б",
                  "чёрно-бел", "черно-бел", "строг", "сдержан", "лаконичн"],
        palette={
            "gradient_start": "F5F5F5", "gradient_end": "E0E0E0",
            "card_bg": "FFFFFF", "kpi_bg": "FAFAFA",
            "text_dark": "1A1A1A", "text_muted": "757575",
            "accent": "424242", "badge": "E53935",
        },
    ),
    ThemePreset(
        name="light_clean",
        keywords=["светл", "light", "белый фон", "чистый", "clean",
                  "воздушн", "лёгк", "легк"],
        palette={
            "gradient_start": "FFFFFF", "gradient_end": "EEF2F6",
            "card_bg": "F4F7FA", "kpi_bg": "FFFFFF",
            "text_dark": "1F2933", "text_muted": "7B8794",
            "accent": "3D7DCA", "badge": "F0883E",
        },
    ),

    # ---- ФИРМЕННЫЙ ШАБЛОН SberF1 ----
    ThemePreset(
        name="sberf1",
        keywords=["sberf1", "sber f1", "сберф1", "сбер f1", "фирменн",
                  "брендов", "бренд-стиль", "шаблон sber", "шаблон сбер",
                  "корпоративный шаблон", "f1", "ф1", "блок технологии",
                  "командасбера", "команда сбера"],
        palette={
            # фон контентных слайдов — светлый брендовый меш (картинка),
            # поэтому стопы держим светлыми: текст на фоне выходит тёмным
            "gradient_start": "FAFAF7", "gradient_end": "FFFFFF",
            "card_bg": "FFFFFF", "kpi_bg": "111827",   # тёмные KPI как в шаблоне
            "text_dark": "111827", "text_muted": "64748B",
            "accent": "0669E0", "badge": "933EFF",     # брендовые синий/фиолет
        },
        heading_font="SB Sans Display",
        body_font="SB Sans Display",
        bg_cover="bg_cover.png",
        bg_content="bg_content.png",
    ),


    ThemePreset(
        name="default_teal_peach",
        keywords=["персик", "peach", "пастель", "голос it", "референс",
                  "как образец", "стандартн"],
        palette={
            "gradient_start": "A8D8D5", "gradient_end": "F5D5BA",
            "card_bg": "FFE8EA", "kpi_bg": "FFFFFF",
            "text_dark": "1A1A2E", "text_muted": "6B7280",
            "accent": "2C5F5D", "badge": "F97316",
        },
    ),
]

DEFAULT_PRESET = THEME_PRESETS[-1]  # бирюзово-персиковый


def detect_theme(style_prompt: str) -> Optional[ThemePreset]:
    """
    Определяет пресет по ключевым словам в промпте.

    Возвращает пресет с самым ДЛИННЫМ совпавшим ключом (специфичные
    побеждают общие: «тёмно-зелёный» → dark_emerald, а не dark).
    None, если ничего не совпало.
    """
    if not style_prompt:
        return None

    text = style_prompt.lower().replace("ё", "ё")  # нормализация не трогает ё
    best: Optional[ThemePreset] = None
    best_key_len = 0

    for preset in THEME_PRESETS:
        for kw in preset.keywords:
            if kw in text and len(kw) > best_key_len:
                best = preset
                best_key_len = len(kw)

    return best


def resolve_palette(style_prompt: str, fallback_to_default: bool = True) -> Optional[Dict[str, str]]:
    """Палитра для промпта. fallback_to_default=False → None, если не нашли."""
    preset = detect_theme(style_prompt)
    if preset:
        return dict(preset.palette)
    if fallback_to_default:
        return dict(DEFAULT_PRESET.palette)
    return None


def get_theme_by_name(name: str) -> Optional[ThemePreset]:
    """Пресет по точному имени (для force_theme). Регистронезависимо."""
    if not name:
        return None
    key = name.lower().strip()
    for preset in THEME_PRESETS:
        if preset.name.lower() == key:
            return preset
    return None


def available_themes() -> List[str]:
    """Список имён доступных тем."""
    return [p.name for p in THEME_PRESETS]
