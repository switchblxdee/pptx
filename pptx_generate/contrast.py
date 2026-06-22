"""
contrast.py — авто-контраст для дайджест-билдера.

Идея: ни один цвет текста/фигуры в презентации не должен задаваться «на глаз».
Каждый раз, когда мы кладём текст на фон (слайд, KPI-карточку, pill, плашку),
цвет текста вычисляется так, чтобы он гарантированно читался — по метрике WCAG.

Ничего не знает про python-pptx по сути (работает на чистых числах),
но умеет принимать/возвращать RGBColor, если pptx установлен — для удобной
интеграции в builder.py.

--------------------------------------------------------------------------- #
Немного теории (формулы — основа всего модуля)
--------------------------------------------------------------------------- #

1. sRGB-канал -> линейный (gamma decode):
       cs = c / 255
       c_lin = cs/12.92                      , если cs <= 0.04045
               ((cs + 0.055)/1.055) ** 2.4   , иначе

2. Относительная яркость (relative luminance), WCAG 2.x:
       L = 0.2126*R_lin + 0.7152*G_lin + 0.0722*B_lin
   Веса — это чувствительность глаза: зелёный самый «яркий», синий почти не
   добавляет воспринимаемой светлоты.

3. Контраст двух цветов:
       CR = (L_светлый + 0.05) / (L_тёмный + 0.05)
   Диапазон 1..21. Пороги WCAG:
       - 4.5 : обычный текст, уровень AA  (дефолт)
       - 3.0 : крупный текст (>=18pt bold / >=24pt) и UI-элементы
       - 7.0 : обычный текст, уровень AAA
"""
from __future__ import annotations

from typing import Iterable, Sequence, Tuple, Union

# RGBColor опционален — модуль работает и без pptx.
try:
    from pptx.dml.color import RGBColor  # type: ignore
    _HAS_PPTX = True
except Exception:  # pragma: no cover
    RGBColor = None  # type: ignore
    _HAS_PPTX = False


RGB = Tuple[int, int, int]
ColorLike = Union[str, RGB, Sequence[int], "RGBColor"]

# --------------------------------------------------------------------------- #
# Пороги контраста
# --------------------------------------------------------------------------- #

AA_NORMAL = 4.5
AA_LARGE = 3.0
AAA_NORMAL = 7.0
UI_COMPONENT = 3.0

BLACK: RGB = (0, 0, 0)
WHITE: RGB = (255, 255, 255)


# --------------------------------------------------------------------------- #
# Парсинг / нормализация цвета
# --------------------------------------------------------------------------- #

def parse_color(c: ColorLike) -> RGB:
    """Любой разумный формат -> (r, g, b) c int 0..255."""
    if c is None:
        raise ValueError("color is None")

    # RGBColor (наследник bytes длиной 3) и обычные кортежи/списки
    if _HAS_PPTX and isinstance(c, RGBColor):
        return (c[0], c[1], c[2])
    if isinstance(c, (tuple, list)) and len(c) == 3:
        r, g, b = (int(x) for x in c)
        return (_clamp8(r), _clamp8(g), _clamp8(b))

    if isinstance(c, str):
        s = c.strip().lstrip("#")
        if len(s) == 3:  # короткая форма #abc
            s = "".join(ch * 2 for ch in s)
        if len(s) != 6:
            raise ValueError(f"bad hex color: {c!r}")
        return (int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))

    raise ValueError(f"unsupported color: {c!r}")


def to_hex(c: ColorLike, with_hash: bool = True) -> str:
    r, g, b = parse_color(c)
    h = f"{r:02X}{g:02X}{b:02X}"
    return f"#{h}" if with_hash else h


def to_rgbcolor(c: ColorLike):
    """-> RGBColor для python-pptx (бросает, если pptx не установлен)."""
    if not _HAS_PPTX:
        raise RuntimeError("python-pptx is not installed")
    r, g, b = parse_color(c)
    return RGBColor(r, g, b)


def _clamp8(x: int) -> int:
    return max(0, min(255, int(round(x))))


# --------------------------------------------------------------------------- #
# Яркость и контраст
# --------------------------------------------------------------------------- #

def _linearize(channel: int) -> float:
    cs = channel / 255.0
    if cs <= 0.04045:
        return cs / 12.92
    return ((cs + 0.055) / 1.055) ** 2.4


def relative_luminance(c: ColorLike) -> float:
    """L по WCAG, диапазон 0.0 (чёрный) .. 1.0 (белый)."""
    r, g, b = parse_color(c)
    return 0.2126 * _linearize(r) + 0.7152 * _linearize(g) + 0.0722 * _linearize(b)


def contrast_ratio(c1: ColorLike, c2: ColorLike) -> float:
    """CR в диапазоне 1.0 .. 21.0. Порядок аргументов не важен."""
    l1 = relative_luminance(c1)
    l2 = relative_luminance(c2)
    lighter, darker = (l1, l2) if l1 >= l2 else (l2, l1)
    return (lighter + 0.05) / (darker + 0.05)


def is_dark(c: ColorLike, threshold: float = 0.5) -> bool:
    """True, если цвет «тёмный» (по воспринимаемой светлоте L*).

    Используем L* из CIELab (perceptual), а не сырую luminance — она лучше
    бьётся с человеческим ощущением «тёмный/светлый» на средних тонах.
    """
    return perceived_lightness(c) < threshold


def perceived_lightness(c: ColorLike) -> float:
    """L* (CIELab) нормированный в 0..1. 0 — чёрный, 1 — белый."""
    y = relative_luminance(c)
    if y <= 0.008856:
        lstar = y * 903.3
    else:
        lstar = 116.0 * (y ** (1.0 / 3.0)) - 16.0
    return lstar / 100.0


# --------------------------------------------------------------------------- #
# Главное: подбор читаемого цвета текста
# --------------------------------------------------------------------------- #

def best_text_color(
    bg: ColorLike,
    candidates: Iterable[ColorLike] = (BLACK, WHITE),
) -> RGB:
    """Из кандидатов выбирает тот, у кого максимальный контраст с фоном.

    По умолчанию кандидаты — чистый чёрный и чистый белый, поэтому:
        белый фон  -> чёрный текст
        чёрный фон -> белый текст
        синяя плашка -> что контрастнее (обычно белый)
    Можно передать «фирменные» тёмный/светлый из палитры, чтобы текст
    был не #000/#FFF, а мягче, но всё ещё читаемый.
    """
    best, best_cr = None, -1.0
    for cand in candidates:
        cr = contrast_ratio(cand, bg)
        if cr > best_cr:
            best, best_cr = parse_color(cand), cr
    return best  # type: ignore[return-value]


def ensure_contrast(
    fg: ColorLike,
    bg: ColorLike,
    min_ratio: float = AA_NORMAL,
    candidates: Iterable[ColorLike] = (BLACK, WHITE),
) -> RGB:
    """Гарантирует контраст >= min_ratio.

    Сначала пытаемся сохранить желаемый цвет fg (бренд/акцент), лишь
    подкручивая его светлоту в сторону чёрного или белого — так бренд-цвет
    «выживает» по тону. Если дотянуть не удалось — отдаём чистый best_text_color.
    """
    fg_rgb = parse_color(fg)
    if contrast_ratio(fg_rgb, bg) >= min_ratio:
        return fg_rgb

    # Куда тянуть: к белому, если фон тёмный; к чёрному, если светлый.
    target = WHITE if is_dark(bg) else BLACK
    # 20 шагов смешивания fg -> target
    for i in range(1, 21):
        t = i / 20.0
        mixed = _mix(fg_rgb, target, t)
        if contrast_ratio(mixed, bg) >= min_ratio:
            return mixed

    # не дотянули тоном — берём гарантированный максимум
    return best_text_color(bg, candidates)


def min_ratio_for_font(size_pt: float, bold: bool = False) -> float:
    """Порог WCAG в зависимости от размера шрифта.

    «Крупный текст» = >=18pt, либо >=14pt bold -> достаточно 3.0,
    иначе нужен 4.5.
    """
    large = size_pt >= 18 or (bold and size_pt >= 14)
    return AA_LARGE if large else AA_NORMAL


# --------------------------------------------------------------------------- #
# Градиенты: текст должен читаться над ВСЕМ градиентом, а не над одной точкой
# --------------------------------------------------------------------------- #

def best_text_color_over_gradient(
    stops: Sequence[ColorLike],
    candidates: Iterable[ColorLike] = (BLACK, WHITE),
) -> RGB:
    """Выбирает текст, у которого МАКСИМАЛЕН минимальный контраст по всем
    стопам градиента (maximin). То есть текст, который нигде не «провалится».
    """
    cand_list = [parse_color(c) for c in candidates]
    best, best_worst = None, -1.0
    for cand in cand_list:
        worst = min(contrast_ratio(cand, s) for s in stops)
        if worst > best_worst:
            best, best_worst = cand, worst
    return best  # type: ignore[return-value]


def worst_contrast_over_gradient(fg: ColorLike, stops: Sequence[ColorLike]) -> float:
    return min(contrast_ratio(fg, s) for s in stops)


# --------------------------------------------------------------------------- #
# Полупрозрачные слои (карточка с alpha над градиентом)
# --------------------------------------------------------------------------- #

def composite_over(fg: ColorLike, bg: ColorLike, alpha: float) -> RGB:
    """Цвет полупрозрачного fg (alpha 0..1) поверх непрозрачного bg.

    Нужен, когда KPI-карточка/плашка кладётся с прозрачностью на градиент:
    реальный фон под текстом = композит, по нему и считаем контраст.
    """
    fr, fg_, fb = parse_color(fg)
    br, bg_, bb = parse_color(bg)
    a = max(0.0, min(1.0, alpha))
    return (
        _clamp8(fr * a + br * (1 - a)),
        _clamp8(fg_ * a + bg_ * (1 - a)),
        _clamp8(fb * a + bb * (1 - a)),
    )


# --------------------------------------------------------------------------- #
# Фигура на фоне: подобрать и заливку фигуры, и текст на ней
# --------------------------------------------------------------------------- #

def pick_shape_fill_and_text(
    slide_bg: ColorLike,
    accent: ColorLike | None = None,
    min_fill_vs_bg: float = UI_COMPONENT,  # фигура должна отделяться от фона (>=3:1)
    min_text_vs_fill: float = AA_NORMAL,
) -> Tuple[RGB, RGB]:
    """Возвращает (fill, text) для фигуры (KPI-карточка, плашка, бейдж).

    Логика «микромоментов» из запроса:
      - на белом фоне чёрная фигура -> белый текст;
      - заливка фигуры обязана отделяться от фона (>=3:1), иначе фигура
        сливается со слайдом;
      - текст на фигуре считается уже от ИТОГОВОЙ заливки фигуры.
    """
    bg = parse_color(slide_bg)

    # 1. Кандидаты заливки: акцент (если дан и отделяется), иначе контрастный
    #    «противоположный» тон фона.
    if accent is not None and contrast_ratio(accent, bg) >= min_fill_vs_bg:
        fill = parse_color(accent)
    else:
        # фон светлый -> тёмная фигура, и наоборот
        fill = BLACK if not is_dark(bg) else WHITE
        # если бренд-акцент есть, но не дотягивает — притянем его к нужному тону
        if accent is not None:
            fill = ensure_contrast(accent, bg, min_fill_vs_bg)

    # 2. Текст на фигуре — от итоговой заливки
    text = ensure_contrast(best_text_color(fill), fill, min_text_vs_fill)
    return fill, text


# --------------------------------------------------------------------------- #
# Утилиты
# --------------------------------------------------------------------------- #

def _mix(c1: ColorLike, c2: ColorLike, t: float) -> RGB:
    """Линейная интерполяция c1 -> c2, t in [0,1] (в sRGB, для простоты)."""
    r1, g1, b1 = parse_color(c1)
    r2, g2, b2 = parse_color(c2)
    return (
        _clamp8(r1 + (r2 - r1) * t),
        _clamp8(g1 + (g2 - g1) * t),
        _clamp8(b1 + (b2 - b1) * t),
    )


def lighten(c: ColorLike, t: float) -> RGB:
    return _mix(c, WHITE, t)


def darken(c: ColorLike, t: float) -> RGB:
    return _mix(c, BLACK, t)
