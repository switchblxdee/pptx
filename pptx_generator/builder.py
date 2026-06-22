"""
Сборка дайджеста-презентации.

Реализует визуальный язык референса «Голос IT»:
- Cover-слайд: градиентный фон, крупный заголовок, pill-метки источников,
  4 белые KPI-карточки с лёгкой тенью, шапка с метаданными.
- Topic-слайды: градиентный фон, плашка с названием темы, 4 KPI-карточки
  в линию (с обводкой), список тем-проблем на пастельной карточке,
  бейджи `new`, pill-метки источников, футер с метаданными.

Технические решения:
- Градиент фона делается через прямой XML-патч background-fill
  (python-pptx не предоставляет высокоуровневого API для градиентов).
- Pill-метки — rounded rectangle с динамической шириной от длины текста.
- Все размеры — в Emu/Inches, единая константа LAYOUT-зоны.
"""
from __future__ import annotations

from pathlib import Path
from typing import List, Optional, Tuple

from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

from . import contrast as C
from .icons import draw_icon, has_icon
from .brand_icons import draw_brand_icon, has_brand_icon
from .safe import with_fallback
from .schemas import (
    AttentionSlide,
    ChartSlide,
    ClosingSlide,
    CoverSlide,
    DigestSpec,
    DigestStyle,
    ExecutiveSummarySlide,
    KPICard,
    PatternsSlide,
    TopicItem,
    TopicSlide,
)

# --------------------------------------------------------------------------- #
# Константы
# --------------------------------------------------------------------------- #

SLIDE_WIDTH = Inches(13.333)
SLIDE_HEIGHT = Inches(7.5)

# Зона безопасных отступов
MARGIN_X = Inches(0.55)

# Зона шапки/футера с метаданными
META_HEADER_TOP = Inches(0.25)
META_HEADER_HEIGHT = Inches(0.4)
FOOTER_BOTTOM = Inches(0.2)
FOOTER_HEIGHT = Inches(0.35)

# XML namespaces для прямого патчинга OOXML
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"


# --------------------------------------------------------------------------- #
# Builder
# --------------------------------------------------------------------------- #

class DigestBuilder:
    """Собирает .pptx-дайджест по DigestSpec."""

    def __init__(self, spec: DigestSpec):
        self.spec = spec
        self.style = spec.style
        # Роли, заданные пользователем явно — их авто-контраст не трогает.
        self.locked = set(getattr(spec.style, "locked_roles", None) or [])
        self.palette = self._heal_palette(spec.style.palette)
        # Брендовые фоновые картинки (например шаблон SberF1)
        self.bg_cover = self._resolve_asset(getattr(spec.style, "background_cover", None))
        self.bg_content = self._resolve_asset(getattr(spec.style, "background_content", None))
        # Брендовые иконки SberF1 включаем вместе с брендовым фоном
        self.use_brand_icons = bool(self.bg_content or self.bg_cover)
        self.prs = Presentation()
        self.prs.slide_width = SLIDE_WIDTH
        self.prs.slide_height = SLIDE_HEIGHT

    def _locked(self, *roles: str) -> bool:
        """True, если любая из ролей зафиксирована пользователем."""
        return any(r in self.locked for r in roles)

    # Авто-подбор иконки по тексту подписи KPI — чтобы значки появлялись
    # даже если LLM не проставила icon_hint (актуально для GLM и подобных).
    _AUTO_ICON = [
        (("сигнал", "обращен", "упомин", "запрос"), "signal"),
        (("безопас", "кибербез", "защит"), "security"),
        (("риск", "инцидент", "угроз"), "warning"),
        (("нов",), "bell"),
        (("тем", "рост", "динамик", "актив", "увелич"), "growth"),
        (("источник", "команд", "сотрудн", "польз", "людей"), "team"),
        (("продукт", "инструмент", "сервис", "систем"), "tools"),
        (("период", "срок", "недел", "время", "дата"), "clock"),
        (("баг", "ошибк", "дефект", "сбой"), "bug"),
        (("поиск", "релевант", "наход"), "search"),
        (("деньг", "финанс", "затрат", "бюджет", "стоим"), "money"),
        (("почт", "письм", "обратн"), "email"),
    ]

    def _auto_icon_hint(self, label: Optional[str]) -> Optional[str]:
        if not label:
            return "signal"
        t = label.lower()
        for keys, hint in self._AUTO_ICON:
            if any(k in t for k in keys):
                return hint
        return "signal"  # нейтральный дефолт (гистограмма)

    def _has_any_icon(self, hint: Optional[str]) -> bool:
        if not hint:
            return False
        return (self.use_brand_icons and has_brand_icon(hint)) or has_icon(hint)

    def _draw_icon(self, slide, hint, left, top, size, color) -> None:
        """Рисует иконку: сначала брендовую (SberF1), иначе встроенную."""
        if self.use_brand_icons and draw_brand_icon(slide, hint, left, top, size, color):
            return
        if has_icon(hint):
            draw_icon(slide, hint, left, top, size, color)

    @staticmethod
    def _resolve_asset(name: Optional[str]) -> Optional[str]:
        """Путь к упакованному ассету (assets/<name>) или None."""
        if not name:
            return None
        p = Path(__file__).resolve().parent / "assets" / name
        return str(p) if p.exists() else None

    # ----------------------------------------------------------------------- #
    # Самолечение палитры: гарантируем читаемость основного/приглушённого
    # текста на поверхностях карточек. Для согласованных пресетов — no-op;
    # срабатывает только на «сломанной» кастомной палитре (например от LLM).
    # ----------------------------------------------------------------------- #

    def _heal_palette(self, pal):
        surfaces = [pal.card_bg, pal.kpi_bg]
        # Зафиксированные пользователем роли не лечим — рендерим как задано.
        new_dark = (C.to_hex(pal.text_dark, with_hash=False) if self._locked("text")
                    else self._heal_ink(pal.text_dark, surfaces, C.AA_NORMAL))
        new_muted = (C.to_hex(pal.text_muted, with_hash=False)
                     if self._locked("text", "text_muted")
                     else self._heal_ink(pal.text_muted, surfaces, C.AA_LARGE))
        try:
            # Pydantic v2 — не мутируем исходную спеку, делаем копию
            return pal.model_copy(update={"text_dark": new_dark,
                                          "text_muted": new_muted})
        except AttributeError:
            # dataclass-mock из smoke_test — правим поля на месте
            pal.text_dark = new_dark
            pal.text_muted = new_muted
            return pal

    @staticmethod
    def _heal_ink(color: str, surfaces: list[str], min_ratio: float) -> str:
        """Если color читается на всех surfaces — оставляем; иначе берём
        цвет с максимальным минимальным контрастом по поверхностям."""
        worst = min(C.contrast_ratio(color, s) for s in surfaces)
        if worst >= min_ratio:
            return C.to_hex(color, with_hash=False)
        rgb = C.best_text_color_over_gradient(
            surfaces, candidates=[color, "0A0A0A", "FFFFFF"])
        return C.to_hex(rgb, with_hash=False)

    def build(self, output_path: str | Path) -> Path:
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        # Считаем общее число слайдов для нумерации в футере
        total = 1  # cover
        if self.spec.executive_summary:
            total += 1
        total += len(self.spec.topics)
        total += len(self.spec.charts)
        if self.spec.patterns:
            total += 1
        if self.spec.attention:
            total += 1
        if self.spec.closing:
            total += 1

        page = 1

        # 1. Обложка
        cover_slide = self._new_slide(cover=True)
        self._render_cover(cover_slide, self.spec.cover)
        self._render_meta_header(cover_slide)

        # 2. Executive summary (сразу после обложки)
        if self.spec.executive_summary:
            page += 1
            slide = self._new_slide()
            self._render_executive_summary(slide, self.spec.executive_summary)
            self._render_meta_footer(slide, page_number=page, total=total)

        # 3. Topic-слайды
        for topic in self.spec.topics:
            page += 1
            slide = self._new_slide()
            self._render_topic(slide, topic)
            self._render_meta_footer(slide, page_number=page, total=total)

        # 4. Слайды-графики (после тем)
        for chart in self.spec.charts:
            page += 1
            slide = self._new_slide()
            self._render_chart(slide, chart)
            self._render_meta_footer(slide, page_number=page, total=total)

        # 5. Сквозные паттерны
        if self.spec.patterns:
            page += 1
            slide = self._new_slide()
            self._render_patterns(slide, self.spec.patterns)
            self._render_meta_footer(slide, page_number=page, total=total)

        # 5. На что обратить внимание
        if self.spec.attention:
            page += 1
            slide = self._new_slide()
            self._render_attention(slide, self.spec.attention)
            self._render_meta_footer(slide, page_number=page, total=total)

        # 6. Итоги
        if self.spec.closing:
            page += 1
            slide = self._new_slide()
            self._render_closing(slide, self.spec.closing)
            self._render_meta_footer(slide, page_number=page, total=total)

        self._finalize_unique_ids()
        self.prs.save(output_path)
        return output_path

    def _finalize_unique_ids(self) -> None:
        """Делает все id фигур уникальными внутри каждого слайда.

        Иконки из шаблона и автогенерация python-pptx могут давать
        повторяющиеся cNvPr@id. PowerPoint/LibreOffice это терпят, но
        R7-Офис/OnlyOffice из-за дублей теряет/путает фигуры. Финальный
        проход переназначает id последовательно — гарантия уникальности.
        """
        for slide in self.prs.slides:
            nid = 1
            for cnv in slide._element.iter(qn("p:cNvPr")):
                nid += 1
                cnv.set("id", str(nid))

    def _new_slide(self, cover: bool = False):
        """Создаёт пустой слайд с фоном.

        Если в стиле заданы брендовые фоновые картинки (например шаблон
        SberF1) — кладём картинку full-bleed (cover-фон на обложку, content
        на остальные). Иначе — градиент (премиум → простой fallback).
        Презентация не падает ни при каких условиях.
        """
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        bg_img = self.bg_cover if cover else self.bg_content
        if bg_img:
            with_fallback(
                lambda: self._set_image_background(slide, bg_img),
                lambda: self._set_gradient_background(slide),
                label="bg_image",
            )
        else:
            with_fallback(
                lambda: self._set_premium_background(slide),
                lambda: self._set_gradient_background(slide),
                label="background",
            )
        return slide

    def _set_image_background(self, slide, image_path: str) -> None:
        """Кладёт картинку как фон слайда (full-bleed). Картинка добавляется
        первой, поэтому оказывается позади всего контента."""
        slide.shapes.add_picture(image_path, 0, 0, SLIDE_WIDTH, SLIDE_HEIGHT)

    # ----------------------------------------------------------------------- #
    # COVER
    # ----------------------------------------------------------------------- #

    def _render_cover(self, slide, cover: CoverSlide) -> None:
        # Цвета текста, лежащего прямо на градиенте — контрастные к фону
        on_bg = self._text_on_background()
        on_bg_muted = self._muted_on_background()

        # Заголовок («Голос IT»)
        self._add_text(
            slide, cover.title,
            left=Inches(0.8), top=Inches(2.0),
            width=Inches(11.5), height=Inches(0.9),
            font=self.style.typography.heading_font,
            size=44, bold=True, color=on_bg,
        )

        # Подзаголовок («дайджест для руководства Блока T»)
        self._add_text(
            slide, cover.subtitle,
            left=Inches(0.8), top=Inches(2.85),
            width=Inches(11.5), height=Inches(0.9),
            font=self.style.typography.heading_font,
            size=32, bold=True, color=on_bg,
        )

        # Описание (мелким текстом)
        if cover.description:
            self._add_text(
                slide, cover.description,
                left=Inches(0.8), top=Inches(3.85),
                width=Inches(11.5), height=Inches(0.7),
                font=self.style.typography.body_font,
                size=14, color=on_bg_muted,
            )

        # Pill-метки источников
        self._render_pills_row(
            slide, cover.source_tags,
            top=Inches(4.7),
            font_size=11,
            text_color=self.palette.text_dark,
            bg_color=self.palette.kpi_bg,
        )

        # KPI-карточки внизу
        self._render_kpi_row(
            slide, cover.kpis,
            top=Inches(5.5),
            height=Inches(1.6),
            value_size=36,
            label_size=11,
            value_color=self.palette.text_dark,
            label_color=self.palette.text_muted,
            card_bg=self.palette.kpi_bg,
            border_color=None,  # на cover карточки без обводки
        )

    # ----------------------------------------------------------------------- #
    # TOPIC
    # ----------------------------------------------------------------------- #

    def _render_topic(self, slide, topic: TopicSlide) -> None:
        # Зона заголовка темы слева (плашка как на референсе)
        title_box_left = Inches(0.6)
        title_box_top = Inches(0.95)
        title_box_width = Inches(4.2)
        title_box_height = Inches(1.5)

        # Адаптивный размер шрифта по длине названия темы
        if len(topic.title) > 30:
            title_size = 24
        elif len(topic.title) > 18:
            title_size = 28
        else:
            title_size = 32

        # Чтобы тема визуально «звучала» — лёгкий блок-плашка с акцентом
        self._add_text(
            slide, topic.title,
            left=title_box_left, top=title_box_top,
            width=title_box_width, height=title_box_height,
            font=self.style.typography.heading_font,
            size=title_size, bold=True, color=self._text_on_background(),
            anchor=MSO_ANCHOR.MIDDLE,
        )

        # KPI-карточки в линию справа от заголовка темы
        kpi_left_start = Inches(5.0)
        kpi_total_width = Inches(7.8)
        kpi_gap = Inches(0.2)
        n_kpi = len(topic.kpis)
        kpi_card_width = Emu(int((kpi_total_width - kpi_gap * (n_kpi - 1)) / n_kpi))

        on_bg = self._text_on_background()
        on_bg_muted = self._muted_on_background()
        for i, kpi in enumerate(topic.kpis):
            left = kpi_left_start + (kpi_card_width + kpi_gap) * i
            self._render_kpi_card(
                slide, kpi,
                left=left, top=Inches(0.95),
                width=kpi_card_width, height=Inches(1.5),
                value_size=28,
                label_size=10,
                value_color=on_bg,
                label_color=on_bg_muted,
                card_bg=None,                # прозрачный (на градиенте)
                border_color=on_bg,          # обводка контрастна фону
            )

        # Карточка-список тем (пастельный фон).
        # Высота АДАПТИВНАЯ: при 1-2 темах фиксированные 3.5" оставляли
        # полупустую карточку. Теперь ~1.15" на тему + внутренние отступы,
        # с минимумом 1.6" и максимумом 3.6".
        list_top = Inches(2.7)
        n_items = len(topic.items)
        raw_height = 0.6 + n_items * 1.05
        list_height = Inches(min(3.6, max(1.6, raw_height)))
        self._render_topic_items_card(
            slide, topic.items,
            left=MARGIN_X, top=list_top,
            width=SLIDE_WIDTH - MARGIN_X * 2,
            height=list_height,
        )

        # Pill-метки источников внизу
        if topic.source_tags:
            self._add_text(
                slide, "Источники:",
                left=MARGIN_X, top=Inches(6.4),
                width=Inches(1.2), height=Inches(0.3),
                font=self.style.typography.body_font,
                size=11, color=self._muted_on_background(),
                anchor=MSO_ANCHOR.MIDDLE,
            )
            self._render_pills_row(
                slide, topic.source_tags,
                top=Inches(6.4),
                start_left=Inches(1.65),
                font_size=10,
                text_color=self.palette.text_dark,
                bg_color=self.palette.kpi_bg,
            )

    def _render_topic_items_card(
        self, slide, items: List[TopicItem],
        left, top, width, height,
    ) -> None:
        """Карточка с пастельным фоном, содержащая список тем-проблем."""
        # Фон карточки
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.adjustments[0] = 0.04
        self._fill_solid(card, self.palette.card_bg)
        card.line.fill.background()
        # Лёгкая тень (без неё карточка плоская)
        self._apply_subtle_shadow(card)

        # Раскладка items внутри карточки
        n = len(items)
        inner_pad_x = Inches(0.5)
        inner_pad_y = Inches(0.3)

        # Высота каждой строки = (общая высота - отступы) / n
        usable_height = height - inner_pad_y * 2
        row_height = Emu(int(usable_height / n))

        # Колонки внутри строки:
        # [номер|заголовок+цитата .............. период | mentions]
        col_num_width = Inches(0.4)
        col_period_width = Inches(2.0)
        col_mentions_width = Inches(0.9)
        col_content_left = left + inner_pad_x + col_num_width + Inches(0.1)
        col_content_width = (
            width - inner_pad_x * 2 - col_num_width - col_period_width - col_mentions_width - Inches(0.3)
        )
        col_period_left = left + width - inner_pad_x - col_mentions_width - col_period_width - Inches(0.1)
        col_mentions_left = left + width - inner_pad_x - col_mentions_width

        # Максимум упоминаний — для нормировки мини-баров
        max_mentions = max((it.mentions for it in items), default=0) or 1

        for i, item in enumerate(items):
            row_top = top + inner_pad_y + row_height * i

            # Номер
            self._add_text(
                slide, f"{i + 1}.",
                left=left + inner_pad_x, top=row_top,
                width=col_num_width, height=Inches(0.4),
                font=self.style.typography.body_font,
                size=14, bold=True, on=self.palette.card_bg, role="strong",
                anchor=MSO_ANCHOR.TOP,
            )

            # Бейдж `new` (если есть) — рисуем ПОСЛЕ заголовка,
            # чтобы он не наезжал. Заголовок сжимаем по ширине.
            badge_width = Inches(0.5) if item.is_new else Inches(0)
            content_left = col_content_left
            # При наличии бейджа отступаем заголовок вправо на ширину бейджа
            if item.is_new:
                self._render_new_badge(
                    slide,
                    left=col_content_left,
                    top=row_top + Inches(0.05),
                )
                content_left = col_content_left + badge_width + Inches(0.1)
            content_width = col_content_width - (
                badge_width + Inches(0.1) if item.is_new else Emu(0)
            )

            # Заголовок темы
            self._add_text(
                slide, item.title,
                left=content_left, top=row_top,
                width=content_width, height=Inches(0.4),
                font=self.style.typography.heading_font,
                size=15, bold=True, on=self.palette.card_bg, role="strong",
            )

            # Цитата под заголовком (курсив) — начинается с того же left,
            # что и заголовок (если бейдж был — после него)
            if item.quote:
                self._add_text(
                    slide, f"«{item.quote}»",
                    left=col_content_left, top=row_top + Inches(0.4),
                    width=col_content_width, height=Inches(0.6),
                    font=self.style.typography.body_font,
                    size=11, italic=True, on=self.palette.card_bg, role="muted",
                    line_spacing=1.2,
                )

            # Период (справа)
            self._add_text(
                slide, item.period,
                left=col_period_left, top=row_top + Inches(0.05),
                width=col_period_width, height=Inches(0.4),
                font=self.style.typography.body_font,
                size=12, on=self.palette.card_bg, role="strong",
                align=PP_ALIGN.RIGHT,
            )

            # Mentions (число + подпись)
            self._add_text(
                slide, str(item.mentions),
                left=col_mentions_left, top=row_top,
                width=col_mentions_width, height=Inches(0.4),
                font=self.style.typography.heading_font,
                size=20, bold=True, on=self.palette.card_bg, role="strong",
                align=PP_ALIGN.RIGHT,
            )
            self._add_text(
                slide, "упом.",
                left=col_mentions_left, top=row_top + Inches(0.4),
                width=col_mentions_width, height=Inches(0.3),
                font=self.style.typography.body_font,
                size=10, on=self.palette.card_bg, role="muted",
                align=PP_ALIGN.RIGHT,
            )

            # Мини-бар под числом упоминаний — визуализация относительного веса.
            # Длина пропорциональна mentions / max_mentions.
            bar_full_width = col_mentions_width
            bar_ratio = item.mentions / max_mentions
            bar_width = Emu(max(int(bar_full_width * bar_ratio), Inches(0.1)))
            bar_top = row_top + Inches(0.72)
            # Фоновая дорожка (бледная)
            track = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                col_mentions_left, bar_top, bar_full_width, Inches(0.07),
            )
            track.adjustments[0] = 0.5
            self._fill_solid(track, self.palette.text_muted)
            track.line.fill.background()
            track.fill.fore_color.rgb = self._rgb(self.palette.text_muted)
            # Заполнение (акцентный цвет), выровнено по правому краю
            fill_bar = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                col_mentions_left + bar_full_width - bar_width, bar_top,
                bar_width, Inches(0.07),
            )
            fill_bar.adjustments[0] = 0.5
            self._fill_solid(fill_bar, self.palette.accent)
            fill_bar.line.fill.background()

            # Разделитель между строками (кроме последней)
            if i < n - 1:
                sep_y = row_top + row_height - Inches(0.05)
                separator = slide.shapes.add_connector(
                    1,  # MSO_CONNECTOR.STRAIGHT
                    left + inner_pad_x, sep_y,
                    left + width - inner_pad_x, sep_y,
                )
                separator.line.color.rgb = self._rgb(self.palette.text_muted)
                separator.line.width = Pt(0.5)

    def _render_new_badge(self, slide, left, top) -> None:
        """Оранжевый бейдж `new` как на референсе."""
        badge = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            left, top,
            Inches(0.4), Inches(0.22),
        )
        badge.adjustments[0] = 0.5
        self._fill_solid(badge, self.palette.badge)
        badge.line.fill.background()
        tf = badge.text_frame
        tf.margin_left = Emu(0); tf.margin_right = Emu(0)
        tf.margin_top = Emu(0); tf.margin_bottom = Emu(0)
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.text = "new"
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        badge_ink = self._ink_on(self.palette.badge, prefer="FFFFFF",
                                 size_pt=9, bold=True)
        for run in p.runs:
            run.font.name = self.style.typography.body_font
            run.font.size = Pt(9)
            run.font.bold = True
            run.font.color.rgb = self._rgb(badge_ink)

    # ----------------------------------------------------------------------- #
    # АНАЛИТИЧЕСКИЕ СЛАЙДЫ
    # ----------------------------------------------------------------------- #

    def _render_slide_title(self, slide, title: str, accent_bar: bool = True) -> None:
        """Общий заголовок аналитического слайда с акцентным маркером."""
        if accent_bar:
            bar = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                Inches(0.6), Inches(0.7),
                Inches(0.1), Inches(0.6),
            )
            self._fill_solid(bar, self.palette.accent)
            bar.line.fill.background()

        self._add_text(
            slide, title,
            left=Inches(0.85), top=Inches(0.6),
            width=Inches(11.5), height=Inches(0.8),
            font=self.style.typography.heading_font,
            size=32, bold=True, color=self._text_on_background(),
            anchor=MSO_ANCHOR.MIDDLE,
        )

    def _render_executive_summary(self, slide, s: ExecutiveSummarySlide) -> None:
        """Executive summary: вводный абзац + крупные тезисы."""
        self._render_slide_title(slide, s.title)

        top = Inches(1.7)

        # Вводный абзац на пастельной плашке
        if s.intro:
            intro_card = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                MARGIN_X, top, SLIDE_WIDTH - MARGIN_X * 2, Inches(0.9),
            )
            intro_card.adjustments[0] = 0.06
            self._fill_solid(intro_card, self.palette.card_bg)
            intro_card.line.fill.background()
            self._apply_subtle_shadow(intro_card)

            self._add_text(
                slide, s.intro,
                left=MARGIN_X + Inches(0.35), top=top + Inches(0.1),
                width=SLIDE_WIDTH - MARGIN_X * 2 - Inches(0.7), height=Inches(0.7),
                font=self.style.typography.body_font,
                size=14, on=self.palette.card_bg, role="strong",
                anchor=MSO_ANCHOR.MIDDLE,
            )
            top = top + Inches(1.2)

        # Тезисы: каждый — headline крупно + detail мельче
        n = len(s.points)
        usable_height = SLIDE_HEIGHT - top - Inches(0.7)
        row_height = Emu(int(usable_height / n))

        for i, point in enumerate(s.points):
            row_top = top + row_height * i

            # Номерной акцент-кружок
            num_circle = slide.shapes.add_shape(
                MSO_SHAPE.OVAL,
                MARGIN_X, row_top + Inches(0.05),
                Inches(0.45), Inches(0.45),
            )
            self._fill_solid(num_circle, self.palette.accent)
            num_circle.line.fill.background()
            tf = num_circle.text_frame
            tf.margin_left = Emu(0); tf.margin_right = Emu(0)
            tf.margin_top = Emu(0); tf.margin_bottom = Emu(0)
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf.text = str(i + 1)
            pp = tf.paragraphs[0]
            pp.alignment = PP_ALIGN.CENTER
            for run in pp.runs:
                run.font.name = self.style.typography.heading_font
                run.font.size = Pt(18)
                run.font.bold = True
                run.font.color.rgb = self._rgb(self._ink_on(
                    self.palette.accent, prefer=self.palette.kpi_bg,
                    size_pt=18, bold=True))

            content_left = MARGIN_X + Inches(0.7)
            content_width = SLIDE_WIDTH - content_left - MARGIN_X

            # Headline
            self._add_text(
                slide, point.headline,
                left=content_left, top=row_top,
                width=content_width, height=Inches(0.45),
                font=self.style.typography.heading_font,
                size=18, bold=True, color=self._text_on_background(),
            )
            # Detail
            if point.detail:
                self._add_text(
                    slide, point.detail,
                    left=content_left, top=row_top + Inches(0.45),
                    width=content_width, height=Inches(0.5),
                    font=self.style.typography.body_font,
                    size=13, color=self._muted_on_background(),
                )

    def _render_patterns(self, slide, s: PatternsSlide) -> None:
        """Сквозные паттерны: карточки-блоки с описанием закономерностей."""
        self._render_slide_title(slide, s.title)

        top = Inches(1.7)
        if s.intro:
            self._add_text(
                slide, s.intro,
                left=MARGIN_X, top=top,
                width=SLIDE_WIDTH - MARGIN_X * 2, height=Inches(0.6),
                font=self.style.typography.body_font,
                size=14, color=self._muted_on_background(),
            )
            top = top + Inches(0.75)

        n = len(s.patterns)
        usable_height = SLIDE_HEIGHT - top - Inches(0.7)
        gap = Inches(0.2)
        card_height = Emu(int((usable_height - gap * (n - 1)) / n))

        for i, pattern in enumerate(s.patterns):
            card_top = top + (card_height + gap) * i
            card = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                MARGIN_X, card_top,
                SLIDE_WIDTH - MARGIN_X * 2, card_height,
            )
            card.adjustments[0] = 0.08
            self._fill_solid(card, self.palette.card_bg)
            card.line.fill.background()
            self._apply_subtle_shadow(card)

            inner_x = MARGIN_X + Inches(0.4)

            # Заголовок паттерна
            title_width = SLIDE_WIDTH - MARGIN_X * 2 - Inches(0.8) - Inches(1.6)
            self._add_text(
                slide, pattern.title,
                left=inner_x, top=card_top + Inches(0.15),
                width=title_width, height=Inches(0.4),
                font=self.style.typography.heading_font,
                size=16, bold=True, on=self.palette.card_bg, role="strong",
            )
            # Описание
            self._add_text(
                slide, pattern.description,
                left=inner_x, top=card_top + Inches(0.55),
                width=title_width, height=card_height - Inches(0.6),
                font=self.style.typography.body_font,
                size=12, on=self.palette.card_bg, role="muted",
            )

            # Бейдж «затронуто N тем» справа
            if pattern.affected_count is not None:
                badge_left = SLIDE_WIDTH - MARGIN_X - Inches(1.7)
                self._add_text(
                    slide, str(pattern.affected_count),
                    left=badge_left, top=card_top + Inches(0.15),
                    width=Inches(1.4), height=Inches(0.5),
                    font=self.style.typography.heading_font,
                    size=28, bold=True, color=self._accent_ink(self.palette.card_bg),
                    align=PP_ALIGN.RIGHT,
                )
                self._add_text(
                    slide, "тем затронуто",
                    left=badge_left, top=card_top + Inches(0.65),
                    width=Inches(1.4), height=Inches(0.3),
                    font=self.style.typography.body_font,
                    size=10, on=self.palette.card_bg, role="muted",
                    align=PP_ALIGN.RIGHT,
                )

    def _render_attention(self, slide, s: AttentionSlide) -> None:
        """На что обратить внимание: список с цветовыми маркерами важности."""
        self._render_slide_title(slide, s.title)

        # Цвета severity
        severity_colors = {
            "высокий": self.palette.badge,
            "средний": self.palette.accent,
            "низкий": self.palette.text_muted,
        }

        top = Inches(1.8)
        n = len(s.items)
        gap = Inches(0.2)
        usable_height = SLIDE_HEIGHT - top - Inches(0.7)
        card_height = Emu(int((usable_height - gap * (n - 1)) / n))

        for i, item in enumerate(s.items):
            card_top = top + (card_height + gap) * i
            color = severity_colors.get(item.severity, self.palette.accent)

            # Карточка
            card = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                MARGIN_X, card_top,
                SLIDE_WIDTH - MARGIN_X * 2, card_height,
            )
            card.adjustments[0] = 0.08
            self._fill_solid(card, self.palette.kpi_bg)
            card.line.fill.background()
            self._apply_subtle_shadow(card)

            # Цветная полоса severity слева
            sev_bar = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                MARGIN_X + Inches(0.15), card_top + Inches(0.15),
                Inches(0.12), card_height - Inches(0.3),
            )
            sev_bar.adjustments[0] = 0.5
            self._fill_solid(sev_bar, color)
            sev_bar.line.fill.background()

            inner_x = MARGIN_X + Inches(0.5)
            content_width = SLIDE_WIDTH - inner_x - MARGIN_X - Inches(1.5)

            # Заголовок
            self._add_text(
                slide, item.title,
                left=inner_x, top=card_top + Inches(0.15),
                width=content_width, height=Inches(0.4),
                font=self.style.typography.heading_font,
                size=15, bold=True, on=self.palette.kpi_bg, role="strong",
            )
            # Обоснование
            self._add_text(
                slide, item.rationale,
                left=inner_x, top=card_top + Inches(0.55),
                width=content_width, height=card_height - Inches(0.6),
                font=self.style.typography.body_font,
                size=12, on=self.palette.kpi_bg, role="muted",
            )

            # Бейдж severity справа
            sev_badge_left = SLIDE_WIDTH - MARGIN_X - Inches(1.3)
            badge = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                sev_badge_left, card_top + (card_height - Inches(0.35)) / 2,
                Inches(1.1), Inches(0.35),
            )
            badge.adjustments[0] = 0.5
            self._fill_solid(badge, color)
            badge.line.fill.background()
            tf = badge.text_frame
            tf.margin_left = Emu(0); tf.margin_right = Emu(0)
            tf.margin_top = Emu(0); tf.margin_bottom = Emu(0)
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf.text = item.severity
            pp = tf.paragraphs[0]
            pp.alignment = PP_ALIGN.CENTER
            for run in pp.runs:
                run.font.name = self.style.typography.body_font
                run.font.size = Pt(11)
                run.font.bold = True
                run.font.color.rgb = self._rgb(self._ink_on(
                    color, prefer=self.palette.kpi_bg, size_pt=11, bold=True))

    def _render_closing(self, slide, s: ClosingSlide) -> None:
        """Финальный слайд: обобщающий текст + опциональные итоговые KPI."""
        self._render_slide_title(slide, s.title)

        # Обобщающий текст на крупной плашке
        summary_card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            MARGIN_X, Inches(1.8),
            SLIDE_WIDTH - MARGIN_X * 2, Inches(2.2),
        )
        summary_card.adjustments[0] = 0.05
        self._fill_solid(summary_card, self.palette.card_bg)
        summary_card.line.fill.background()
        self._apply_subtle_shadow(summary_card)

        self._add_text(
            slide, s.summary,
            left=MARGIN_X + Inches(0.5), top=Inches(2.0),
            width=SLIDE_WIDTH - MARGIN_X * 2 - Inches(1.0), height=Inches(1.8),
            font=self.style.typography.body_font,
            size=16, on=self.palette.card_bg, role="strong",
            anchor=MSO_ANCHOR.MIDDLE,
            line_spacing=1.3,
        )

        # Итоговые KPI, если есть
        if s.kpis:
            self._render_kpi_row(
                slide, s.kpis,
                top=Inches(4.6),
                height=Inches(1.6),
                value_size=40, label_size=12,
                value_color=self.palette.text_dark,
                label_color=self.palette.text_muted,
                card_bg=self.palette.kpi_bg,
                border_color=None,
            )

    # ----------------------------------------------------------------------- #
    # СЛАЙД-ГРАФИК (нативный pptx-чарт)
    # ----------------------------------------------------------------------- #

    # Маппинг наших типов в pptx-типы
    _CHART_TYPE_MAP = {
        "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "bar": XL_CHART_TYPE.BAR_CLUSTERED,
        "line": XL_CHART_TYPE.LINE_MARKERS,
        "pie": XL_CHART_TYPE.PIE,
    }

    # Рисовать столбчатые графики фигурами (прямоугольниками), а не нативным
    # чарт-объектом. Фигуры рендерятся одинаково везде (PowerPoint/LibreOffice/
    # Google Slides), поэтому «столбцы не появились» становится невозможным.
    DRAW_CHARTS_AS_SHAPES = True

    def _render_chart(self, slide, s: ChartSlide) -> None:
        """Слайд с графиком.

        Для column/bar по умолчанию рисуем гистограмму ФИГУРАМИ (надёжно во
        всех редакторах). Для pie/line — нативный pptx-чарт. Нативный путь
        остаётся fallback'ом, если отрисовка фигурами вдруг упадёт.
        """
        self._render_slide_title(slide, s.title)

        chart_top = Inches(1.7)
        if s.subtitle:
            self._add_text(
                slide, s.subtitle,
                left=MARGIN_X, top=Inches(1.55),
                width=SLIDE_WIDTH - MARGIN_X * 2, height=Inches(0.5),
                font=self.style.typography.body_font,
                size=14, color=self._muted_on_background(),
            )
            chart_top = Inches(2.15)

        chart_height = Inches(3.9) if s.insight else Inches(4.5)
        chart_width = SLIDE_WIDTH - MARGIN_X * 2

        if self.DRAW_CHARTS_AS_SHAPES and s.chart_type in ("column", "bar"):
            with_fallback(
                lambda: self._render_bar_chart_shapes(
                    slide, s, MARGIN_X, chart_top, chart_width, chart_height),
                lambda: self._render_native_chart(
                    slide, s, chart_top, chart_height, chart_width),
                label="bar_chart_shapes",
            )
        else:
            self._render_native_chart(slide, s, chart_top, chart_height, chart_width)

        if s.insight:
            insight_top = chart_top + chart_height + Inches(0.1)
            card = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                MARGIN_X, insight_top,
                SLIDE_WIDTH - MARGIN_X * 2, Inches(0.7),
            )
            card.adjustments[0] = 0.12
            self._fill_solid(card, self.palette.card_bg)
            card.line.fill.background()
            self._apply_subtle_shadow(card)
            self._add_text(
                slide, s.insight,
                left=MARGIN_X + Inches(0.35), top=insight_top + Inches(0.05),
                width=SLIDE_WIDTH - MARGIN_X * 2 - Inches(0.7), height=Inches(0.6),
                font=self.style.typography.body_font,
                size=13, color=self._ink(self.palette.card_bg, role="strong", size_pt=13),
                anchor=MSO_ANCHOR.MIDDLE,
            )

    @staticmethod
    def _nice_ceil(v: float) -> float:
        """Округляет максимум вверх до «красивого» числа для оси."""
        import math
        if v <= 0:
            return 1.0
        mag = 10 ** math.floor(math.log10(v))
        for m in (1, 2, 2.5, 5, 10):
            if v <= m * mag:
                return m * mag
        return 10 * mag

    def _render_bar_chart_shapes(
        self, slide, s: ChartSlide, left, top, width, height,
    ) -> None:
        """Гистограмма, нарисованная прямоугольниками (column или bar).

        Полностью независима от чарт-движка редактора: оси — линии-коннекторы,
        столбцы — autoshapes, подписи — текстбоксы. Цвет столбцов контрастен
        фону (никогда не «исчезает»), у каждого столбца есть контур.
        """
        horizontal = (s.chart_type == "bar")
        values = list(s.values)
        cats = list(s.categories)
        n = len(values)
        if n == 0:
            return

        vmax = max(values + [0])
        nice_max = self._nice_ceil(vmax)

        ink = self._text_on_background()
        grid = self._muted_on_background()
        bar_fill = (self.palette.accent if self._locked("accent")
                    else self._contrasting_fill(self.palette.accent, 3.0))
        bar_outline = self._contrasting_fill(bar_fill, 1.6)

        L, T, W, H = int(left), int(top), int(width), int(height)
        y_gutter = int(Inches(0.55))   # под подписи оси значений / категорий (для bar)
        x_gutter = int(Inches(0.45))   # под подписи категорий / значений

        if not horizontal:
            # --- ВЕРТИКАЛЬНЫЕ СТОЛБЦЫ (column) ---
            plot_left = L + y_gutter
            plot_top = T
            plot_w = W - y_gutter
            plot_h = H - x_gutter
            baseline = plot_top + plot_h

            self._chart_axis_line(slide, plot_left, baseline, plot_left + plot_w, baseline, ink)
            self._chart_axis_line(slide, plot_left, plot_top, plot_left, baseline, ink)

            ticks = 5
            for t in range(ticks + 1):
                gy = baseline - int(plot_h * t / ticks)
                if t > 0:
                    self._chart_axis_line(slide, plot_left, gy, plot_left + plot_w, gy,
                                          grid, width_pt=0.5)
                self._add_text(
                    slide, self._fmt_num(nice_max * t / ticks),
                    left=Emu(L), top=Emu(gy - int(Inches(0.12))),
                    width=Emu(y_gutter - int(Inches(0.08))), height=Inches(0.24),
                    font=self.style.typography.body_font, size=9, color=ink,
                    align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE,
                )

            slot = plot_w / n
            bar_w = int(slot * 0.6)
            for i, v in enumerate(values):
                bx = int(plot_left + slot * i + (slot - bar_w) / 2)
                bh = int(plot_h * (v / nice_max)) if nice_max else 0
                by = baseline - bh
                if bh > 0:
                    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(bx), Emu(by),
                                                 Emu(bar_w), Emu(bh))
                    self._fill_solid(bar, bar_fill)
                    bar.line.color.rgb = self._rgb(bar_outline)
                    bar.line.width = Pt(1.0)
                # значение над столбцом
                self._add_text(
                    slide, self._fmt_num(v),
                    left=Emu(int(plot_left + slot * i)), top=Emu(by - int(Inches(0.28))),
                    width=Emu(int(slot)), height=Inches(0.26),
                    font=self.style.typography.heading_font, size=12, bold=True,
                    color=ink, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
                )
                # подпись категории
                self._add_text(
                    slide, cats[i] if i < len(cats) else "",
                    left=Emu(int(plot_left + slot * i)), top=Emu(baseline + int(Inches(0.06))),
                    width=Emu(int(slot)), height=Inches(0.35),
                    font=self.style.typography.body_font, size=10, color=ink,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP,
                )
        else:
            # --- ГОРИЗОНТАЛЬНЫЕ ПОЛОСЫ (bar) ---
            cat_gutter = int(Inches(1.9))  # слева под названия категорий
            plot_left = L + cat_gutter
            plot_top = T
            plot_w = W - cat_gutter
            plot_h = H

            self._chart_axis_line(slide, plot_left, plot_top, plot_left, plot_top + plot_h, ink)

            slot = plot_h / n
            bar_h = int(slot * 0.55)
            for i, v in enumerate(values):
                by = int(plot_top + slot * i + (slot - bar_h) / 2)
                bl = int(plot_w * (v / nice_max)) if nice_max else 0
                if bl > 0:
                    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(plot_left), Emu(by),
                                                 Emu(bl), Emu(bar_h))
                    self._fill_solid(bar, bar_fill)
                    bar.line.color.rgb = self._rgb(bar_outline)
                    bar.line.width = Pt(1.0)
                # категория слева
                self._add_text(
                    slide, cats[i] if i < len(cats) else "",
                    left=Emu(L), top=Emu(by - int(Inches(0.02))),
                    width=Emu(cat_gutter - int(Inches(0.1))), height=Emu(bar_h + int(Inches(0.04))),
                    font=self.style.typography.body_font, size=10, color=ink,
                    align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE,
                )
                # значение в конце полосы
                self._add_text(
                    slide, self._fmt_num(v),
                    left=Emu(int(plot_left + bl + int(Inches(0.05)))), top=Emu(by),
                    width=Inches(0.9), height=Emu(bar_h),
                    font=self.style.typography.heading_font, size=12, bold=True,
                    color=ink, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE,
                )

    def _chart_axis_line(self, slide, x1, y1, x2, y2, color, width_pt: float = 1.0) -> None:
        conn = slide.shapes.add_connector(1, Emu(int(x1)), Emu(int(y1)),
                                          Emu(int(x2)), Emu(int(y2)))
        conn.line.color.rgb = self._rgb(color)
        conn.line.width = Pt(width_pt)

    @staticmethod
    def _fmt_num(v: float) -> str:
        return str(int(v)) if float(v).is_integer() else f"{v:.1f}"

    def _render_native_chart(self, slide, s: ChartSlide, chart_top, chart_height, chart_width) -> None:
        """Нативный pptx-чарт (pie/line, и fallback для column/bar)."""
        chart_data = CategoryChartData()
        chart_data.categories = s.categories
        chart_data.add_series(s.value_label, s.values)
        xl_type = self._CHART_TYPE_MAP.get(s.chart_type, XL_CHART_TYPE.COLUMN_CLUSTERED)
        graphic_frame = slide.shapes.add_chart(
            xl_type, MARGIN_X, chart_top, chart_width, chart_height, chart_data,
        )
        self._style_chart(graphic_frame.chart, s)

    def _style_chart(self, chart, s: ChartSlide) -> None:
        """Применяет палитру дайджеста и базовое оформление к графику."""
        chart.has_title = False

        # Легенду показываем только для pie (там она осмысленна)
        if s.chart_type == "pie":
            chart.has_legend = True
            chart.legend.position = XL_LEGEND_POSITION.RIGHT
            chart.legend.include_in_layout = False
            try:
                chart.legend.font.size = Pt(11)
                chart.legend.font.name = self.style.typography.body_font
                chart.legend.font.color.rgb = self._rgb(self.palette.text_dark)
            except Exception:
                pass
        else:
            chart.has_legend = False

        plot = chart.plots[0]

        # Палитра для точек: accent + вариации. Для pie каждый сектор свой цвет.
        palette_cycle = [
            self.palette.accent,
            self.palette.badge,
            self.palette.text_muted,
            self.palette.gradient_start,
            self.palette.gradient_end,
        ]

        series = plot.series[0]

        if s.chart_type == "pie":
            # Каждый сектор — свой цвет через прямой XML
            self._color_pie_points(series, len(s.categories), palette_cycle)
            plot.has_data_labels = True
            dl = plot.data_labels
            dl.show_percentage = True
            dl.show_value = False
            dl.number_format = "0%"
            dl.number_format_is_linked = False
            try:
                dl.font.size = Pt(11)
                dl.font.color.rgb = self._rgb(self.palette.text_dark)
                dl.font.bold = True
            except Exception:
                pass
        else:
            # bar/column/line — единый акцентный цвет серии.
            # Цвет столбцов считаем КОНТРАСТНЫМ к фону слайда: иначе при
            # светлом акценте на светлом фоне столбцы сливаются и «исчезают»
            # (как было на скрине в PowerPoint). Если акцент задан явно
            # пользователем — оставляем его, но всё равно даём обводку.
            bar_color = (self.palette.accent if self._locked("accent")
                         else self._contrasting_fill(self.palette.accent, 3.0))
            # Контур столбца — гарантирует видимость, даже если PowerPoint
            # по какой-то причине проигнорирует заливку spPr.
            outline = self._contrasting_fill(bar_color, 1.6)

            fmt = series.format
            fmt.fill.solid()
            fmt.fill.fore_color.rgb = self._rgb(bar_color)
            if s.chart_type == "line":
                fmt.line.color.rgb = self._rgb(bar_color)
                fmt.line.width = Pt(2.5)
            else:
                fmt.line.color.rgb = self._rgb(outline)
                fmt.line.width = Pt(1.25)

            # Для одиночной серии выключаем авто-разноцветность: PowerPoint
            # иначе может перекрасить/обнулить заливку через стиль чарта.
            try:
                plot.vary_by_categories = False
            except Exception:
                pass
            # Толщина столбцов (меньше зазор — столбцы заметнее)
            try:
                if s.chart_type in ("column", "bar"):
                    plot.gap_width = 60
            except Exception:
                pass

            # Подписи значений на столбцах/точках — цвет контрастен фону
            label_ink = self._chart_label_ink()
            plot.has_data_labels = True
            dl = plot.data_labels
            dl.show_value = True
            try:
                dl.font.size = Pt(11)
                dl.font.bold = True
                dl.font.color.rgb = self._rgb(label_ink)
                if s.chart_type in ("column", "bar"):
                    dl.position = XL_LABEL_POSITION.OUTSIDE_END
            except Exception:
                pass

            # Оформление осей — тоже контрастно фону
            axis_ink = self._chart_label_ink()
            try:
                cat_ax = chart.category_axis
                cat_ax.tick_labels.font.size = Pt(11)
                cat_ax.tick_labels.font.name = self.style.typography.body_font
                cat_ax.tick_labels.font.color.rgb = self._rgb(axis_ink)
            except Exception:
                pass
            try:
                val_ax = chart.value_axis
                val_ax.tick_labels.font.size = Pt(10)
                val_ax.tick_labels.font.color.rgb = self._rgb(axis_ink)
                val_ax.has_major_gridlines = True
            except Exception:
                pass

    def _color_pie_points(self, series, n_points: int, palette_cycle) -> None:
        """Красит каждый сектор pie в свой цвет (прямой XML-патч dPt)."""
        series_xml = series._element
        # Точка вставки — после служебных элементов
        insert_tags = {qn(t) for t in ("c:idx", "c:order", "c:tx", "c:spPr")}
        insert_idx = 0
        for i, child in enumerate(series_xml):
            if child.tag in insert_tags:
                insert_idx = i + 1

        for i in range(n_points):
            color = palette_cycle[i % len(palette_cycle)].lstrip("#")
            dpt_xml = (
                f'<c:dPt xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
                f'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
                f'<c:idx val="{i}"/><c:bubble3D val="0"/>'
                f'<c:spPr><a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
                f'<a:ln w="19050"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln>'
                f'</c:spPr></c:dPt>'
            )
            series_xml.insert(insert_idx + i, etree.fromstring(dpt_xml))

    # ----------------------------------------------------------------------- #
    # KPI карточки (одинаковая логика для cover и topic)
    # ----------------------------------------------------------------------- #

    def _render_kpi_row(
        self, slide, kpis: List[KPICard],
        top, height,
        value_size: int, label_size: int,
        value_color: str, label_color: str,
        card_bg: Optional[str] = None,
        border_color: Optional[str] = None,
    ) -> None:
        """Ряд KPI-карточек по всей ширине слайда."""
        n = len(kpis)
        total_width = SLIDE_WIDTH - MARGIN_X * 2
        gap = Inches(0.25)
        card_width = Emu(int((total_width - gap * (n - 1)) / n))

        for i, kpi in enumerate(kpis):
            left = MARGIN_X + (card_width + gap) * i
            self._render_kpi_card(
                slide, kpi,
                left=left, top=top,
                width=card_width, height=height,
                value_size=value_size, label_size=label_size,
                value_color=value_color, label_color=label_color,
                card_bg=card_bg, border_color=border_color,
            )

    def _render_kpi_card(
        self, slide, kpi: KPICard,
        left, top, width, height,
        value_size: int, label_size: int,
        value_color: str, label_color: str,
        card_bg: Optional[str], border_color: Optional[str],
    ) -> None:
        """Одна KPI-карточка: значение крупно, подпись мельче."""
        # Фон карточки
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.adjustments[0] = 0.08
        if card_bg:
            self._fill_solid(card, card_bg)
            card.line.fill.background()
            self._apply_subtle_shadow(card)
        else:
            # Прозрачный фон, только обводка
            card.fill.background()

        if border_color:
            card.line.color.rgb = self._rgb(border_color)
            card.line.width = Pt(1.0)

        # Авто-контраст: у карточки со сплошной заливкой цвет текста
        # вычисляется от её фактического фона (а не вслепую из палитры).
        if card_bg:
            value_color = self._ink(card_bg, role="strong",
                                    size_pt=value_size, bold=True)
            label_color = self._ink(card_bg, role="muted", size_pt=label_size)

        # Иконка в правом верхнем углу. Если LLM не задала icon_hint —
        # подбираем брендовую иконку по тексту подписи (для SberF1).
        icon_hint = getattr(kpi, "icon_hint", None)
        if not icon_hint and self.use_brand_icons:
            icon_hint = self._auto_icon_hint(getattr(kpi, "label", None))
        if icon_hint and self._has_any_icon(icon_hint):
            icon_size = Inches(0.28)
            icon_color = value_color if card_bg else self.palette.accent
            self._draw_icon(
                slide, icon_hint,
                left=left + width - icon_size - Inches(0.18),
                top=top + Inches(0.16),
                size=icon_size,
                color=icon_color,
            )

        # Значение
        self._add_text(
            slide, kpi.value,
            left=left, top=top + Inches(0.15),
            width=width, height=Emu(int(height * 0.55)),
            font=self.style.typography.heading_font,
            size=value_size, bold=True, color=value_color,
            align=PP_ALIGN.CENTER,
            anchor=MSO_ANCHOR.MIDDLE,
        )

        # Подпись
        label_top = top + Emu(int(height * 0.6))
        self._add_text(
            slide, kpi.label,
            left=left + Inches(0.1), top=label_top,
            width=width - Inches(0.2), height=Emu(int(height * 0.35)),
            font=self.style.typography.body_font,
            size=label_size, color=label_color,
            align=PP_ALIGN.CENTER,
            anchor=MSO_ANCHOR.TOP,
        )

    # ----------------------------------------------------------------------- #
    # Pill-метки
    # ----------------------------------------------------------------------- #

    def _render_pills_row(
        self, slide, tags: List[str],
        top,
        font_size: int,
        text_color: str,
        bg_color: str,
        start_left=None,
    ) -> None:
        """Ряд pill-меток с динамической шириной."""
        if start_left is None:
            start_left = MARGIN_X

        # Эвристика: ширина пилки = длина текста * шрифт * коэффициент + отступы
        # px_per_char ≈ font_size * 0.55 для русского текста
        # 1pt = 12700 EMU
        pill_height = Inches(0.36)
        pill_gap = Inches(0.15)
        inner_pad_x_emu = Inches(0.3)

        current_left = start_left
        # Авто-контраст: цвет текста pill — от её заливки. Желаемый text_color
        # сохраняется, если читается на bg_color, иначе берётся чёрный/белый.
        pill_ink = self._ink_on(bg_color, prefer=text_color,
                                size_pt=font_size, bold=False)
        for tag in tags:
            # Эвристика ширины pill. Кириллица шире латиницы, поэтому
            # коэффициент щедрый. Размер pt в EMU: 1pt ≈ 12700 EMU.
            # Эмпирически 0.75 даёт корректные ширины для смешанного текста.
            char_width_emu = int(font_size * 0.75 * 12700)
            pill_width = Emu(len(tag) * char_width_emu) + inner_pad_x_emu * 2

            # Если выходит за правую границу — не рисуем больше.
            if current_left + pill_width > SLIDE_WIDTH - MARGIN_X:
                break

            pill = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                current_left, top, pill_width, pill_height,
            )
            pill.adjustments[0] = 0.5  # полная скругленность
            self._fill_solid(pill, bg_color)
            pill.line.fill.background()
            self._apply_subtle_shadow(pill)

            tf = pill.text_frame
            tf.margin_left = inner_pad_x_emu
            tf.margin_right = inner_pad_x_emu
            tf.margin_top = Emu(0); tf.margin_bottom = Emu(0)
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf.text = tag
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            for run in p.runs:
                run.font.name = self.style.typography.body_font
                run.font.size = Pt(font_size)
                run.font.color.rgb = self._rgb(pill_ink)

            current_left = current_left + pill_width + pill_gap

    # ----------------------------------------------------------------------- #
    # Метаданные (шапка и футер)
    # ----------------------------------------------------------------------- #

    def _render_meta_header(self, slide) -> None:
        """Шапка с датой/периодом/номером выпуска (для cover-слайда)."""
        meta = self.spec.meta
        text = f"{meta.issue_date}    •    Период: {meta.period}    •    {meta.issue_number}"

        # Полупрозрачная плашка справа
        pill_width = Inches(7.5)
        pill_left = SLIDE_WIDTH - MARGIN_X - pill_width

        self._add_text(
            slide, text,
            left=pill_left, top=META_HEADER_TOP,
            width=pill_width, height=META_HEADER_HEIGHT,
            font=self.style.typography.body_font,
            size=11, color=self._muted_on_background(),
            align=PP_ALIGN.RIGHT,
            anchor=MSO_ANCHOR.MIDDLE,
        )

    def _render_meta_footer(self, slide, page_number: int, total: int) -> None:
        """Футер с расширенной мета-информацией (для topic-слайдов)."""
        meta = self.spec.meta
        parts = [meta.issue_date, f"Период: {meta.period}", meta.issue_number]
        if meta.next_issue:
            parts.append(f"Следующий выпуск: {meta.next_issue}")
        if meta.note:
            parts.append(meta.note)
        text = "    •    ".join(parts)

        # Слева — мета
        self._add_text(
            slide, text,
            left=MARGIN_X, top=SLIDE_HEIGHT - FOOTER_BOTTOM - FOOTER_HEIGHT,
            width=SLIDE_WIDTH - MARGIN_X * 2 - Inches(0.7),
            height=FOOTER_HEIGHT,
            font=self.style.typography.body_font,
            size=9, color=self._muted_on_background(),
            anchor=MSO_ANCHOR.MIDDLE,
        )

        # Справа — номер страницы
        self._add_text(
            slide, f"стр. {page_number}",
            left=SLIDE_WIDTH - MARGIN_X - Inches(0.7),
            top=SLIDE_HEIGHT - FOOTER_BOTTOM - FOOTER_HEIGHT,
            width=Inches(0.7), height=FOOTER_HEIGHT,
            font=self.style.typography.body_font,
            size=9, color=self._muted_on_background(),
            align=PP_ALIGN.RIGHT,
            anchor=MSO_ANCHOR.MIDDLE,
        )

    # ----------------------------------------------------------------------- #
    # Градиентный фон (прямой XML)
    # ----------------------------------------------------------------------- #

    def _contrasting_fill(self, color: str, min_ratio: float = 3.0) -> str:
        """Цвет заливки, который отделяется от фона слайда (по обоим стопам
        градиента) минимум на min_ratio. Если исходный color уже виден —
        возвращаем его; иначе берём ближайший видимый вариант (затемнённый/
        осветлённый или, в крайнем случае, почти чёрный/белый)."""
        stops = self._gradient_stops()
        if C.worst_contrast_over_gradient(color, stops) >= min_ratio:
            return C.to_hex(color, with_hash=False)
        cands = [
            color,
            C.to_hex(C.darken(color, 0.4), with_hash=False),
            C.to_hex(C.lighten(color, 0.4), with_hash=False),
            "111111", "F0F0F0",
        ]
        return C.to_hex(
            C.best_text_color_over_gradient(stops, candidates=cands),
            with_hash=False,
        )

    def _chart_label_ink(self) -> str:
        """Цвет подписей/осей графика — читаемый поверх фона слайда."""
        return self._text_on_background()

    def _apply_background_gradient(self, slide, angle_deg: float) -> None:
        """Кладёт градиент на фон слайда ЧЕРЕЗ штатный API python-pptx.

        Важно: раньше мы дописывали <p:bgPr> сырым XML прямо в <p:cSld>,
        из-за чего получалась невалидная структура — PowerPoint её
        игнорировал и оставлял фон белым (хотя LibreOffice рисовал). API
        создаёт <p:bg><p:bgPr><a:gradFill> в правильном месте, и фон
        корректно рендерится во всех редакторах.
        """
        start = self.palette.gradient_start.lstrip("#")
        end = self.palette.gradient_end.lstrip("#")

        fill = slide.background.fill
        fill.gradient()
        stops = fill.gradient_stops
        stops[0].position = 0.0
        stops[0].color.rgb = self._rgb(start)
        stops[1].position = 1.0
        stops[1].color.rgb = self._rgb(end)
        try:
            fill.gradient_angle = angle_deg
        except Exception:
            pass

    def _set_premium_background(self, slide) -> None:
        """«Премиум»-фон: диагональный градиент (угол 45°)."""
        self._apply_background_gradient(slide, 45.0)

    def _set_gradient_background(self, slide) -> None:
        """Простой горизонтальный градиент фона (надёжный fallback)."""
        self._apply_background_gradient(slide, 0.0)

    # ----------------------------------------------------------------------- #
    # Утилиты текста и стиля
    # ----------------------------------------------------------------------- #

    def _add_text(
        self, slide, text: str,
        left, top, width, height,
        font: str, size: int, color: str = None,
        bold: bool = False, italic: bool = False,
        align=PP_ALIGN.LEFT,
        anchor=MSO_ANCHOR.TOP,
        line_spacing: float | None = None,
        on: str | None = None,
        role: str = "strong",
    ) -> None:
        """Единая точка создания текстового бокса.

        on   — hex заливки, НА КОТОРОЙ лежит текст. Если задан, цвет текста
               вычисляется от этого фона (тёмный объект → светлый текст,
               светлый объект → тёмный), с учётом «замков» пользователя.
               Это и есть пообъектный контраст.
        role — 'strong' (основной) или 'muted' (вторичный) — влияет на выбор
               предпочтительного цвета и порог контраста.
        color — явный цвет (для текста на градиенте, где фон не сплошной).
                Используется, только если on не задан.
        line_spacing — множитель межстрочного интервала (1.0 = одинарный).
        """
        if on is not None:
            color = self._ink(on, role=role, size_pt=size, bold=bold)
        elif color is None:
            color = self._text_on_background()
        box = slide.shapes.add_textbox(left, top, width, height)
        tf = box.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = anchor
        tf.margin_left = Emu(0); tf.margin_right = Emu(0)
        tf.margin_top = Emu(0); tf.margin_bottom = Emu(0)
        tf.text = text
        p = tf.paragraphs[0]
        p.alignment = align
        if line_spacing:
            p.line_spacing = line_spacing
        for run in p.runs:
            run.font.name = font
            run.font.size = Pt(size)
            run.font.bold = bold
            run.font.italic = italic
            run.font.color.rgb = self._rgb(color)

    def _apply_subtle_shadow(self, shape) -> None:
        """Лёгкая тень — выделяет карточки на градиентном фоне."""
        # python-pptx не имеет высокоуровневого API для shadow,
        # патчим напрямую через spPr/effectLst
        sp_pr = shape.fill._xPr  # spPr
        # Если уже есть effectLst — не дублируем
        if sp_pr.find(qn("a:effectLst")) is not None:
            return
        effect_xml = """
        <a:effectLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:outerShdw blurRad="50800" dist="25400" dir="5400000" algn="t" rotWithShape="0">
                <a:srgbClr val="000000"><a:alpha val="15000"/></a:srgbClr>
            </a:outerShdw>
        </a:effectLst>
        """
        sp_pr.append(etree.fromstring(effect_xml))

    def _fill_solid(self, shape, hex_color: str) -> None:
        shape.fill.solid()
        shape.fill.fore_color.rgb = self._rgb(hex_color)

    @staticmethod
    def _rgb(hex_color: str) -> RGBColor:
        return RGBColor.from_string(hex_color.lstrip("#"))

    @staticmethod
    def _luminance(hex_color: str) -> float:
        """
        Относительная яркость цвета (0=чёрный, 1=белый).

        Нужна, чтобы выбирать контрастный цвет текста на любом фоне
        независимо от темы. Формула — взвешенная по восприятию (sRGB).
        """
        h = hex_color.lstrip("#")
        if len(h) != 6:
            return 0.5
        r, g, b = (int(h[i:i+2], 16) / 255.0 for i in (0, 2, 4))
        return 0.2126 * r + 0.7152 * g + 0.0722 * b

    # ----------------------------------------------------------------------- #
    # Авто-контраст (WCAG). Единая точка выбора цвета текста под фон.
    # ----------------------------------------------------------------------- #

    def _gradient_stops(self) -> list[str]:
        return [self.palette.gradient_start, self.palette.gradient_end]

    def _ink(self, bg: str, *, role: str = "strong",
             size_pt: float = 14, bold: bool = False) -> str:
        """Цвет текста, читаемого на СПЛОШНОМ фоне bg (hex).

        role='strong' — основной текст, предпочитаем palette.text_dark;
        role='muted'  — вторичный текст, предпочитаем palette.text_muted.

        Если пользователь явно зафиксировал цвет текста — возвращаем его
        как есть (даже при низком контрасте: «делаю как просил»). Иначе:
        сохраняем фирменный тон, если он проходит WCAG; не проходит —
        подтягиваем тоном, в крайнем случае чистый чёрный/белый.
        """
        prefer = self.palette.text_muted if role == "muted" else self.palette.text_dark
        lock_roles = ("text", "text_muted") if role == "muted" else ("text",)
        if self._locked(*lock_roles):
            return C.to_hex(prefer, with_hash=False)
        min_ratio = C.min_ratio_for_font(size_pt, bold)
        if role == "muted":
            min_ratio = max(C.AA_LARGE, min_ratio - 1.0)
        # Фирменный цвет проходит по контрасту — оставляем его.
        if C.contrast_ratio(prefer, bg) >= min_ratio:
            return C.to_hex(prefer, with_hash=False)
        # Не проходит — ЧЁТКИЙ чёрный/белый (без мутных полутонов).
        base = C.best_text_color(bg)
        if role == "muted":
            soft = C._mix(base, C.parse_color(bg), 0.28)
            if C.contrast_ratio(soft, bg) >= C.AA_LARGE:
                return C.to_hex(soft, with_hash=False)
        return C.to_hex(base, with_hash=False)

    def _ink_on(self, fill_bg: str, *, prefer: Optional[str] = None,
                size_pt: float = 12, bold: bool = True) -> str:
        """Цвет текста на ЦВЕТНОЙ заливке фигуры (бейдж, кружок, плашка).

        prefer — желаемый «вывернутый» цвет (например palette.kpi_bg для
        белого текста на акценте). Если он проходит контраст — берём его,
        иначе — лучший чёрный/белый. Так на белом фоне чёрная фигура
        получает белый текст, на жёлтом бейдже — чёрный, и т.д.
        """
        min_ratio = C.min_ratio_for_font(size_pt, bold)
        if prefer is not None and C.contrast_ratio(prefer, fill_bg) >= min_ratio:
            return C.to_hex(prefer, with_hash=False)
        return C.to_hex(C.best_text_color(fill_bg), with_hash=False)

    def _accent_ink(self, bg: str) -> str:
        """Акцентный цвет текста (число), но гарантированно видимый на bg."""
        return C.to_hex(C.ensure_contrast(self.palette.accent, bg, C.AA_LARGE),
                        with_hash=False)

    def _text_on_background(self) -> str:
        """Основной текст поверх ГРАДИЕНТА фона.

        Считаем по maximin сразу по обоим стопам: текст не должен
        провалиться ни на светлом, ни на тёмном конце градиента.
        Фирменный text_dark сохраняем, если он читается на всём градиенте.
        """
        stops = self._gradient_stops()
        prefer = self.palette.text_dark
        # Пользователь явно задал цвет текста — рендерим как есть.
        if self._locked("text"):
            return C.to_hex(prefer, with_hash=False)
        if C.worst_contrast_over_gradient(prefer, stops) >= C.AA_NORMAL:
            return C.to_hex(prefer, with_hash=False)
        rgb = C.best_text_color_over_gradient(
            stops, candidates=[prefer, "FFFFFF", "0A0A0A"])
        return C.to_hex(rgb, with_hash=False)

    def _muted_on_background(self) -> str:
        """Приглушённый текст поверх градиента (подзаголовки, мета)."""
        stops = self._gradient_stops()
        prefer = self.palette.text_muted
        if self._locked("text", "text_muted"):
            return C.to_hex(prefer, with_hash=False)
        if C.worst_contrast_over_gradient(prefer, stops) >= C.AA_LARGE:
            return C.to_hex(prefer, with_hash=False)
        rgb = C.best_text_color_over_gradient(
            stops, candidates=[prefer, "C8CCD8", "5A6B7D", "FFFFFF", "1A1A1A"])
        return C.to_hex(rgb, with_hash=False)
