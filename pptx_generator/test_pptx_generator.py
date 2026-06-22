"""
Тесты builder'а и схем без обращения к LLM.

Идея: builder и схемы должны быть полностью покрыты тестами,
потому что они детерминированы. LLM-часть отдельно через мок.
"""
import json
import tempfile
from pathlib import Path

import pandas as pd
import pytest

from pptx_generator import (
    BulletsSlide,
    ChartData,
    ChartSeries,
    ChartSlide,
    ChartType,
    ChartWithTextSlide,
    ClosingSlide,
    ColorPalette,
    ExcelReader,
    KPIItem,
    KPISlide,
    PresentationBuilder,
    PresentationSpec,
    PresentationStyle,
    QuoteSlide,
    SectionHeaderSlide,
    TitleSlide,
    TwoColumnSlide,
    Typography,
)
from pptx_generator.tool import extract_json, parse_spec


# --------------------------------------------------------------------------- #
# Тесты схем
# --------------------------------------------------------------------------- #

class TestColorPalette:
    def test_valid_hex_accepted(self):
        p = ColorPalette(
            primary="1E2761", secondary="CADCFC", accent="FFFFFF",
            background="FFFFFF", text_dark="000000", text_light="FFFFFF",
        )
        assert p.primary == "1E2761"

    def test_hash_prefix_stripped(self):
        p = ColorPalette(
            primary="#1E2761", secondary="CADCFC", accent="FFFFFF",
            background="FFFFFF", text_dark="000000", text_light="FFFFFF",
        )
        assert p.primary == "1E2761"

    def test_invalid_hex_rejected(self):
        with pytest.raises(ValueError, match="Некорректный hex-цвет"):
            ColorPalette(
                primary="XYZXYZ", secondary="CADCFC", accent="FFFFFF",
                background="FFFFFF", text_dark="000000", text_light="FFFFFF",
            )

    def test_short_hex_rejected(self):
        with pytest.raises(ValueError):
            ColorPalette(
                primary="FFF", secondary="CADCFC", accent="FFFFFF",
                background="FFFFFF", text_dark="000000", text_light="FFFFFF",
            )


class TestChartData:
    def test_pie_with_multiple_series_rejected(self):
        with pytest.raises(ValueError, match="Pie chart"):
            ChartData(
                chart_type=ChartType.PIE,
                title="x", categories=["a", "b"],
                series=[
                    ChartSeries(name="s1", values=[1, 2]),
                    ChartSeries(name="s2", values=[3, 4]),
                ],
            )

    def test_mismatched_lengths_rejected(self):
        with pytest.raises(ValueError, match="должны совпадать"):
            ChartData(
                chart_type=ChartType.BAR,
                title="x", categories=["a", "b", "c"],
                series=[ChartSeries(name="s", values=[1, 2])],
            )

    def test_scatter_allows_mismatched_categories(self):
        # Scatter не требует категорий
        cd = ChartData(
            chart_type=ChartType.SCATTER,
            title="x", categories=[],
            series=[
                ChartSeries(name="X", values=[1, 2, 3]),
                ChartSeries(name="Y", values=[10, 20, 30]),
            ],
        )
        assert cd.chart_type == ChartType.SCATTER


# --------------------------------------------------------------------------- #
# Тесты ExcelReader
# --------------------------------------------------------------------------- #

@pytest.fixture
def sample_xlsx():
    """Создаёт тестовый xlsx с обоими нужными листами."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = Path(f.name)

    raw_df = pd.DataFrame({
        "Регион": ["Москва", "СПб", "Казань", "Москва", "СПб"],
        "Выручка": [1200000, 850000, 430000, 1500000, 920000],
        "Жалобы": [12, 8, 4, 15, 9],
    })
    summary_df = pd.DataFrame([
        ["Основные проблемы:"],
        ["1. Высокий рост жалоб в Москве: с 12 до 15 за квартал"],
        ["2. Падение качества обслуживания в регионах"],
        ["3. Необходима ревизия процессов в топ-3 городах"],
    ])

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        raw_df.to_excel(writer, sheet_name="Сырье", index=False)
        summary_df.to_excel(writer, sheet_name="Суммаризация", index=False, header=False)

    yield path
    path.unlink(missing_ok=True)


class TestExcelReader:
    def test_reads_both_sheets(self, sample_xlsx):
        ctx = ExcelReader(sample_xlsx).read()
        assert ctx.row_count == 5
        assert ctx.column_count == 3
        assert "Москва" in ctx.summary_text or "проблем" in ctx.summary_text.lower()

    def test_numeric_column_profiled(self, sample_xlsx):
        ctx = ExcelReader(sample_xlsx).read()
        revenue = next(c for c in ctx.columns if c.name == "Выручка")
        assert revenue.min_value == 430000
        assert revenue.max_value == 1500000

    def test_categorical_column_has_top_values(self, sample_xlsx):
        ctx = ExcelReader(sample_xlsx).read()
        region = next(c for c in ctx.columns if c.name == "Регион")
        assert "Москва" in region.top_values
        assert region.top_values["Москва"] == 2

    def test_missing_sheet_raises(self, tmp_path):
        # xlsx только с одним листом
        bad_path = tmp_path / "bad.xlsx"
        pd.DataFrame({"a": [1]}).to_excel(bad_path, sheet_name="Other", index=False)

        with pytest.raises(ValueError, match="отсутствуют листы"):
            ExcelReader(bad_path).read()

    def test_nonexistent_file_raises(self):
        with pytest.raises(FileNotFoundError):
            ExcelReader("/nonexistent/file.xlsx").read()


# --------------------------------------------------------------------------- #
# Тесты PresentationBuilder
# --------------------------------------------------------------------------- #

def _make_style() -> PresentationStyle:
    return PresentationStyle(
        palette=ColorPalette(
            primary="1E2761", secondary="CADCFC", accent="F96167",
            background="FFFFFF", text_dark="212121", text_light="FFFFFF",
        ),
        typography=Typography(heading_font="Calibri", body_font="Calibri"),
        mood="professional",
    )


class TestPresentationBuilder:
    def test_builds_all_layouts(self, tmp_path):
        """Smoke test: каждый тип слайда должен рендериться без падений."""
        spec = PresentationSpec(
            title="Тестовая презентация",
            author="QA",
            style=_make_style(),
            slides=[
                TitleSlide(title="Главный заголовок", subtitle="Подзаголовок"),
                SectionHeaderSlide(title="Раздел 1", description="Описание"),
                BulletsSlide(title="Буллеты", bullets=["Раз", "Два", "Три"]),
                TwoColumnSlide(
                    title="Сравнение",
                    left_heading="Плюсы", left_bullets=["a", "b"],
                    right_heading="Минусы", right_bullets=["c", "d"],
                ),
                ChartSlide(
                    title="Bar chart",
                    chart=ChartData(
                        chart_type=ChartType.BAR,
                        title="Выручка",
                        categories=["Москва", "СПб", "Казань"],
                        series=[ChartSeries(name="Q1", values=[1200, 850, 430])],
                    ),
                ),
                ChartSlide(
                    title="Line chart",
                    chart=ChartData(
                        chart_type=ChartType.LINE,
                        title="Тренд",
                        categories=["Янв", "Фев", "Мар", "Апр"],
                        series=[
                            ChartSeries(name="2024", values=[100, 120, 150, 180]),
                            ChartSeries(name="2025", values=[110, 140, 170, 220]),
                        ],
                    ),
                ),
                ChartSlide(
                    title="Pie chart",
                    chart=ChartData(
                        chart_type=ChartType.PIE,
                        title="Доли",
                        categories=["A", "B", "C", "D"],
                        series=[ChartSeries(name="Share", values=[40, 30, 20, 10])],
                    ),
                ),
                ChartSlide(
                    title="Scatter chart",
                    chart=ChartData(
                        chart_type=ChartType.SCATTER,
                        title="Корреляция",
                        categories=[],
                        series=[
                            ChartSeries(name="X", values=[1, 2, 3, 4, 5]),
                            ChartSeries(name="Y", values=[2, 4, 5, 4, 6]),
                        ],
                    ),
                ),
                ChartWithTextSlide(
                    title="График + текст",
                    chart=ChartData(
                        chart_type=ChartType.COLUMN,
                        title="По кварталам",
                        categories=["Q1", "Q2", "Q3"],
                        series=[ChartSeries(name="Продажи", values=[100, 130, 160])],
                    ),
                    insights=["Рост 30% за период", "Q3 — лучший квартал"],
                ),
                KPISlide(
                    title="KPI",
                    kpis=[
                        KPIItem(value="+24%", label="Рост выручки"),
                        KPIItem(value="1.2M", label="Активных пользователей"),
                        KPIItem(value="98%", label="Удовлетворённость"),
                    ],
                ),
                QuoteSlide(quote="Данные — это новая нефть.", attribution="К. Хамби"),
                ClosingSlide(title="Спасибо!", subtitle="Вопросы?"),
            ],
        )

        out = tmp_path / "test.pptx"
        result = PresentationBuilder(spec).build(out)
        assert result.exists()
        assert result.stat().st_size > 0


# --------------------------------------------------------------------------- #
# Тесты парсинга LLM-ответа
# --------------------------------------------------------------------------- #

class TestExtractJSON:
    def test_plain_json(self):
        assert extract_json('{"a": 1}') == '{"a": 1}'

    def test_with_fence(self):
        raw = '```json\n{"a": 1}\n```'
        assert extract_json(raw) == '{"a": 1}'

    def test_with_preamble(self):
        raw = 'Вот ваша презентация:\n{"a": 1}\nГотово!'
        assert extract_json(raw) == '{"a": 1}'

    def test_handles_nested_braces(self):
        raw = 'prefix {"a": {"b": 2}} suffix'
        assert extract_json(raw) == '{"a": {"b": 2}}'


class TestParseSpec:
    def test_parses_minimal_valid(self):
        data = {
            "title": "Test",
            "style": {
                "palette": {
                    "primary": "1E2761", "secondary": "CADCFC", "accent": "F96167",
                    "background": "FFFFFF", "text_dark": "212121", "text_light": "FFFFFF",
                },
                "typography": {"heading_font": "Calibri", "body_font": "Calibri"},
                "mood": "professional",
            },
            "slides": [{"layout": "title", "title": "Hello"}],
        }
        spec = parse_spec(json.dumps(data))
        assert spec.title == "Test"
        assert len(spec.slides) == 1

    def test_rejects_bad_json(self):
        with pytest.raises(ValueError, match="невалидный JSON"):
            parse_spec("это не json")

    def test_rejects_schema_violation(self):
        data = {"title": "Test"}  # нет style и slides
        with pytest.raises(ValueError, match="схеме"):
            parse_spec(json.dumps(data))


# --------------------------------------------------------------------------- #
# Тесты программной защиты от рекомендаций
# --------------------------------------------------------------------------- #

class TestRecommendationSanitizer:
    """
    Проверяет, что схема программно фильтрует рекомендации даже если
    LLM проигнорировал запрет в промпте.
    """

    def _base_spec(self, slides):
        return PresentationSpec(
            title="Test",
            style=_make_style(),
            slides=slides,
        )

    def test_renames_recommendation_title(self):
        from pptx_generator import BulletsSlide
        spec = self._base_spec([
            TitleSlide(title="Start"),
            BulletsSlide(title="Рекомендации", bullets=["Факт 1", "Факт 2"]),
        ])
        assert spec.slides[1].title == "Итоги анализа"

    def test_renames_action_plan_title(self):
        from pptx_generator import BulletsSlide
        spec = self._base_spec([
            TitleSlide(title="Start"),
            BulletsSlide(title="План действий на квартал", bullets=["Факт"]),
        ])
        assert spec.slides[1].title == "Итоги анализа"

    def test_keeps_neutral_title(self):
        from pptx_generator import BulletsSlide
        original = "Главные наблюдения"
        spec = self._base_spec([
            TitleSlide(title="Start"),
            BulletsSlide(title=original, bullets=["Факт"]),
        ])
        assert spec.slides[1].title == original

    def test_filters_imperative_bullets(self):
        from pptx_generator import BulletsSlide
        spec = self._base_spec([
            TitleSlide(title="Start"),
            BulletsSlide(
                title="Итоги",
                bullets=[
                    "Выручка выросла на 24%",                # факт — оставить
                    "Следует внедрить новую систему SLA",    # императив — убрать
                    "Москва генерирует 47% жалоб",           # факт — оставить
                    "Необходимо провести обучение",          # императив — убрать
                ],
            ),
        ])
        bullets = spec.slides[1].bullets
        assert "выросла на 24%" in bullets[0]
        assert "47% жалоб" in bullets[1]
        assert len(bullets) == 2

    def test_filters_in_two_column(self):
        spec = self._base_spec([
            TitleSlide(title="Start"),
            TwoColumnSlide(
                title="Сравнение",
                left_heading="Левое",
                left_bullets=["Факт левый", "Необходимо что-то сделать"],
                right_heading="Правое",
                right_bullets=["Следует переделать", "Факт правый"],
            ),
        ])
        s = spec.slides[1]
        assert len(s.left_bullets) == 1 and "Факт левый" in s.left_bullets[0]
        assert len(s.right_bullets) == 1 and "Факт правый" in s.right_bullets[0]

    def test_placeholder_when_all_filtered(self):
        """Если все буллеты — императивы, остаётся плейсхолдер."""
        from pptx_generator import BulletsSlide
        spec = self._base_spec([
            TitleSlide(title="Start"),
            BulletsSlide(
                title="Итоги",
                bullets=["Следует внедрить X", "Необходимо сделать Y"],
            ),
        ])
        bullets = spec.slides[1].bullets
        assert len(bullets) == 1
        assert "данных" in bullets[0].lower()  # placeholder упоминает «данные»


# --------------------------------------------------------------------------- #
# Тесты лимитов длины
# --------------------------------------------------------------------------- #

class TestLengthLimits:
    """Защита от переполнения вёрстки за счёт длинных текстов."""

    def test_bullet_truncated(self):
        from pptx_generator import BulletsSlide
        slide = BulletsSlide(
            title="Test",
            bullets=["x" * 300],  # очень длинный
        )
        assert len(slide.bullets[0]) <= 140

    def test_category_truncated(self):
        from pptx_generator import ChartData, ChartSeries, ChartType
        cd = ChartData(
            chart_type=ChartType.BAR,
            title="Test",
            categories=["a" * 100, "b"],
            series=[ChartSeries(name="s", values=[1, 2])],
        )
        assert len(cd.categories[0]) <= 25

    def test_title_max_length_enforced(self):
        # Заголовок > 80 симв должен вызвать ValidationError
        with pytest.raises(ValueError):
            TitleSlide(title="x" * 200)
