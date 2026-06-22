"""
pptx_generator — LangChain-инструмент для генерации дайджеста .pptx через GigaChat.

Главный публичный API:
    GenerateDigestTool — LangChain Tool
    DigestBuilder      — низкоуровневая сборка из спеки
    ExcelReader        — чтение и профилирование xlsx
    DigestSpec         — Pydantic-схема дайджеста
"""
from .builder import DigestBuilder
from .excel_reader import DataContext, ExcelReader
from .json_repair import repair_json
from .themes import detect_theme, resolve_palette, available_themes, THEME_PRESETS
from .contrast import (
    best_text_color,
    contrast_ratio,
    ensure_contrast,
    relative_luminance,
)
from .safe import with_fallback
from .style_resolver import parse_style_directives, merge_palette
from .brand_icons import available_brand_hints, has_brand_icon
from .schemas import (
    AttentionItem,
    AttentionSlide,
    ChartSlide,
    ClosingSlide,
    ColorPalette,
    CoverSlide,
    DigestMeta,
    DigestSpec,
    DigestStyle,
    ExecutiveSummarySlide,
    KPICard,
    PatternItem,
    PatternsSlide,
    SummaryPoint,
    TopicItem,
    TopicSlide,
    Typography,
)
from .tool import (
    GenerateDigestInput,
    GenerateDigestTool,
    # backward-compat aliases
    GeneratePresentationInput,
    GeneratePresentationTool,
)

__all__ = [
    # Tool
    "GenerateDigestTool",
    "GenerateDigestInput",
    # Backward compatibility
    "GeneratePresentationTool",
    "GeneratePresentationInput",
    # Builder
    "DigestBuilder",
    # Reader
    "ExcelReader",
    "DataContext",
    # Schemas
    "DigestSpec",
    "DigestStyle",
    "DigestMeta",
    "ColorPalette",
    "Typography",
    "CoverSlide",
    "TopicSlide",
    "TopicItem",
    "KPICard",
    "ExecutiveSummarySlide",
    "SummaryPoint",
    "PatternsSlide",
    "PatternItem",
    "AttentionSlide",
    "AttentionItem",
    "ClosingSlide",
    "ChartSlide",
    # JSON repair (полезно для отладки)
    "repair_json",
    "detect_theme",
    "resolve_palette",
    "available_themes",
    # Contrast / style engine
    "best_text_color",
    "contrast_ratio",
    "ensure_contrast",
    "relative_luminance",
    "with_fallback",
    "parse_style_directives",
    "merge_palette",
    # Brand icons (SberF1)
    "available_brand_hints",
    "has_brand_icon",
]
__version__ = "0.5.0"
