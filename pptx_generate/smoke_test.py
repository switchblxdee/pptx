"""
Smoke-test для дайджест-формата.

Так как pydantic недоступен в этой среде, подменяем модели dataclass'ами
с тем же интерфейсом. Это проверяет рендер-код независимо от валидации.
"""
import sys
import types
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional


# --------------------------------------------------------------------------- #
# Mock-схемы (dataclass-аналоги pydantic-моделей)
# --------------------------------------------------------------------------- #

@dataclass
class ColorPalette:
    gradient_start: str
    gradient_end: str
    card_bg: str
    kpi_bg: str
    text_dark: str
    text_muted: str
    accent: str
    badge: str


@dataclass
class Typography:
    heading_font: str = "Calibri"
    body_font: str = "Calibri"


@dataclass
class DigestStyle:
    palette: ColorPalette
    typography: Typography = field(default_factory=Typography)


@dataclass
class DigestMeta:
    issue_date: str
    period: str
    issue_number: str
    next_issue: Optional[str] = None
    note: Optional[str] = "Подготовлен автоматически"


@dataclass
class KPICard:
    value: str
    label: str
    icon_hint: Optional[str] = None


@dataclass
class CoverSlide:
    title: str
    subtitle: str
    source_tags: List[str]
    kpis: List[KPICard]
    description: Optional[str] = None


@dataclass
class TopicItem:
    title: str
    period: str
    mentions: int
    quote: Optional[str] = None
    is_new: bool = False


@dataclass
class TopicSlide:
    title: str
    kpis: List[KPICard]
    items: List[TopicItem]
    source_tags: List[str] = field(default_factory=list)


@dataclass
class DigestSpec:
    style: DigestStyle
    meta: DigestMeta
    cover: CoverSlide
    topics: List[TopicSlide]


# Подменяем schemas-модуль
mock_schemas = types.ModuleType("schemas")
for cls in [ColorPalette, Typography, DigestStyle, DigestMeta, KPICard,
            CoverSlide, TopicItem, TopicSlide, DigestSpec]:
    setattr(mock_schemas, cls.__name__, cls)
sys.modules["pptx_generator.schemas"] = mock_schemas

# Создаём пакет
pkg = types.ModuleType("pptx_generator")
pkg.__path__ = [str(Path(__file__).parent / "src" / "pptx_generator")]
sys.modules["pptx_generator"] = pkg
sys.modules["pptx_generator.schemas"] = mock_schemas

# Импортируем builder
from pptx_generator.builder import DigestBuilder  # noqa: E402


# --------------------------------------------------------------------------- #
# Создание дайджеста (близко к референсу «Голос IT»)
# --------------------------------------------------------------------------- #

def main():
    spec = DigestSpec(
        style=DigestStyle(
            palette=ColorPalette(
                # Бирюзовый → персиковый, как на референсе
                gradient_start="A8D8D5",   # бирюзовый слева
                gradient_end="F5D5BA",     # персиковый справа
                card_bg="FFE8EA",          # пастельный розовый для карточки тем
                kpi_bg="FFFFFF",           # белые KPI-карточки
                text_dark="1A1A2E",        # тёмно-синий текст
                text_muted="6B7280",       # серый для метаданных
                accent="2C5F5D",           # тёмно-бирюзовый акцент
                badge="F97316",            # оранжевый для бейджа new
            ),
            typography=Typography(heading_font="Calibri", body_font="Calibri"),
        ),
        meta=DigestMeta(
            issue_date="28 мая 2026",
            period="21–27 мая 2026",
            issue_number="№ 1 / еженедельный",
            next_issue="1 июня 2026",
            note="Подготовлен автоматически",
        ),
        cover=CoverSlide(
            title="Голос IT",
            subtitle="дайджест для руководства Блока T",
            description=(
                "Темы, волнующие сотрудников по продуктам программы PDLC "
                "и вендорозамещениям"
            ),
            source_tags=[
                "Чаты поддержки и сообщества в СберЧате",
                "Обращения в SberF1",
                "Открытые диалоги",
                "Опросы",
            ],
            kpis=[
                KPICard(value="847", label="Сигналов проанализировано"),
                KPICard(value="6", label="Продуктов под наблюдением"),
                KPICard(value="1", label="Активных тем"),
                KPICard(value="4", label="Новые темы"),
            ],
        ),
        topics=[
            TopicSlide(
                title="PDLC: GigaCode CLI",
                kpis=[
                    KPICard(value="4", label="Источника"),
                    KPICard(value="108", label="Сигналов"),
                    KPICard(value="3", label="Активных темы"),
                    KPICard(value="1", label="Новая тема"),
                ],
                items=[
                    TopicItem(
                        title="Замедление в работе",
                        quote="Второй день после обеда все жутко висит, работать не возможно",
                        period="2Q 2026",
                        mentions=23,
                        is_new=False,
                    ),
                    TopicItem(
                        title="Прерывание работа из-за жестких критериев работы Цензора",
                        quote="разработчик с фамилией Медведев — это блочит нам работу с ИИ",
                        period="2Q 2026",
                        mentions=10,
                        is_new=False,
                    ),
                    TopicItem(
                        title="Проблемы при вызове инструментов из-за затирания секретов",
                        quote=(
                            "При перезапуске варм под windows зашифрованные секреты слетают "
                            "и приходится каждый раз восстанавливать"
                        ),
                        period="на анализе у команды",
                        mentions=8,
                        is_new=True,
                    ),
                ],
                source_tags=["#Ai in Dev Community", "#GigaIDE Support"],
            ),
            TopicSlide(
                title="SberF1: Поиск",
                kpis=[
                    KPICard(value="2", label="Источника"),
                    KPICard(value="54", label="Сигналов"),
                    KPICard(value="2", label="Активных темы"),
                    KPICard(value="0", label="Новые темы"),
                ],
                items=[
                    TopicItem(
                        title="Медленный поиск по большим документам",
                        quote=(
                            "Когда индекс больше 10к страниц, поиск выдаёт результат "
                            "через 30+ секунд"
                        ),
                        period="2Q 2026",
                        mentions=17,
                    ),
                    TopicItem(
                        title="Релевантность результатов снизилась после обновления",
                        quote="После апдейта на прошлой неделе ищет совсем не то, что нужно",
                        period="2Q 2026",
                        mentions=12,
                    ),
                ],
                source_tags=["#SberF1 Support", "#Поиск-чат"],
            ),
        ],
    )

    out = Path("/home/claude/gigachat_pptx_tool/examples/digest_sample.pptx")
    out.parent.mkdir(parents=True, exist_ok=True)
    result = DigestBuilder(spec).build(out)
    print(f"OK: {result} ({result.stat().st_size:,} bytes)")


if __name__ == "__main__":
    main()
