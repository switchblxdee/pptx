"""
Чтение и предварительный анализ входного xlsx.

Ожидаем два листа:
- "Сырье" — табличные данные (любые: числа, текст, мешанина)
- "Суммаризация" — текст с описанием проблем

Модуль возвращает структурированный DataContext, который потом
передаётся в LLM в текстовом виде. LLM не видит сырой xlsx —
он видит уже подготовленный, очищенный, компактный summary.

Это критично: подавать сырые 10к строк в контекст бессмысленно
и дорого. Лучше дать LLM качественную сводку + сэмплы.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd


RAW_SHEET = "Сырье"
SUMMARY_SHEET = "Суммаризация"

# Эвристика: если колонка числовая и в ней разумное число уникальных значений —
# она годится для агрегаций/графиков.
MAX_SAMPLE_ROWS = 15           # сколько строк показать LLM как сэмпл
MAX_UNIQUE_FOR_CATEGORY = 50   # больше — уже не категориальная переменная

# Кандидаты на колонку группировки (1 уникальное значение = 1 topic-слайд).
# Поиск нечувствителен к регистру и частичному совпадению.
GROUPING_COLUMN_CANDIDATES = (
    "объект сигнала",
    "объект",
    "продукт",
    "тема",
    "сервис",
    "система",
    "категория",
    "направление",
    "компонент",
)


@dataclass
class ColumnProfile:
    """Профиль одной колонки — то, что увидит LLM."""
    name: str
    dtype: str
    non_null_count: int
    null_count: int
    unique_count: int
    sample_values: List[Any] = field(default_factory=list)
    # Только для числовых
    min_value: Optional[float] = None
    max_value: Optional[float] = None
    mean_value: Optional[float] = None
    # Топ-значения для категориальных
    top_values: Dict[str, int] = field(default_factory=dict)

    def to_prompt_text(self) -> str:
        lines = [f"  • {self.name} ({self.dtype})"]
        lines.append(f"      не пусто: {self.non_null_count}, пусто: {self.null_count}, уникальных: {self.unique_count}")

        if self.min_value is not None:
            lines.append(
                f"      min={self.min_value:.2f}, max={self.max_value:.2f}, mean={self.mean_value:.2f}"
            )
        if self.top_values:
            top_str = ", ".join(f"{k!r}={v}" for k, v in list(self.top_values.items())[:5])
            lines.append(f"      топ: {top_str}")
        if self.sample_values:
            sample_str = ", ".join(repr(v) for v in self.sample_values[:5])
            lines.append(f"      примеры: {sample_str}")
        return "\n".join(lines)


@dataclass
class GroupingInfo:
    """
    Информация о группировке данных по слайдам.

    Определяет, по какой колонке резать данные на topic-слайды и
    сколько уникальных групп есть. LLM использует это, чтобы решить,
    сколько слайдов делать.
    """
    column_name: Optional[str]               # найденная колонка группировки
    unique_values: List[str] = field(default_factory=list)
    value_counts: Dict[str, int] = field(default_factory=dict)
    detected_automatically: bool = True

    @property
    def group_count(self) -> int:
        return len(self.unique_values)

    def to_prompt_text(self) -> str:
        if not self.column_name:
            return (
                "=== ГРУППИРОВКА ===\n"
                "Колонка для группировки по темам не обнаружена автоматически. "
                "Сгруппируй данные по смыслу самостоятельно (по продуктам/темам/проблемам)."
            )
        lines = [
            "=== ГРУППИРОВКА ПО ТЕМАМ ===",
            f"Колонка группировки: '{self.column_name}'",
            f"Уникальных значений (= потенциальных тем): {self.group_count}",
            "",
            "Распределение сигналов по темам (отсортировано по частоте):",
        ]
        # Сортируем по убыванию частоты
        sorted_counts = sorted(
            self.value_counts.items(), key=lambda x: x[1], reverse=True
        )
        for value, count in sorted_counts:
            lines.append(f"  • {value}: {count} сигналов")
        return "\n".join(lines)


@dataclass
class DataContext:
    """Полный контекст данных для LLM."""
    source_path: str
    row_count: int
    column_count: int
    columns: List[ColumnProfile]
    sample_rows: List[Dict[str, Any]]
    summary_text: str
    grouping: Optional[GroupingInfo] = None
    # Сколько слайдов хочет пользователь (None = пусть LLM решит сам)
    requested_slide_count: Optional[int] = None

    def to_prompt_text(self) -> str:
        """Сериализация в компактный текст для системного промпта."""
        parts = [
            f"=== ИСТОЧНИК: {self.source_path} ===",
            f"Строк: {self.row_count}, колонок: {self.column_count}",
            "",
            "=== СТРУКТУРА ДАННЫХ (лист 'Сырье') ===",
        ]
        for col in self.columns:
            parts.append(col.to_prompt_text())

        parts.append("")
        parts.append("=== СЭМПЛ ДАННЫХ (первые строки) ===")
        for i, row in enumerate(self.sample_rows[:MAX_SAMPLE_ROWS], 1):
            row_repr = ", ".join(f"{k}={v!r}" for k, v in row.items() if v is not None)
            parts.append(f"  [{i}] {row_repr}")

        # Информация о группировке
        if self.grouping:
            parts.append("")
            parts.append(self.grouping.to_prompt_text())

        # Указание по количеству слайдов
        parts.append("")
        parts.append("=== КОЛИЧЕСТВО ТЕМ (TOPIC-СЛАЙДОВ) ===")
        if self.requested_slide_count:
            n = self.requested_slide_count
            parts.append(
                f"⚠️ КРИТИЧНО: массив \"topics\" должен содержать РОВНО {n} "
                f"элементов. Не {n - 1}, не {n + 1}, а именно {n}.\n"
                f"Отбери {n} самых важных тем по числу сигналов/упоминаний. "
                f"Перед возвратом JSON ПОСЧИТАЙ элементы в \"topics\" — их должно "
                f"быть ровно {n}. Если получилось больше — удали наименее значимые. "
                f"Если меньше — добавь следующие по важности из данных."
            )
        elif self.grouping and self.grouping.column_name:
            parts.append(
                f"Пользователь не указал число слайдов. Реши сам по важности: "
                f"всего {self.grouping.group_count} уникальных тем по колонке "
                f"'{self.grouping.column_name}'. Если их много (>10), отбери "
                f"самые значимые по числу сигналов. Если мало — раскрой все."
            )
        else:
            parts.append(
                "Пользователь не указал число слайдов. Сгруппируй данные по смыслу "
                "и сделай столько тем, сколько действительно есть в данных (обычно 3-8)."
            )

        parts.append("")
        parts.append("=== СУММАРИЗАЦИЯ ПРОБЛЕМ (лист 'Суммаризация') ===")
        parts.append(self.summary_text.strip())
        return "\n".join(parts)


# --------------------------------------------------------------------------- #
# Чтение и профилирование
# --------------------------------------------------------------------------- #

class ExcelReader:
    """Читает xlsx и собирает DataContext."""

    def __init__(
        self,
        path: str | Path,
        grouping_column: Optional[str] = None,
        requested_slide_count: Optional[int] = None,
    ):
        """
        :param path: путь к xlsx
        :param grouping_column: явное имя колонки группировки. Если None —
            ищем автоматически по GROUPING_COLUMN_CANDIDATES.
        :param requested_slide_count: сколько topic-слайдов хочет пользователь.
            Если None — решает LLM.
        """
        self.path = Path(path)
        self.grouping_column = grouping_column
        self.requested_slide_count = requested_slide_count
        if not self.path.exists():
            raise FileNotFoundError(f"Файл не найден: {self.path}")
        if self.path.suffix.lower() not in {".xlsx", ".xlsm"}:
            raise ValueError(f"Ожидается xlsx-файл, получено: {self.path.suffix}")

    def read(self) -> DataContext:
        """Главная точка входа."""
        xls = pd.ExcelFile(self.path)
        available = set(xls.sheet_names)

        missing = {RAW_SHEET, SUMMARY_SHEET} - available
        if missing:
            raise ValueError(
                f"В файле {self.path.name} отсутствуют листы: {missing}. "
                f"Найдены: {sorted(available)}"
            )

        raw_df = pd.read_excel(xls, sheet_name=RAW_SHEET)
        summary_df = pd.read_excel(xls, sheet_name=SUMMARY_SHEET, header=None)

        grouping = self._detect_grouping(raw_df)

        return DataContext(
            source_path=str(self.path),
            row_count=len(raw_df),
            column_count=len(raw_df.columns),
            columns=[self._profile_column(raw_df[c]) for c in raw_df.columns],
            sample_rows=self._extract_samples(raw_df),
            summary_text=self._extract_summary(summary_df),
            grouping=grouping,
            requested_slide_count=self.requested_slide_count,
        )

    def _detect_grouping(self, df: pd.DataFrame) -> GroupingInfo:
        """
        Определяет колонку группировки.

        Приоритет:
        1. Явно заданная пользователем (self.grouping_column).
        2. Автопоиск по кандидатам (частичное совпадение, регистронезависимо).
        3. Эвристика: первая текстовая колонка с разумным числом уникальных
           значений (2..MAX_UNIQUE_FOR_CATEGORY).
        """
        col = self._resolve_grouping_column(df)
        if col is None:
            return GroupingInfo(column_name=None, detected_automatically=True)

        series = df[col].dropna().astype(str)
        counts = series.value_counts()
        return GroupingInfo(
            column_name=col,
            unique_values=[str(v) for v in counts.index.tolist()],
            value_counts={str(k): int(v) for k, v in counts.items()},
            detected_automatically=(self.grouping_column is None),
        )

    def _resolve_grouping_column(self, df: pd.DataFrame) -> Optional[str]:
        """Находит имя колонки группировки или None."""
        columns = list(df.columns)
        columns_lower = {str(c).lower().strip(): c for c in columns}

        # 1. Явно задано пользователем
        if self.grouping_column:
            # точное совпадение
            if self.grouping_column in columns:
                return self.grouping_column
            # регистронезависимое
            key = self.grouping_column.lower().strip()
            if key in columns_lower:
                return columns_lower[key]
            # частичное
            for low, orig in columns_lower.items():
                if key in low or low in key:
                    return orig
            # не нашли явно заданную — продолжаем автопоиск

        # 2. Автопоиск по кандидатам
        for candidate in GROUPING_COLUMN_CANDIDATES:
            # точное
            if candidate in columns_lower:
                return columns_lower[candidate]
            # частичное совпадение имени колонки с кандидатом
            for low, orig in columns_lower.items():
                if candidate in low:
                    return orig

        # 3. Эвристика: первая текстовая колонка с 2..MAX уникальных
        for c in columns:
            series = df[c].dropna()
            if series.empty:
                continue
            if pd.api.types.is_numeric_dtype(series):
                continue
            n_unique = series.nunique()
            if 2 <= n_unique <= MAX_UNIQUE_FOR_CATEGORY:
                return c

        return None

    # --- private ------------------------------------------------------------

    @staticmethod
    def _profile_column(series: pd.Series) -> ColumnProfile:
        """Собирает профиль колонки в зависимости от её типа."""
        non_null = series.dropna()
        profile = ColumnProfile(
            name=str(series.name),
            dtype=str(series.dtype),
            non_null_count=int(non_null.count()),
            null_count=int(series.isna().sum()),
            unique_count=int(non_null.nunique()),
        )

        if pd.api.types.is_numeric_dtype(series) and not non_null.empty:
            profile.min_value = float(non_null.min())
            profile.max_value = float(non_null.max())
            profile.mean_value = float(non_null.mean())

        # Сэмпл уникальных значений
        unique_vals = non_null.unique()[:5]
        profile.sample_values = [_to_jsonable(v) for v in unique_vals]

        # Топ-значения для категориальных
        if profile.unique_count <= MAX_UNIQUE_FOR_CATEGORY and profile.unique_count > 0:
            counts = non_null.value_counts().head(7)
            profile.top_values = {str(k): int(v) for k, v in counts.items()}

        return profile

    @staticmethod
    def _extract_samples(df: pd.DataFrame) -> List[Dict[str, Any]]:
        """Берёт первые строки и приводит значения к JSON-совместимым."""
        sample = df.head(MAX_SAMPLE_ROWS).fillna("")
        return [
            {str(k): _to_jsonable(v) for k, v in row.items()}
            for row in sample.to_dict(orient="records")
        ]

    @staticmethod
    def _extract_summary(df: pd.DataFrame) -> str:
        """
        Лист 'Суммаризация' может быть в любом формате: одна ячейка,
        несколько строк, табличка. Склеиваем всё непустое.
        """
        chunks: List[str] = []
        for _, row in df.iterrows():
            for cell in row:
                if pd.notna(cell):
                    text = str(cell).strip()
                    if text:
                        chunks.append(text)
        return "\n".join(chunks)


def _to_jsonable(value: Any) -> Any:
    """Преобразует numpy/pandas-типы в нативные Python."""
    if pd.isna(value):
        return None
    if hasattr(value, "item"):  # numpy scalars
        return value.item()
    if isinstance(value, (pd.Timestamp,)):
        return value.isoformat()
    return value
