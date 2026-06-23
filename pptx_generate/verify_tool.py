"""LangChain-инструмент: проверка классификации моделью + «кол-во от модели»."""
from __future__ import annotations

import logging
from typing import Any, Optional, Type

from langchain_core.callbacks import CallbackManagerForToolRun
from langchain_core.tools import BaseTool
from pydantic import BaseModel, Field

from .classify_verify import verify_and_count, format_verify_report

logger = logging.getLogger(__name__)


class VerifyClassificationInput(BaseModel):
    xlsx_path: str = Field(
        description="Абсолютный путь к xlsx с листами 'исх' и 'динамика'."
    )
    output_path: Optional[str] = Field(
        default=None,
        description="Куда сохранить результат. По умолчанию перезаписывает исходный файл.",
    )


class VerifyClassificationTool(BaseTool):
    name: str = "verify_classification"
    description: str = (
        "Проверяет МОДЕЛЬЮ построчно лист 'исх': соответствует ли 'Текст из источника "
        "обратной связи' присвоенному 'Кластер сигналов 2 уровня'. Пишет True/False в "
        "новую колонку 'Соответствие модели'. Затем считает число True по каждой теме "
        "(классы берутся из листа 'динамика', колонка 'тема') и записывает это число в "
        "лист 'динамика' в новую колонку 'кол-во от модели'. Возвращает путь к "
        "обновлённому файлу и отчёт. Вызывай ДО генерации дайджеста, когда нужна "
        "проверка классификации."
    )
    args_schema: Type[BaseModel] = VerifyClassificationInput
    llm: Any = None

    def _run(self, xlsx_path: str, output_path: Optional[str] = None,
             run_manager: Optional[CallbackManagerForToolRun] = None) -> str:
        if self.llm is None:
            return ("ОШИБКА: модель не передана. Создавай инструмент как "
                    "VerifyClassificationTool(llm=<твоя GigaChat-модель>).")
        try:
            out, rep = verify_and_count(xlsx_path, self.llm, output_path=output_path)
            return f"Готово. Обновлённый файл: {out}\n\n" + format_verify_report(rep)
        except Exception as e:
            logger.exception("Ошибка проверки классификации")
            return f"ОШИБКА: {type(e).__name__}: {e}"

    async def _arun(self, *args: Any, **kwargs: Any) -> str:
        return self._run(*args, **kwargs)
