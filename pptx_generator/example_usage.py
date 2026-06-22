"""
Полный пример использования: GigaChat + tool + xlsx → дайджест.pptx.

Запуск:
    pip install python-pptx pandas openpyxl pydantic lxml langchain-core langchain-gigachat
    export GIGACHAT_CREDENTIALS="ваш_ключ"
    python example_usage.py
"""
import os
from pathlib import Path

import pandas as pd

from pptx_generator import GenerateDigestTool


def create_demo_xlsx(path: Path) -> Path:
    """Создаёт демонстрационный xlsx с двумя нужными листами."""
    raw_df = pd.DataFrame({
        "Продукт": ["GigaCode CLI"] * 5 + ["SberF1"] * 4 + ["Lokit"] * 3,
        "Категория": [
            "скорость", "цензор", "секреты", "скорость", "цензор",
            "поиск", "релевантность", "поиск", "ui",
            "стабильность", "интеграция", "стабильность",
        ],
        "Цитата": [
            "Второй день после обеда все жутко висит, работать не возможно",
            "разработчик с фамилией Медведев — это блочит нам работу с ИИ",
            "При перезапуске варм под windows зашифрованные секреты слетают",
            "После полудня запросы идут по 5-10 секунд",
            "Цензор блокирует обычные технические тексты",
            "Когда индекс больше 10к страниц, поиск выдаёт результат через 30+ секунд",
            "После апдейта ищет совсем не то, что нужно",
            "Поиск не находит документы по точному заголовку",
            "Кнопки в новой версии плохо различимы",
            "Lokit падает на сложных кейсах с ошибкой 500",
            "Не интегрируется с нашим CI/CD пайплайном",
            "После часа работы вылетает с потерей контекста",
        ],
        "Источник": [
            "#Ai in Dev Community", "#GigaIDE Support", "#GigaIDE Support",
            "#Ai in Dev Community", "#Ai in Dev Community",
            "#SberF1 Support", "#SberF1 Support", "#Поиск-чат", "#SberF1 Support",
            "#Lokit Users", "#Lokit Users", "#Lokit Users",
        ],
    })

    summary_df = pd.DataFrame([
        ["Голос IT за период 21-27 мая 2026."],
        ["Проанализировано 847 сигналов от сотрудников через 4 источника."],
        [""],
        ["Главные темы:"],
        ["1. GigaCode CLI: 108 сигналов, 3 активных темы — замедление, цензор, секреты."],
        ["2. SberF1: 54 сигнала, 2 активных темы — поиск и релевантность."],
        ["3. Lokit: 31 сигнал, 2 активных темы — стабильность и интеграции."],
    ])

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        raw_df.to_excel(writer, sheet_name="Сырье", index=False)
        summary_df.to_excel(writer, sheet_name="Суммаризация", index=False, header=False)

    return path


def main():
    xlsx_path = Path("/tmp/digest_demo.xlsx")
    create_demo_xlsx(xlsx_path)
    print(f"Создан тестовый xlsx: {xlsx_path}")

    from langchain_gigachat import GigaChat

    llm = GigaChat(
        credentials=os.environ["GIGACHAT_CREDENTIALS"],
        model="GigaChat-Pro",
        verify_ssl_certs=False,
        scope="GIGACHAT_API_PERS",
        temperature=0.2,
    )

    tool = GenerateDigestTool(llm=llm)

    result_path = tool.invoke({
        "xlsx_path": str(xlsx_path),
        "style_prompt": (
            "Корпоративный дайджест 'Голос IT' для руководства Блока T. "
            "Спокойная бирюзово-персиковая палитра с пастельно-розовыми "
            "карточками тем. Период — еженедельный отчёт."
        ),
        "output_path": "/tmp/voice_it_digest.pptx",
    })
    print(f"Готово: {result_path}")


if __name__ == "__main__":
    main()
