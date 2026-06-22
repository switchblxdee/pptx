# GigaChat Digest Tool — SberF1 edition

LangChain-инструмент для генерации корпоративных аналитических **дайджестов**
в формате PowerPoint на основе xlsx-данных и GigaChat. Поддерживает фирменное
оформление **SberF1**: брендовые фоны, палитру, шрифт и иконки из шаблона.

## Что это

Дайджест «Голос IT»: обложка + темы + аналитические слайды (executive summary,
паттерны, «на что обратить внимание», графики, closing). Стиль настраивается
**промптом**.

```
xlsx → ExcelReader → DataContext ──┐
                                   ├→ GigaChat → JSON → DigestSpec → DigestBuilder → .pptx
                  style_prompt ────┘
```

## Установка

```bash
pip install -r requirements.txt
# или как пакет (включая ассеты — фоны и иконки):
pip install .
```

## Быстрый старт

```python
from langchain_gigachat import GigaChat
from pptx_generator import GenerateDigestTool

llm = GigaChat(credentials="...", model="GigaChat-Pro", temperature=0.2)
tool = GenerateDigestTool(llm=llm)

result = tool.invoke({
    "xlsx_path": "/path/to/data.xlsx",
    "style_prompt": "Фирменный дайджест SberF1",   # включает бренд-стиль
    "output_path": "/path/to/digest.pptx",
})
```

## Управление стилем по промпту

Стиль резолвится послойно (`style_resolver.py` + `themes.py`):

1. **База** — `force_theme` → keyword-пресет → палитра LLM.
2. **Явные директивы** из текста накладываются поверх по ролям, причём роли
   фона и текста разбираются ОТДЕЛЬНО: «тёмный фон и тёмный текст» больше не
   превращает текст в белый.
3. **Замки** — заданные пользователем роли фиксируются (`locked_roles`) и
   авто-контраст их не трогает: рендерится ровно как просили.

Примеры:
```python
"Фирменный дайджест SberF1"               # бренд: фон+палитра+шрифт+иконки
"белый фон, чёрный текст, синий акцент"    # явные цвета по ролям
"тёмный фон и тёмный текст"                # уважается как есть (низкий контраст — выбор пользователя)
"светлая презентация с чёрным текстом"
```

## Бренд SberF1

Пресет `sberf1` (триггеры: «SberF1», «фирменный», «брендовый», «Блок технологии»,
«КомандаСбера», …) подставляет:

- **Фоны** — реальные подложки из шаблона (`assets/bg_cover.png`,
  `assets/bg_content.png`): титульный диагональный меш с логотипом и светлый
  контентный меш.
- **Палитру** — фирменные цвета (`0669E0`, `933EFF`, `FFB900`, `0FB880`,
  текст `111827`), тёмные KPI-карточки.
- **Шрифт** — `SB Sans Display`.
- **Иконки** — извлечённый из шаблона набор (`assets/icons/`), см. ниже.

## Иконки

`brand_icons.py` кладёт иконки из шаблона как PNG (`add_picture`) — это
универсально и **корректно рендерится в Р7-Офис / OnlyOffice**, PowerPoint и
LibreOffice. Каждая иконка в двух вариантах (тёмный/светлый) — выбор по фону
объекта (на тёмных карточках белые, на светлых тёмные).

Доступные хинты: `growth, chart, check, error, new, info, warning, security,
integration, calendar, tag, document, users, list, data, wifi, rocket`
(+ синонимы и авто-подбор по тексту подписи KPI).

Добавить ещё иконок из шаблона:
```bash
# контактный лист с номерами
python -m pptx_generator.tools.extract_brand_icons \
    --template SberF1.pptx --slide 6 --contact-sheet contact.png
# извлечь выбранные (правишь CHOSEN в скрипте или --map map.json)
python -m pptx_generator.tools.extract_brand_icons \
    --template SberF1.pptx --slide 6 --out pptx_generator/assets/icons
```

## Надёжность рендера

Болевые места OOXML, из-за которых раньше что-то не отображалось в PowerPoint/Р7,
закрыты:

- **Фон** ставится штатным API python-pptx (валидный `<p:bg>`), а не сырым XML
  не в тот элемент — иначе PowerPoint оставлял фон белым.
- **Графики** (column/bar) рисуются фигурами-прямоугольниками, а не нативным
  чарт-объектом — столбцы видны всегда.
- **Контраст** считается по WCAG для каждого объекта от его заливки
  (`contrast.py`): на тёмном — светлый текст, на светлом — тёмный.
- Премиум-приёмы обёрнуты в «красиво → fallback» (`safe.py`): при ошибке
  откат на простой вариант, без падения сборки.

## Структура

```
pptx_generator/
├── builder.py        — рендер .pptx (фон, KPI, графики, иконки, контраст)
├── tool.py           — LangChain BaseTool + резолв стиля
├── schemas.py        — DigestSpec и модели
├── excel_reader.py   — чтение xlsx
├── prompts.py        — системный промпт + few-shot
├── themes.py         — пресеты палитр + бренд SberF1
├── style_resolver.py — разбор стиля из промпта (роли + замки)
├── contrast.py       — WCAG-контраст
├── safe.py           — «красиво → fallback»
├── icons.py          — встроенные векторные иконки (fallback)
├── brand_icons.py    — иконки из шаблона SberF1 (PNG, R7-совместимо)
├── json_repair.py    — починка JSON от LLM
├── assets/           — bg_cover.png, bg_content.png, icons/*.png
└── tools/extract_brand_icons.py — извлечение иконок из шаблона
```

## Backward compatibility

`GeneratePresentationTool` / `GeneratePresentationInput` — алиасы
`GenerateDigestTool` / `GenerateDigestInput`.
