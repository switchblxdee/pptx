"""
Ремонт «почти валидного» JSON от слабых LLM.

GigaChat (особенно Pro) часто возвращает JSON с типичными дефектами:
- markdown-обёртки ```json ... ```
- преамбулы «Вот ваш дайджест:» / постамбулы «Готово!»
- одинарные кавычки вместо двойных
- трейлинг-запятые перед } и ]
- обрезанный хвост (ответ упёрся в лимит токенов)
- «умные» кавычки «ёлочки» и “fancy quotes”
- Python-литералы True/False/None вместо true/false/null
- комментарии // и /* */

Этот модуль чинит максимум дефектов БЕЗ обращения к LLM, чтобы
не тратить лишний раунд ретрая. Стратегия многоуровневая: пробуем
самый дешёвый способ, при неудаче — следующий по агрессивности.
"""
from __future__ import annotations

import json
import re
from typing import Any, Optional


_FENCE_RE = re.compile(r"```(?:json|JSON)?\s*(.*?)```", re.DOTALL)
_LINE_COMMENT_RE = re.compile(r"//[^\n\r]*")
_BLOCK_COMMENT_RE = re.compile(r"/\*.*?\*/", re.DOTALL)
_TRAILING_COMMA_RE = re.compile(r",(\s*[}\]])")

# Reasoning-теги, которыми thinking-модели (DeepSeek V4, GLM, QwQ и др.)
# оборачивают свои размышления перед финальным ответом.
_REASONING_TAGS_RE = re.compile(
    r"<(think|thinking|reasoning|thought|reflection)>.*?</\1>",
    re.DOTALL | re.IGNORECASE,
)


def _strip_reasoning(text: str) -> str:
    """
    Убирает блоки рассуждений thinking-моделей.

    DeepSeek V4, GLM и другие reasoning-модели могут вернуть
    <think>...</think> перед JSON. Если такой блок есть — вырезаем его,
    чтобы он не мешал извлечению JSON.
    """
    return _REASONING_TAGS_RE.sub("", text)


def extract_json_text(raw: str) -> str:
    """
    Достаёт JSON-фрагмент из текста ответа.

    Порядок: убрать reasoning-теги → fenced-блок → баланс { }.
    """
    raw = _strip_reasoning(raw).strip()

    fence = _FENCE_RE.search(raw)
    if fence:
        candidate = fence.group(1).strip()
        if candidate:
            return candidate

    # Берём от первой { до последней }
    first = raw.find("{")
    last = raw.rfind("}")
    if first != -1 and last != -1 and last > first:
        return raw[first:last + 1]

    return raw


def _normalize_quotes(text: str) -> str:
    """Заменяет «умные» кавычки на обычные ASCII."""
    replacements = {
        "\u201c": '"', "\u201d": '"',   # " "
        "\u201e": '"', "\u201f": '"',
        "\u2018": "'", "\u2019": "'",   # ' '
        # Ёлочки оставляем ВНУТРИ строк (это часть цитат),
        # их трогать нельзя — поэтому не заменяем « ».
    }
    for bad, good in replacements.items():
        text = text.replace(bad, good)
    return text


def _strip_comments(text: str) -> str:
    """Убирает // и /* */ комментарии, которые иногда вставляет LLM."""
    text = _BLOCK_COMMENT_RE.sub("", text)
    # Строчные комментарии убираем осторожно: только если // не внутри
    # строки. Простая эвристика — убираем // до конца строки, но это
    # может задеть "https://". Поэтому проверяем, что перед // нет ":".
    lines = []
    for line in text.split("\n"):
        # ищем // не предварённое ":" или '"'
        m = re.search(r'(?<![:"\'/])//', line)
        if m:
            line = line[:m.start()]
        lines.append(line)
    return "\n".join(lines)


def _fix_python_literals(text: str) -> str:
    """True/False/None → true/false/null (вне строк)."""
    # Заменяем только как отдельные токены (с границами слова),
    # чтобы не задеть слова внутри строк типа "TrueType".
    text = re.sub(r'\bTrue\b', 'true', text)
    text = re.sub(r'\bFalse\b', 'false', text)
    text = re.sub(r'\bNone\b', 'null', text)
    return text


def _remove_trailing_commas(text: str) -> str:
    """Убирает запятые перед } и ]."""
    prev = None
    while prev != text:
        prev = text
        text = _TRAILING_COMMA_RE.sub(r"\1", text)
    return text


def _close_unbalanced(text: str) -> str:
    """
    Чинит обрезанный хвост: если ответ упёрся в лимит токенов,
    закрываем недостающие } и ].

    Считаем баланс скобок ВНЕ строк и дописываем недостающие закрытия.
    """
    in_string = False
    escape = False
    stack = []
    for ch in text:
        if escape:
            escape = False
            continue
        if ch == "\\":
            escape = True
            continue
        if ch == '"':
            in_string = not in_string
            continue
        if in_string:
            continue
        if ch in "{[":
            stack.append(ch)
        elif ch == "}":
            if stack and stack[-1] == "{":
                stack.pop()
        elif ch == "]":
            if stack and stack[-1] == "[":
                stack.pop()

    # Если строка не закрыта — закрываем кавычку
    suffix = ""
    if in_string:
        suffix += '"'

    # Закрываем оставшиеся скобки в обратном порядке
    for opener in reversed(stack):
        suffix += "}" if opener == "{" else "]"

    return text + suffix


def _fix_single_quotes(text: str) -> str:
    """
    Заменяет одинарные кавычки-делимитеры на двойные.

    Это рискованная операция: апострофы внутри текста («д'Артаньян»,
    «it's») трогать нельзя. Поэтому действуем консервативно — меняем
    только кавычки, стоящие в позиции делимитера JSON: после { , : [
    и перед } , : ] (с учётом пробелов).
    """
    # ' в роли открывающей кавычки ключа/значения: после { [ , :
    text = re.sub(r'([{\[,:]\s*)\'', r'\1"', text)
    # ' в роли закрывающей: перед } ] , :
    text = re.sub(r'\'(\s*[}\],:])', r'"\1', text)
    return text


def repair_json(raw: str) -> dict[str, Any]:
    """
    Главная функция: пытается распарсить JSON, последовательно применяя
    всё более агрессивные починки.

    Возвращает dict или бросает ValueError, если ничего не помогло.
    """
    text = extract_json_text(raw)

    # Уровень 0: вдруг уже валидный
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    # Уровень 1: базовая чистка (кавычки, комментарии, литералы, запятые)
    cleaned = text
    cleaned = _normalize_quotes(cleaned)
    cleaned = _strip_comments(cleaned)
    cleaned = _fix_python_literals(cleaned)
    cleaned = _remove_trailing_commas(cleaned)
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        pass

    # Уровень 1.5: + аккуратная замена одинарных кавычек
    single_fixed = _fix_single_quotes(cleaned)
    single_fixed = _remove_trailing_commas(single_fixed)
    try:
        return json.loads(single_fixed)
    except json.JSONDecodeError:
        pass

    # Уровень 2: + закрытие оборванного хвоста
    repaired = _close_unbalanced(single_fixed)
    repaired = _remove_trailing_commas(repaired)
    try:
        return json.loads(repaired)
    except json.JSONDecodeError:
        pass

    # Уровень 3: последняя попытка — обрезать до последней закрытой
    # верхнеуровневой структуры и закрыть. Иногда хвост безнадёжно битый.
    truncated = _truncate_to_last_valid(single_fixed)
    if truncated:
        try:
            return json.loads(truncated)
        except json.JSONDecodeError:
            pass

    raise ValueError(
        "Не удалось починить JSON даже после всех уровней ремонта. "
        f"Начало ответа:\n{raw[:300]}"
    )


def _truncate_to_last_valid(text: str) -> Optional[str]:
    """
    Обрезает текст до последней позиции, на которой JSON балансируется,
    и закрывает структуру. Помогает при сильно битом хвосте.
    """
    in_string = False
    escape = False
    stack = []
    last_safe = None  # позиция после последнего «хорошего» закрытия элемента

    for i, ch in enumerate(text):
        if escape:
            escape = False
            continue
        if ch == "\\":
            escape = True
            continue
        if ch == '"':
            in_string = not in_string
            continue
        if in_string:
            continue
        if ch in "{[":
            stack.append(ch)
        elif ch in "}]":
            if stack:
                stack.pop()
            # После закрытия элемента запоминаем безопасную точку
            if len(stack) <= 1:
                last_safe = i + 1

    if last_safe is None:
        return None

    candidate = text[:last_safe]
    return _close_unbalanced(candidate)
