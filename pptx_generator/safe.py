"""
safe.py — паттерн «сначала красиво, если упало — как раньше».

Любой рискованный визуальный приём (премиум-градиент, glow-блобы, XML-патчинг,
тени) оборачиваем так, чтобы при ЛЮБОЙ ошибке рендера слайд всё равно собрался
в простом надёжном варианте, а не падал целиком.

Использование 1 — декоратор пары (fancy, fallback):

    @fancy_or_fallback
    def _bg(self, slide):
        ...рискованный премиум-фон...

    @_bg.fallback
    def _bg(self, slide):
        ...простой двухцветный градиент как сейчас...

Использование 2 — инлайн:

    with_fallback(
        lambda: self._fancy_background(slide),
        lambda: self._simple_background(slide),
        label="background",
    )
"""
from __future__ import annotations

import functools
import logging
from typing import Callable, Optional, TypeVar

log = logging.getLogger("pptx_generator.render")

T = TypeVar("T")


def with_fallback(
    fancy: Callable[[], T],
    fallback: Callable[[], T],
    label: str = "",
) -> T:
    """Пробует fancy(); при исключении логирует и выполняет fallback()."""
    try:
        return fancy()
    except Exception as e:  # noqa: BLE001 — намеренно ловим всё
        log.warning("fancy render failed%s: %s -> fallback",
                    f" [{label}]" if label else "", e)
        return fallback()


class fancy_or_fallback:  # noqa: N801 — это декоратор, имя в нижнем регистре намеренно
    """Декоратор: основной метод = «красиво», .fallback = «как сейчас».

    При исключении в основном методе вызывается fallback с теми же
    аргументами. Логируется warning, presentation не падает.
    """

    def __init__(self, primary: Callable):
        functools.update_wrapper(self, primary)
        self._primary = primary
        self._fallback: Optional[Callable] = None

    def fallback(self, fn: Callable) -> "fancy_or_fallback":
        self._fallback = fn
        return self

    def __set_name__(self, owner, name):
        self._name = name

    def __get__(self, obj, objtype=None):
        # поддержка методов: связываем self
        if obj is None:
            return self
        return functools.partial(self.__call__, obj)

    def __call__(self, *args, **kwargs):
        try:
            return self._primary(*args, **kwargs)
        except Exception as e:  # noqa: BLE001
            log.warning("fancy '%s' failed: %s -> fallback",
                        getattr(self, "_name", self._primary.__name__), e)
            if self._fallback is None:
                raise
            return self._fallback(*args, **kwargs)
