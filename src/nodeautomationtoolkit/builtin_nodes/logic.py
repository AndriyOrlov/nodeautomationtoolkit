from __future__ import annotations

from typing import Any

from nodeautomationtoolkit.core.definition import node


@node(name="Умова", category="Логіка", type_id="builtin.logic.if_else")
def if_else(condition: bool, when_true: Any, when_false: Any) -> Any:
    return when_true if condition else when_false


@node(name="Фільтр списку", category="Логіка", type_id="builtin.logic.filter_contains")
def filter_contains(items: list, text: str, ignore_case: bool = True) -> list:
    needle = text.casefold() if ignore_case else text
    result = []
    for item in items:
        candidate = str(item).casefold() if ignore_case else str(item)
        if needle in candidate:
            result.append(item)
    return result


@node(name="Без дублікатів", category="Логіка", type_id="builtin.logic.unique")
def unique(items: list) -> list:
    result = []
    for item in items:
        if item not in result:
            result.append(item)
    return result

