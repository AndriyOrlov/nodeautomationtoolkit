from __future__ import annotations

import re

from nodeautomationtoolkit.core.definition import node


@node(name="Текст", category="Текст", type_id="builtin.text.value")
def text_value(value: str = "") -> str:
    return value


@node(name="Об'єднати текст", category="Текст", type_id="builtin.text.concat")
def concat(first: str, second: str, separator: str = "") -> str:
    return f"{first}{separator}{second}"


@node(name="Нормалізувати пробіли", category="Текст", type_id="builtin.text.normalize")
def normalize_whitespace(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


@node(name="Пошук за Regex", category="Текст", type_id="builtin.text.regex_find")
def regex_find(text: str, pattern: str, ignore_case: bool = True) -> list[str]:
    flags = re.IGNORECASE if ignore_case else 0
    return [match.group(0) for match in re.finditer(pattern, text, flags)]


@node(name="Замінити текст", category="Текст", type_id="builtin.text.replace")
def replace(text: str, old: str, new: str) -> str:
    return text.replace(old, new)

