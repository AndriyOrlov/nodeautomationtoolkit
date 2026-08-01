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


@node(
    name="Розбити текст на слова",
    category="Текст",
    description=(
        "Розбиває текст і після виконання додає окремий підписаний вихід для "
        "кожного слова. Порти зберігаються разом зі сценарієм."
    ),
    type_id="builtin.text.split_words_dynamic",
    outputs={"слова": "List"},
    dynamic_outputs=True,
)
def split_words_dynamic(text: str = "", maximum_outputs: int = 30) -> dict:
    words = re.findall(r"[^\W_]+(?:[-'’][^\W_]+)*", text, re.UNICODE)
    limit = max(1, min(int(maximum_outputs), 200))
    result: dict[str, object] = {"слова": words}
    duplicates: dict[str, int] = {}
    for index, word in enumerate(words[:limit], start=1):
        base = word[:42]
        duplicates[base] = duplicates.get(base, 0) + 1
        suffix = f" #{duplicates[base]}" if duplicates[base] > 1 else ""
        result[f"{index}. {base}{suffix}"] = word
    return result
