from __future__ import annotations

import re
from datetime import datetime
from typing import Any

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
    name="Замінити за Regex",
    category="Текст",
    description="Замінює знайдені збіги регулярного виразу на заданий рядок. Підтримує групи ($1, $2).",
    type_id="builtin.text.regex_replace",
    outputs={"result": "str", "replaced_count": "int", "summary": "str"},
)
def regex_replace(
    text: str = "",
    pattern: str = "",
    replacement: str = "",
    ignore_case: bool = False,
) -> dict:
    if not pattern:
        return {"result": text, "replaced_count": 0, "summary": "Шаблон порожній"}
    flags = re.IGNORECASE if ignore_case else 0
    repl = replacement.replace("$", "\\")
    result, count = re.subn(pattern, repl, text, flags=flags)
    return {"result": result, "replaced_count": count, "summary": f"Виконано замін: {count}"}


@node(
    name="Очистити текст",
    category="Текст",
    description="Прибирає зайві пробіли, подвійні переноси рядків, BOM-символи та нерозривні пробіли.",
    type_id="builtin.text.clean",
    outputs={"result": "str", "summary": "str"},
)
def clean_text(
    text: str = "",
    strip_lines: bool = True,
    remove_double_newlines: bool = True,
    normalize_spaces: bool = True,
) -> dict:
    result = text.replace("\u00a0", " ").replace("\ufeff", "").replace("\r\n", "\n").replace("\r", "\n")
    if strip_lines:
        result = "\n".join(line.strip() for line in result.splitlines())
    if normalize_spaces:
        result = re.sub(r"[ \t]+", " ", result)
    if remove_double_newlines:
        result = re.sub(r"\n{3,}", "\n\n", result)
    original_len = len(text)
    new_len = len(result)
    return {
        "result": result.strip(),
        "summary": f"Видалено {original_len - new_len} зайвих символів",
    }


@node(
    name="Поточна дата",
    category="Текст",
    description="Повертає поточну дату та час у обраному форматі. Наприклад: '%d %B %Y' → '06 серпня 2026'.",
    type_id="builtin.text.today_date",
    outputs={"date_str": "str", "day": "int", "month": "int", "year": "int"},
)
def today_date(format_str: str = "%d.%m.%Y") -> dict:
    now = datetime.now()
    try:
        date_str = now.strftime(format_str)
    except Exception:
        date_str = now.strftime("%d.%m.%Y")
    return {"date_str": date_str, "day": now.day, "month": now.month, "year": now.year}


@node(
    name="Шаблон тексту",
    category="Текст",
    description=(
        "Заповнює шаблон текстовими значеннями. Поля у шаблоні позначаються фігурними дужками: "
        "{ПІБ}, {дата}, {посада}. Значення передаються у форматі JSON."
    ),
    type_id="builtin.text.fill_template",
    outputs={"result": "str", "summary": "str"},
)
def fill_template(
    template: str = "",
    fields_json: str = "{}",
    strict: bool = False,
) -> dict:
    import json

    try:
        fields: dict[str, Any] = json.loads(fields_json) if fields_json.strip() else {}
    except json.JSONDecodeError as err:
        raise ValueError(f"Помилка JSON-полів: {err}") from err

    result = template
    filled = 0
    for key, value in fields.items():
        placeholder = f"{{{key}}}"
        if placeholder in result:
            result = result.replace(placeholder, str(value))
            filled += 1
        elif strict:
            raise ValueError(f"Поле '{key}' не знайдено у шаблоні")

    unfilled = re.findall(r"\{[^}]+\}", result)
    summary = f"Заповнено полів: {filled}"
    if unfilled:
        summary += f" · Незаповнені: {', '.join(unfilled[:5])}"
    return {"result": result, "summary": summary}


@node(
    name="Розбити на рядки",
    category="Текст",
    description="Розбиває багаторядковий текст на список рядків. Порожні рядки можна пропустити.",
    type_id="builtin.text.split_lines",
    outputs={"lines": "List", "count": "int"},
)
def split_lines(text: str = "", skip_empty: bool = True) -> dict:
    lines = text.splitlines()
    if skip_empty:
        lines = [line for line in lines if line.strip()]
    return {"lines": lines, "count": len(lines)}


@node(
    name="Форматування числа",
    category="Текст",
    description="Форматує ціле чи дробове число у рядок. Наприклад: 1234567 → '1 234 567'.",
    type_id="builtin.text.number_format",
)
def number_format(value: Any = 0, decimals: int = 0, thousands_sep: str = " ") -> str:
    try:
        num = float(str(value).replace(",", ".").replace(" ", ""))
        if decimals == 0:
            formatted = f"{int(round(num)):,}".replace(",", thousands_sep)
        else:
            formatted = f"{num:,.{decimals}f}".replace(",", thousands_sep)
        return formatted
    except (ValueError, TypeError):
        return str(value)


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
    words = re.findall(r"[^\W_]+(?:[-''][^\W_]+)*", text, re.UNICODE)
    limit = max(1, min(int(maximum_outputs), 200))
    result: dict[str, object] = {"слова": words}
    duplicates: dict[str, int] = {}
    for index, word in enumerate(words[:limit], start=1):
        base = word[:42]
        duplicates[base] = duplicates.get(base, 0) + 1
        suffix = f" #{duplicates[base]}" if duplicates[base] > 1 else ""
        result[f"{index}. {base}{suffix}"] = word
    return result

