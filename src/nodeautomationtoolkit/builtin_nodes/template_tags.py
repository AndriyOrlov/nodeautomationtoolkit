"""Спільні імена тегів для примірників, витягів і повідомлень."""


TAG_GROUPS = (
    ("номер_наказу", "номер"),
    ("дата_наказу", "дата"),
    ("примірник", "примірник_номер"),
    ("згідно_з_оригіналом", "засвідчення"),
    ("підписант_посада", "посада_підписанта"),
    ("підписант_звання", "звання_підписанта"),
    ("підписант_піб", "підписант_імя", "підписант_ім'я", "підписант_ім’я", "прізвище_підписанта"),
    ("засвідчувач_посада", "затверджувач_посада", "згідно_з_оригіналом_посада"),
    ("засвідчувач_звання", "затверджувач_звання", "згідно_з_оригіналом_звання"),
    ("засвідчувач_піб", "засвідчувач_імя", "засвідчувач_ім'я", "засвідчувач_ім’я",
     "затверджувач_піб", "затверджувач_імя", "затверджувач_ім'я", "затверджувач_ім’я", "згідно_з_оригіналом_піб"),
    ("засвідчувач", "затверджувач"),
)


def tag_aliases(tag: str) -> tuple[str, ...]:
    name = tag.strip("{}").casefold()
    for group in TAG_GROUPS:
        if name in group:
            return tuple("{{" + alias + "}}" for alias in group)
    return (tag,)


SIGNER_TAGS = ("{{підписант}}",) + tuple(
    "{{" + name + "}}" for group in TAG_GROUPS[4:7] for name in group
)


def expand_common_tags(values: dict[str, str]) -> dict[str, str]:
    """Додає синоніми, не змінюючи явних значень і форматів дати/номера."""
    result = dict(values)
    for tag, value in values.items():
        for alias in tag_aliases(tag):
            result.setdefault(alias, value)
    return result


def certifier_tags(position: str, rank: str, name: str) -> dict[str, str]:
    return expand_common_tags({
        "{{засвідчувач_посада}}": position,
        "{{засвідчувач_звання}}": rank,
        "{{засвідчувач_піб}}": name,
        "{{засвідчувач}}": name,
    })


def signer_tags(position: str, rank: str, name: str) -> dict[str, str]:
    """Однаковий блок підписанта для всіх трьох видів документів."""
    signature_line = " ".join(part for part in (rank, name) if part)
    block = "\r".join(part for part in (position, signature_line) if part)
    return expand_common_tags({
        "{{підписант_посада}}": position,
        "{{підписант_звання}}": rank,
        "{{підписант_піб}}": name,
        "{{підписант}}": block,
    })
