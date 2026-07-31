from nodeautomationtoolkit import node


@node(
    name="Додати префікс",
    category="Приклади",
    description="Приклад користувацької Python-ноди.",
)
def add_prefix(text: str, prefix: str = "ВИТЯГ: ") -> str:
    return prefix + text

