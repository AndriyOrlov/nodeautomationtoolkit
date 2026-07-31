from pathlib import Path

import pytest

from nodeautomationtoolkit.core.node_draft import (
    NodeDraft,
    install_node_draft,
    review_node_code,
)

SAFE_NODE = '''from nodeautomationtoolkit import node

@node(name="Подвоїти", category="Тест")
def double(value: int) -> int:
    return value * 2
'''


def test_accepts_small_pure_node():
    review = review_node_code(SAFE_NODE)
    assert review.installable
    assert review.errors == []
    assert review.permissions == set()


@pytest.mark.parametrize(
    "code, expected",
    [
        ("import subprocess\n" + SAFE_NODE, "Заборонений імпорт"),
        (SAFE_NODE + "\neval('1 + 1')\n", "Заборонений динамічний виклик"),
        ("def plain():\n    return 1\n", "рівно одну функцію"),
    ],
)
def test_rejects_dangerous_or_invalid_plugins(code: str, expected: str):
    review = review_node_code(code)
    assert not review.installable
    assert expected in "\n".join(review.errors)


def test_installs_valid_draft(tmp_path: Path):
    draft = NodeDraft(
        name="Подвоїти",
        function_name="double",
        code=SAFE_NODE,
    )
    target = install_node_draft(draft, tmp_path)
    assert target.name == "generated_double.py"
    assert target.read_text(encoding="utf-8") == SAFE_NODE
