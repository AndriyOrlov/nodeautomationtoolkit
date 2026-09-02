"""Тести поабзацного шифрування назв частин.

`cipher_unit_names` застосовується до ОКРЕМОГО абзацу, скопійованого з наказу
разом із форматуванням, тому вона не має права змінювати структуру рядків —
інакше абзаци перестануть відповідати оригіналу.
"""

import pytest

from nodeautomationtoolkit.builtin_nodes.message_order import (
    cipher_unit_names,
    find_content_start_line,
    generate_decision_order,
)

MAPPING = {
    "55 окремий батальйон безпілотних систем": {
        "open_name": "55 окремий батальйон безпілотних систем",
        "cipher": "А0000",
        "corps": "",
        "abbreviation": "55 обпс",
    },
}


def test_ciphers_single_paragraph():
    text, count, rows = cipher_unit_names(
        "Призначити до 55 окремого батальйону безпілотних систем.", MAPPING
    )
    assert "військової частини А0000" in text
    assert count >= 1
    assert rows


def test_does_not_change_line_count():
    """Головна вимога: кількість рядків лишається незмінною."""
    source = "\n".join(
        [
            "1. Перший пункт 55 окремого батальйону безпілотних систем.",
            "",
            "2. Другий пункт цього самого батальйону.",
            "",
            "3. Третій пункт.",
        ]
    )
    result, _, _ = cipher_unit_names(source, MAPPING)
    assert len(result.splitlines()) == len(source.splitlines())


def test_does_not_add_blank_lines_before_items():
    source = "1. Перший пункт.\n2. Другий пункт."
    result, _, _ = cipher_unit_names(source, MAPPING)
    assert result == source


def test_handles_reference_phrases():
    result, _, _ = cipher_unit_names("Перевести з цього самого батальйону.", MAPPING)
    assert "цієї самої військової частини" in result


def test_empty_text_is_safe():
    assert cipher_unit_names("", MAPPING) == ("", 0, [])
    assert cipher_unit_names(None, MAPPING) == ("", 0, [])


def test_no_mapping_keeps_text_readable():
    source = "Призначити до 55 окремого батальйону безпілотних систем."
    result, count, _ = cipher_unit_names(source, {})
    assert count == 0
    assert "55" in result


@pytest.mark.parametrize(
    "text, expected_line",
    [
        ("НАКАЗ\nкомандира\n\n§ 1\nВідповідно до", 3),
        ("Шапка\n1. Пункт наказу", 1),
        ("Шапка\nнижчепойменованих ЗВІЛЬНИТИ з посад", 1),
        ("Немає жодної ознаки", 0),
    ],
)
def test_find_content_start_line(text, expected_line):
    assert find_content_start_line(text) == expected_line


def test_decision_order_still_normalises_blank_lines():
    """Повний генератор і далі розставляє порожні рядки — на відміну від ядра."""
    source = "§ 1\n1. Перший пункт.\n2. Другий пункт."
    result = generate_decision_order(text=source, mapping=MAPPING)["decision_text"]
    assert len(result.splitlines()) > len(source.splitlines())
