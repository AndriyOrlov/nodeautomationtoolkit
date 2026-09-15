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


@pytest.mark.parametrize("reverse", [False, True])
def test_numbered_column_a_name_wins_over_longer_generic_name(reverse):
    entries = [
        ("902 центр підготовки", {"cipher": "А0902", "abbreviation": "902 цп"}),
        ("окремий центр підготовки", {"cipher": "А0002", "abbreviation": "оцп"}),
    ]
    mapping = dict(reversed(entries) if reverse else entries)

    result, _, _ = cipher_unit_names(
        "офіцера 902 окремого центру підготовки", mapping
    )

    assert result == "офіцера військової частини А0902"


def test_column_c_abbreviation_is_not_a_search_key_for_ciphering():
    mapping = {
        "902 окремий тестовий полк": {
            "cipher": "А0902",
            "abbreviation": "особливе скорочення",
        }
    }
    source = "офіцера особливого скорочення"

    result, count, rows = cipher_unit_names(source, mapping)

    assert result == source
    assert count == 0
    assert rows == []


def test_generic_table_name_does_not_cipher_anaphoric_reference():
    mapping = {"центр": {"cipher": "А0903", "abbreviation": "ц"}}

    result, _, _ = cipher_unit_names("перевести до цього самого центру", mapping)

    assert result == "перевести до цієї самої військової частини"


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
