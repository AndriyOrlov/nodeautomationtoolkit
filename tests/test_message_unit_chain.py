"""Ланцюг підпорядкованості в закритому змісті повідомлення (AGENT.md 9.5.4, 9.5.9).

Скарга користувача (16.09.2026): «5 батальйону 5 штурмової бригади 25 армійського
корпусу» — у повідомленні пропадала ланка. Ланцюг буває з 4 ланок (батальйон →
бригада → корпус → командування). Усі номери й шифри вигадані.
"""

import pytest

import generate_extracts as generator
from nodeautomationtoolkit.builtin_nodes.message_order import (
    _detect_grammatical_case,
    cipher_unit_names,
)


def _entry(open_name, cipher, abbreviation="", corps=""):
    return {
        "open_name": open_name,
        "cipher": cipher,
        "abbreviation": abbreviation,
        "corps": corps,
        "recipient_to": "",
    }


BRIGADE = _entry("5 штурмова бригада", "А2222", "5 шбр", "25 АК")
CORPS = _entry("25 армійський корпус", "А3333", "25 АК", "ОК Тест")
COMMAND = _entry("оперативне командування «Тест»", "А4444", "ОК Тест")
TAIL = "25 армійського корпусу оперативного командування «Тест» Сухопутних військ Збройних Сил України"
FULL_CHAIN = (
    "військової частини А1111 військової частини А2222 "
    "військової частини А3333 військової частини А4444"
)


def _mapping(*entries):
    return {entry["open_name"]: entry for entry in entries}


@pytest.mark.parametrize("battalion_d", ["25 АК", "5 шбр", ""])
def test_separate_battalion_row_gives_all_four_links(battalion_d):
    mapping = _mapping(_entry("5 батальйон", "А1111", "5 бат", battalion_d), BRIGADE, CORPS, COMMAND)

    result, _, _ = cipher_unit_names(f"командира взводу 5 батальйону 5 штурмової бригади {TAIL}", mapping)

    assert result == f"командира взводу {FULL_CHAIN}"


@pytest.mark.parametrize(
    "cell",
    ["5 батальйон 5 штурмової бригади", "5 батальйон 5 штурмової бригади 25 армійського корпусу"],
)
@pytest.mark.parametrize("battalion_d", ["25 АК", "5 шбр", ""])
def test_cell_with_two_units_keeps_the_nested_link(cell, battalion_d):
    """Клітинка з бригадою всередині забирала назву бригади — ланка зникала."""
    mapping = _mapping(_entry(cell, "А1111", "5 бат", battalion_d), BRIGADE, CORPS, COMMAND)
    problems = []

    result, _, _ = cipher_unit_names(
        f"командира взводу 5 батальйону 5 штурмової бригади {TAIL}", mapping, problems=problems
    )

    assert result == f"командира взводу {FULL_CHAIN}"
    assert problems == []


def test_nested_link_keeps_upper_case_of_destination():
    mapping = _mapping(_entry("5 батальйон 5 штурмової бригади", "А1111", "5 бат"), BRIGADE, CORPS)

    result, _, _ = cipher_unit_names("ПРИЗНАЧИТИ до 5 БАТАЛЬЙОНУ 5 ШТУРМОВОЇ БРИГАДИ", mapping)

    assert result == "ПРИЗНАЧИТИ до ВІЙСЬКОВОЇ ЧАСТИНИ А1111 ВІЙСЬКОВОЇ ЧАСТИНИ А2222 ВІЙСЬКОВОЇ ЧАСТИНИ А3333"


def test_cell_with_two_units_and_the_brigade_cipher_is_reported():
    """Шифр батальйону взяти нема звідки — не вгадуємо, а кажемо в журнал."""
    merged = _entry("5 батальйон 5 штурмової бригади", "А2222", "5 шбр")
    mapping = _mapping(merged, _entry("5 штурмова бригада", "А2222", "5 шбр"))
    problems = []

    result, _, _ = cipher_unit_names("командира взводу 5 батальйону 5 штурмової бригади", mapping, problems=problems)

    assert result == "командира взводу військової частини А2222"
    assert problems == [("merged_cell", "5 батальйон 5 штурмової бригади", "5 штурмова бригада")]
    [line] = generator.describe_cipher_problems(problems)
    assert "дві частини" in line and "5 штурмова бригада" in line


def test_unnumbered_row_inside_a_name_adds_no_link():
    """Рядок без номера («штурмова бригада») є в будь-якій назві — ланки не дає."""
    mapping = _mapping(BRIGADE, _entry("штурмова бригада", "А9999"))

    result, _, _ = cipher_unit_names("командира взводу 5 штурмової бригади", mapping)

    assert "А9999" not in result
    assert "А2222" in result


@pytest.mark.parametrize(
    ("text", "expected"),
    [
        ("служить у 5 батальйоні", "служить у військовій частині А1111"),
        ("разом із 5 батальйоном", "разом із військовою частиною А1111"),
        ("до 5 батальйону", "до військової частини А1111"),
    ],
)
def test_case_of_a_name_without_adjective(text, expected):
    mapping = _mapping(_entry("5 батальйон", "А1111", "5 бат"))

    result, _, _ = cipher_unit_names(text, mapping)

    assert result == expected


@pytest.mark.parametrize(
    ("matched", "case_label"),
    [
        ("5 батальйоні", "Д"),
        ("5 полком", "О"),
        ("5 госпіталем", "О"),
        ("5 бригаді", "Д"),
        ("5 бригаду", "З"),
        ("5 бригадою", "О"),
        ("5 батальйону", "Р"),   # родовий і давальний збігаються — лишається родовий
        ("5 бригади", "Р"),
        ("5 дивізії", "Р"),      # родовий і давальний збігаються — не вгадуємо
        ("5 окремому батальйоні", "Д"),  # прикметникова ознака, як і раніше
        ("5 окремого батальйону", "Р"),
    ],
)
def test_detect_case_by_unit_kind(matched, case_label):
    assert _detect_grammatical_case(matched) == case_label
