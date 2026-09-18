"""Повна назва частини береться з таблиці відповідностей, як в Excel-генераторі.

Excel підтягував назву з мережевого файла; ми беремо ту саму таблицю, з якої
працюють витяги (стовпець A — відкрита назва, B — шифр, C — скорочення).
Назви частин у фікстурах вигадані.
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest
from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from nodeautomationtoolkit.order_generator.units import (  # noqa: E402
    find_unit,
    full_unit_name,
    load_units,
    unit_genitive,
)

TABLE = [
    ["Відкрите найменування", "Шифр", "Скорочення", "Корпус", "Кому", "Куди"],
    [
        "72 окрема механізована бригада оперативного командування «Захід» Сухопутних військ",
        "А1111",
        "72 омбр",
        "10 АК",
        "Командиру військової частини А1111",
        "м. Львів",
    ],
    [
        "14 окремий полк зв'язку Командування Сухопутних військ",
        "А2222",
        "14 опз",
        "",
        "Командиру військової частини А2222",
        "м. Київ",
    ],
    [
        "окрема рота охорони 72 окремої механізованої бригади",
        "А3333",
        "оро 72 омбр",
        "10 АК",
        "Командиру військової частини А3333",
        "м. Львів",
    ],
]


@pytest.fixture
def table(tmp_path) -> Path:
    workbook = Workbook()
    sheet = workbook.active
    for row in TABLE:
        sheet.append(row)
    path = tmp_path / "Частини.xlsx"
    workbook.save(path)
    return path


def test_table_is_read(table):
    rows = load_units(table)
    assert len(rows) == 3
    assert rows[0][1] == "А1111"


def test_found_by_cipher(table):
    match = find_unit("А1111", table)
    assert match.how == "шифр"
    assert match.open_name.startswith("72 окрема механізована бригада")


def test_found_by_cipher_written_as_a_unit(table):
    """«військова частина А2222» — теж шифр, просто з обгорткою."""
    match = find_unit("військова частина А 2222", table)
    assert match is not None and match.cipher == "А2222"


def test_found_by_abbreviation(table):
    match = find_unit("72 омбр", table)
    assert match.how == "скорочення" and match.cipher == "А1111"


def test_found_by_words_of_the_name(table):
    match = find_unit("полк зв'язку", table)
    assert match.how == "назва" and match.cipher == "А2222"


def test_shortest_name_wins(table):
    """«72 окрема механізована» є і в назві роти охорони — беремо саму бригаду."""
    match = find_unit("72 окрема механізована", table)
    assert match.cipher == "А1111"


def test_nothing_found_keeps_what_was_typed(table):
    name, how = full_unit_name("А9999", table)
    assert (name, how) == ("А9999", "")


def test_no_table_at_all_keeps_what_was_typed(tmp_path):
    name, how = full_unit_name("72 омбр", tmp_path / "немає.xlsx")
    assert (name, how) == ("72 омбр", "")


def test_full_name_in_genitive_for_the_heading(table):
    name, how = full_unit_name("А1111", table, "Р")
    assert how == "шифр"
    assert name.startswith("72 окремої механізованої бригади оперативного командування")


@pytest.mark.parametrize(
    "nominative, genitive",
    [
        ("72 окрема механізована бригада", "72 окремої механізованої бригади"),
        ("14 окремий полк зв'язку", "14 окремого полку зв'язку"),
        ("окремий батальйон охорони", "окремого батальйону охорони"),
        ("рота забезпечення", "роти забезпечення"),
        ("Головне управління персоналу", "Головного управління персоналу"),
        ("Командування Сухопутних військ", "Командування Сухопутних військ"),
        ("3 армійський корпус", "3 армійського корпусу"),
        ("військова частина А1234", "військової частини А1234"),
    ],
)
def test_unit_genitive(nominative, genitive):
    assert unit_genitive(nominative) == genitive
