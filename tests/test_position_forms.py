"""Довідники відмінків посад і звань: повнота й охайність форм.

Довідники збирає `scripts/order_index/build_position_forms.py` з Excel-генератора
й зведеного переліку. Тести стережуть саме те, що ламалося: порожні відмінки,
форма від чужої (довшої) посади, перевернуті рядки з каталогу, з яких модуль
відмінювання вчив хибні форми слів.
"""

from __future__ import annotations

import csv
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from nodeautomationtoolkit.personnel import declension  # noqa: E402

DICTIONARIES = (
    Path(__file__).resolve().parents[1]
    / "src"
    / "nodeautomationtoolkit"
    / "personnel"
    / "dictionaries"
)
POSITION_CASES = ("Родовий", "Знахідний", "Давальний", "Орудний")


def _rows(name: str) -> list[dict[str, str]]:
    path = DICTIONARIES / name
    return list(csv.DictReader(path.read_text(encoding="utf-8-sig").splitlines(), delimiter=";"))


def test_every_position_has_all_four_cases():
    rows = _rows("positions.csv")
    assert len(rows) > 1200
    missing = [
        row["Називний"]
        for row in rows
        if not all(row.get(case) for case in POSITION_CASES)
    ]
    assert not missing, f"без усіх відмінків: {missing[:5]}"


def test_forms_keep_the_number_of_words():
    """Відмінок не змінює кількості слів: інакше це форма чужої посади."""
    wrong = []
    for row in _rows("positions.csv"):
        words = len(row["Називний"].split())
        for case in POSITION_CASES:
            if len(row[case].split()) != words:
                wrong.append(f"{row['Називний']} → {row[case]}")
    assert not wrong, wrong[:5]


def test_ranks_have_the_forms_the_orders_need():
    rows = {row["Називний"]: row for row in _rows("ranks.csv")}
    assert len(rows) > 40
    for rank in ("солдат", "сержант", "капітан", "майор", "полковник"):
        assert rows[rank]["Знахідний"] and rows[rank]["Давальний"]


@pytest.mark.parametrize(
    "position, case, expected",
    [
        ("командир відділення", "З", "командира відділення"),
        ("командир відділення", "Д", "командиру відділення"),
        ("командир відділення", "О", "командиром відділення"),
        # Назви на -а мають власний знахідний, а не однаковий із родовим.
        ("старшина батальйону", "З", "старшину батальйону"),
        ("медична сестра", "З", "медичну сестру"),
        ("медична сестра", "Д", "медичній сестрі"),
        # Фраза, якої в довіднику немає: форми виводяться послівно.
        ("старша медична сестра", "З", "старшу медичну сестру"),
        ("старша медична сестра", "Д", "старшій медичній сестрі"),
        ("начальник управління особового складу штабу", "З",
         "начальника управління особового складу штабу"),
        ("асистент фармацевта", "Д", "асистенту фармацевта"),
    ],
)
def test_positions_decline_as_expected(position, case, expected):
    result = declension.decline_position(position, case)
    assert result.text == expected


@pytest.mark.parametrize(
    "rank, case, expected",
    [
        ("капітан", "З", "капітана"),
        ("капітан", "Д", "капітану"),
        ("молодший сержант", "Д", "молодшому сержанту"),
        ("старший лейтенант юстиції", "Д", "старшому лейтенанту юстиції"),
    ],
)
def test_ranks_decline_as_expected(rank, case, expected):
    assert declension.decline_rank(rank, case).text == expected


def test_catalogue_artifacts_are_out_of_the_dictionary():
    """«сестра медична» вчила модуль хибній формі «старша» → «старші»."""
    names = {row["Називний"].casefold() for row in _rows("positions.csv")}
    for artifact in ("сестра медична", "старша медична", "головна медична"):
        assert artifact not in names
