"""Тести для модуля генерації повідомлень (message_order)."""

import pytest

from nodeautomationtoolkit.builtin_nodes.message_order import (
    _apply_case_to_closed_text,
    _detect_grammatical_case,
    generate_decision_order,
)

_BPS_MAPPING = {
    "55 окремий батальйон безпілотних систем": {
        "open_name": "55 окремий батальйон безпілотних систем",
        "cipher": "А0000",
        "corps": "",
        "abbreviation": "55 обпс",
    },
    "72 окрема механізована бригада": {
        "open_name": "72 окрема механізована бригада",
        "cipher": "А1111",
        "corps": "",
        "abbreviation": "72 омбр",
    },
}


@pytest.mark.parametrize(
    "phrase, expected",
    [
        ("55 окремий батальйон безпілотних систем", "Н"),
        ("55 окремого батальйону безпілотних систем", "Р"),
        ("55 окремому батальйону безпілотних систем", "Д"),
        ("55 окремим батальйоном безпілотних систем", "О"),
        ("72 окрема механізована бригада", "Н"),
        ("72 окремої механізованої бригади", "Р"),
        ("72 окремій механізованій бригаді", "Д"),
        ("72 окрему механізовану бригаду", "З"),
        ("72 окремою механізованою бригадою", "О"),
    ],
)
def test_detect_grammatical_case(phrase, expected):
    assert _detect_grammatical_case(phrase) == expected


def test_detect_case_defaults_to_genitive_when_unsure():
    """Без впевнених ознак лишається родовий — поведінка не гіршає."""
    assert _detect_grammatical_case("") == "Р"
    assert _detect_grammatical_case("123") == "Р"


def test_detect_case_is_case_insensitive():
    assert _detect_grammatical_case("55 ОКРЕМИМ БАТАЛЬЙОНОМ") == "О"


@pytest.mark.parametrize(
    "case_label, expected",
    [
        ("Н", "військова частина А0000"),
        ("Р", "військової частини А0000"),
        ("Д", "військовій частині А0000"),
        ("З", "військову частину А0000"),
        ("О", "військовою частиною А0000"),
    ],
)
def test_apply_case_to_closed_text(case_label, expected):
    assert _apply_case_to_closed_text("військової частини А0000", case_label) == expected


def test_apply_case_keeps_corps_suffix_in_genitive():
    """Корпус є родовим означенням і має лишатися незмінним."""
    result = _apply_case_to_closed_text(
        "військової частини А0000 військової частини А1111", "О"
    )
    assert result == "військовою частиною А0000 військової частини А1111"


def test_apply_case_does_not_touch_tck_names():
    assert _apply_case_to_closed_text("Львівський ОТЦК та СП", "О") == "Львівський ОТЦК та СП"


@pytest.mark.parametrize(
    "sentence, expected",
    [
        ("1. 55 окремий батальйон безпілотних систем сформовано.", "військова частина А0000"),
        ("1. Призначити до 55 окремого батальйону безпілотних систем.", "військової частини А0000"),
        ("1. У 55 окремому батальйоні безпілотних систем збори.", "військовій частині А0000"),
        ("1. Разом з 55 окремим батальйоном безпілотних систем.", "військовою частиною А0000"),
        ("1. Направити у 72 окрему механізовану бригаду.", "військову частину А1111"),
        ("1. Керувати 72 окремою механізованою бригадою.", "військовою частиною А1111"),
    ],
)
def test_decision_order_substitutes_in_correct_case(sentence, expected):
    res = generate_decision_order(text=sentence, mapping=_BPS_MAPPING)
    assert expected in res["decision_text"]


def test_generate_decision_order_basic():
    mapping = {
        "72 окрема механізована бригада": {
            "cipher": "А2167",
            "open_name": "72 окрема механізована бригада",
            "corps": "",
        },
    }
    text = (
        "НАКАЗ\n\n"
        "1. Капітана призначити до 72 окремої механізованої бригади.\n"
        "2. Солдата звільнити.\n\n"
        "Командир військової частини А0000"
    )
    res = generate_decision_order(text=text, mapping=mapping)
    decision = res["decision_text"]
    assert "А2167" in decision
    assert res["replaced_count"] >= 1
    # Check blank line rules: 1 before item, 2 before signer
    assert "\n\nКомандир" in decision or "\n\n\nКомандир" in decision


def test_generate_decision_order_custom_rules():
    text = "1. Старшого сержанта направити до ВЧ 1234."
    rules = "ВЧ 1234 -> військової частини А9999"
    res = generate_decision_order(text=text, rules=rules)
    assert "військової частини А9999" in res["decision_text"]
