"""Повний ланцюг підпорядкування в повідомленнях: бат → ВЧ → АК/ОК → ОК → КСВ.

ОК і КСВ повідомлення не отримують (03.11.2026, КСВ — тимчасово). Дані вигадані.
"""

from nodeautomationtoolkit.builtin_nodes.message_order import cipher_unit_names
from nodeautomationtoolkit.builtin_nodes.recipient_mapping import _format_full_closed_unit_text
from nodeautomationtoolkit.generator.messages import (
    build_message_recipient_groups,
    is_message_skipped_command,
)


def _row(name, cipher, abbr, superior="", to=""):
    return {
        "open_name": name,
        "cipher": cipher,
        "abbreviation": abbr,
        "corps": superior,
        "recipient_to": to,
    }


MAPPING = {
    "командування сухопутних військ": _row("командування сухопутних військ", "А5555", "КСВ"),
    "1 оперативне командування": _row("1 оперативне командування", "А4444", "ОК Тест", "КСВ"),
    "77 армійський корпус": _row("77 армійський корпус", "А3333", "77 АК", "ОК Тест", "Командиру"),
    "10 окрема механізована бригада": _row(
        "10 окрема механізована бригада", "А1010", "10 омбр", "77 АК", "Командиру"
    ),
    "20 окрема механізована бригада": _row(
        "20 окрема механізована бригада", "А2020", "20 омбр", "ОК Тест", "Командиру"
    ),
    "1 механізований батальйон 10 окремої механізованої бригади": _row(
        "1 механізований батальйон 10 окремої механізованої бригади", "А0101", "1 мб", "10 омбр"
    ),
}


def test_chain_goes_through_corps_up_to_ksv():
    assert _format_full_closed_unit_text(MAPPING["1 механізований батальйон 10 окремої механізованої бригади"], MAPPING, full_chain=True) == (
        "військової частини А0101 військової частини А1010 військової частини А3333 "
        "військової частини А4444 військової частини А5555"
    )


def test_unit_directly_under_ok():
    assert _format_full_closed_unit_text(MAPPING["20 окрема механізована бригада"], MAPPING, full_chain=True) == (
        "військової частини А2020 військової частини А4444 військової частини А5555"
    )


def test_extracts_chain_still_stops_at_corps():
    assert _format_full_closed_unit_text(MAPPING["10 окрема механізована бригада"], MAPPING) == (
        "військової частини А1010 військової частини А3333"
    )


def test_message_text_has_each_cipher_once():
    text = "1. Капітана ТЕСТЕНКА, офіцера 10 окремої механізованої бригади 77 армійського корпусу."
    result = cipher_unit_names(text, MAPPING)[0]
    assert result == (
        "1. Капітана ТЕСТЕНКА, офіцера військової частини А1010 військової частини А3333 "
        "військової частини А4444 військової частини А5555."
    )


def test_ok_and_ksv_are_not_recipients():
    assert is_message_skipped_command(MAPPING["1 оперативне командування"])
    assert is_message_skipped_command(MAPPING["командування сухопутних військ"])
    assert not is_message_skipped_command(MAPPING["77 армійський корпус"])

    routes = {"unit_paragraphs": {
        "А1010": {"open_name": "10 окрема механізована бригада", "unit_code": "А1010"},
        "А2020": {"open_name": "20 окрема механізована бригада", "unit_code": "А2020"},
        "А4444": {"open_name": "1 оперативне командування", "unit_code": "А4444"},
    }}
    groups = build_message_recipient_groups(MAPPING, routes)
    everything = " ".join(groups["corps"] + groups["units"])
    assert "А3333" in " ".join(groups["corps"])  # корпус — адресат
    assert "А1010" in everything and "А2020" in everything
    assert "А4444" not in everything and "А5555" not in everything
