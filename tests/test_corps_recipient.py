"""Маршрутизація частин, підпорядкованих корпусу.

Витяг для такої частини йде НА КОРПУС: і сам витяг, і його «Кому»/«Куди»
мають належати корпусу, а не частині.
"""

from nodeautomationtoolkit.builtin_nodes.recipient_mapping import map_military_units

UNIT = "33 окремий полк зв'язку"
CORPS = "14 армійський корпус"

MAPPING = {
    UNIT: {
        "open_name": UNIT,
        "cipher": "А1111",
        "abbreviation": "33 опз",
        "corps": "14 АК",
        "recipient_to": "Командиру в/ч А1111",
        "destination_where": "м. Львів",
    },
    CORPS: {
        "open_name": CORPS,
        "cipher": "А5555",
        "abbreviation": "14 АК",
        "corps": "",
        "recipient_to": "Командиру в/ч А5555",
        "destination_where": "м. Київ",
    },
}

ORDER = "НАКАЗ\n§ 1\n1. Призначити до 33 окремого полку зв'язку.\n"


def _routes(mapping=None, text=ORDER):
    return map_military_units(text=text, mapping=mapping or MAPPING).get("unit_paragraphs", {})


def test_extract_goes_to_the_corps_not_the_unit():
    keys = list(_routes().keys())
    assert len(keys) == 1
    assert "А5555" in keys[0]
    assert "А1111" not in keys[0]


def test_recipient_belongs_to_the_corps():
    """«Кому» бралося від частини, хоча витяг адресований корпусу."""
    entry = next(iter(_routes().values()))
    assert entry["recipient_to"] == "Командиру в/ч А5555"


def test_destination_belongs_to_the_corps():
    entry = next(iter(_routes().values()))
    assert entry["destination_where"] == "м. Київ"


def test_unit_without_corps_keeps_its_own_recipient():
    mapping = {UNIT: {**MAPPING[UNIT], "corps": ""}}
    entry = next(iter(_routes(mapping).values()))

    assert entry["recipient_to"] == "Командиру в/ч А1111"
    assert entry["destination_where"] == "м. Львів"


def test_corps_mentioned_directly_keeps_its_own_data():
    text = "НАКАЗ\n§ 1\n1. Призначити до 14 армійського корпусу.\n"
    entry = next(iter(_routes(text=text).values()))

    assert entry["recipient_to"] == "Командиру в/ч А5555"
