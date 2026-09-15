"""Частина, підпорядкована корпусу: один витяг, але обидва відправники в повідомленні.

У наказі формулювання «158 окрема бригада підтримки 22 армійського корпусу»:
- ВИТЯГ робиться лише на корпус;
- у ПОВІДОМЛЕННІ в списку відправників мають бути і частина (зашифрована),
  і корпус.
"""

import importlib.util
import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from nodeautomationtoolkit.builtin_nodes.recipient_mapping import map_military_units


def _load_generator_module():
    spec = importlib.util.spec_from_file_location(
        "generate_extracts_corps_tests", PROJECT_ROOT / "generate_extracts.py"
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


generator = _load_generator_module()

UNIT = "158 окрема бригада підтримки"
CORPS = "22 армійський корпус"

ORDER = (
    "§ 1\n"
    "1. Призначити до 158 окремої бригади підтримки 22 армійського корпусу."
)


def _mapping(corps_column: str, corps_abbreviation: str) -> dict:
    """A — повна назва, B — шифр, C — скорочення, D — підпорядкування, E/F — кому/куди."""
    return {
        UNIT: {
            "open_name": UNIT,
            "cipher": "А1158",
            "abbreviation": "158 обрп",
            "corps": corps_column,
            "recipient_to": "Командиру 158 обрп",
            "destination_where": "м. Львів",
        },
        CORPS: {
            "open_name": CORPS,
            "cipher": "А2222",
            "abbreviation": corps_abbreviation,
            "corps": "",
            "recipient_to": "Командиру 22 АК",
            "destination_where": "м. Київ",
        },
    }


@pytest.mark.parametrize(
    "corps_column, corps_abbreviation",
    [
        ("22 АК", "22 АК"),
        (CORPS, "22 АК"),
        # Порожня колонка C у корпуса давала ДВА ключі й два витяги
        ("22 АК", ""),
        (CORPS, ""),
    ],
)
def test_only_one_extract_is_created_for_the_corps(corps_column, corps_abbreviation):
    routes = map_military_units(text=ORDER, mapping=_mapping(corps_column, corps_abbreviation))
    keys = list(routes["unit_paragraphs"].keys())

    assert len(keys) == 1, f"очікувався один витяг, отримано: {keys}"
    assert "А2222" in keys[0]
    assert "А1158" not in keys[0]


def test_extract_recipient_is_the_corps():
    routes = map_military_units(text=ORDER, mapping=_mapping(CORPS, "22 АК"))
    entry = next(iter(routes["unit_paragraphs"].values()))

    assert entry["recipient_to"] == "Командиру 22 АК"
    assert entry["destination_where"] == "м. Київ"


def test_message_lists_both_the_unit_and_its_corps():
    """У повідомленні відправниками є і частина, і корпус — на відміну від витягу."""
    mapping = _mapping(CORPS, "22 АК")
    routes = map_military_units(text=ORDER, mapping=mapping)

    recipients = generator.build_message_recipient_list(mapping, routes)

    assert any("А2222" in line for line in recipients), recipients
    assert any("А1158" in line for line in recipients), recipients


def test_unit_without_corps_still_gets_its_own_extract():
    mapping = _mapping("", "22 АК")
    routes = map_military_units(text=ORDER, mapping=mapping)

    keys = list(routes["unit_paragraphs"].keys())
    assert any("А1158" in key for key in keys)
