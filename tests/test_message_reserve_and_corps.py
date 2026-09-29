"""Повідомлення: відмінок у батальйонах резерву й частина з корпусом у назві.

1. «55 батальйону резерву» — родовий двох іменників чоловічого роду. Правило
   узгодження бачило два слова на «-у» й ставило знахідний, як у «окрему
   механізовану бригаду»: виходило «командира роти військову частину А0055».
   Знахідний на «-у/-ю» тепер — лише з головним словом жіночого роду.
2. Частина, у стовпці A якої записано підпорядкування («… бригада 11
   армійського корпусу»), вважалась корпусом, бо в назві є «корпус…», і група
   частин у {{кому_список}} лишалась без неї.

Дані вигадані.
"""

import importlib.util
import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from nodeautomationtoolkit.builtin_nodes.message_order import (  # noqa: E402
    _detect_grammatical_case,
    cipher_unit_names,
)
from nodeautomationtoolkit.builtin_nodes.recipient_mapping import map_military_units  # noqa: E402


@pytest.mark.parametrize(
    "matched, case_label",
    [
        ("55 батальйону резерву", "Р"),
        ("55 БАТАЛЬЙОНУ РЕЗЕРВУ", "Р"),
        ("10 полку зв'язку", "Р"),
        ("55 батальйоні резерву", "Д"),
        ("55 батальйоном резерву", "О"),
        ("5 окрему механізовану бригаду", "З"),
        ("у 3 тестову базу", "З"),
        ("5 окрема механізована бригада", "Н"),
    ],
)
def test_case_of_reserve_battalion_and_feminine_accusative(matched, case_label):
    assert _detect_grammatical_case(matched) == case_label


def test_reserve_battalion_is_closed_in_genitive():
    mapping = {"55 батальйон резерву": {"open_name": "55 батальйон резерву", "cipher": "А0055", "corps": ""}}
    text, _count, _report = cipher_unit_names(
        text="командира роти 55 батальйону резерву, КОМАНДИРОМ ВЗВОДУ", mapping=mapping
    )
    assert "командира роти військової частини А0055" in text


def _load_generator():
    spec = importlib.util.spec_from_file_location(
        "generate_extracts_reserve_corps_tests", PROJECT_ROOT / "generate_extracts.py"
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _row(name, cipher, abbreviation, corps=""):
    return {
        "open_name": name,
        "cipher": cipher,
        "abbreviation": abbreviation,
        "corps": corps,
        "recipient_to": f"Командиру військової частини {cipher}",
        "destination_where": "м. Тестове",
    }


ORDER = """НАКАЗ

§ 1

Відповідно до пункту 1 Положення нижчепойменованих військовослужбовців 101 окремої механізованої бригади 11 армійського корпусу ЗВІЛЬНИТИ з займаних посад і ПРИЗНАЧИТИ ДО 202 ОКРЕМОЇ МЕХАНІЗОВАНОЇ БРИГАДИ 22 АРМІЙСЬКОГО КОРПУСУ:

1. Капітана ПЕРШОГО Петра, командира роти, КОМАНДИРОМ БАТАЛЬЙОНУ.
"""


@pytest.mark.parametrize(
    "first_unit, second_unit",
    [
        ("101 окрема механізована бригада", "202 окрема механізована бригада"),
        (
            "101 окрема механізована бригада 11 армійського корпусу",
            "202 окрема механізована бригада 22 армійського корпусу",
        ),
    ],
)
def test_heading_path_gives_both_corps_and_both_units(first_unit, second_unit):
    rows = [
        _row("11 армійський корпус", "А1100", "11 АК"),
        _row("22 армійський корпус", "А2200", "22 АК"),
        _row(first_unit, "А0101", "101 омбр", "11 АК"),
        _row(second_unit, "А0202", "202 омбр", "22 АК"),
    ]
    mapping = {row["open_name"]: row for row in rows}
    routes = map_military_units(text=ORDER, mapping=mapping)
    groups = _load_generator().build_message_recipient_groups(mapping, routes)

    assert [recipient.split()[-1] for recipient in groups["corps"]] == ["А1100", "А2200"]
    assert [recipient.split()[-1] for recipient in groups["units"]] == ["А0101", "А0202"]
