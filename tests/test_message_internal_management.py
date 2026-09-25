"""Повідомлення: внутрішнє переміщення в управлінні не переноситься у зміст.

Пункт, де «управління» є і в посаді «звідки» (малими), і в посаді «КУДИ»
(ВЕЛИКИМИ), — переміщення всередині управління: повідомляти про нього нікого.
Витяги до управління за таким пунктом лишаються як були.

Усі назви та реквізити вигадані (публічний репозиторій — без реальних даних).
"""

import importlib.util
import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from nodeautomationtoolkit.builtin_nodes.recipient_mapping import (
    is_internal_management_move,
    map_military_units,
)


def _load_generator():
    spec = importlib.util.spec_from_file_location(
        "generate_extracts_internal_management_tests", PROJECT_ROOT / "generate_extracts.py"
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


NODE = "555 інформаційно-телекомунікаційний вузол"
MAPPING = {
    NODE: {
        "open_name": NODE,
        "cipher": "А0555",
        "abbreviation": "555 ітв",
        "corps": "",
        "recipient_to": "Командиру військової частини А0555",
        "destination_where": "м. Тестове",
    },
}

INTERNAL = (
    "1. Майора ТЕСТЕНКА Олега Васильовича, старшого офіцера відділу планування "
    "управління зв'язку - СТАРШИМ ОФІЦЕРОМ ВІДДІЛУ ОРГАНІЗАЦІЇ УПРАВЛІННЯ ЗВ'ЯЗКУ."
)

ORDER_LINES = [
    "НАКАЗ",
    "§ 1",
    "Відповідно до пункту 1 Положення нижчепойменованих офіцерів ЗВІЛЬНИТИ з "
    "займаних посад і ПРИЗНАЧИТИ:",
    "",
    INTERNAL,
    "1985 р.н., освіта: вища, академія державного управління, у ЗС - з 2005.",
    "",
    "2. Капітана ПРИКЛАДЕНКА Івана Петровича, офіцера 555 інформаційно-"
    "телекомунікаційного вузла - СТАРШИМ ОФІЦЕРОМ 555 ІНФОРМАЦІЙНО-"
    "ТЕЛЕКОМУНІКАЦІЙНОГО ВУЗЛА.",
    "1990 р.н.",
    "",
    "§ 2",
    "Відповідно до пункту 2 Положення нижчепойменованих офіцерів ЗВІЛЬНИТИ з "
    "займаних посад і ПРИЗНАЧИТИ:",
    "",
    "3. Майора ЗРАЗКОВА Петра Івановича, начальника відділу управління персоналу - "
    "ЗАСТУПНИКОМ НАЧАЛЬНИКА УПРАВЛІННЯ ПЕРСОНАЛУ.",
    "1980 р.н.",
    "",
    "4. Капітана ВИГАДАНКА Олега Петровича, офіцера відділу управління персоналу - "
    "СТАРШИМ ОФІЦЕРОМ ВІДДІЛУ КАДРІВ.",
    "1991 р.н.",
]
ORDER = "\n".join(ORDER_LINES)


@pytest.fixture(scope="module")
def routes() -> dict:
    return map_military_units(text=ORDER, mapping=MAPPING)


def test_internal_move_needs_management_in_both_parts():
    assert is_internal_management_move([INTERNAL])
    # Одна посада з двома «управліннями» — це не переміщення.
    assert not is_internal_management_move(
        ["1. Підполковником ТЕСТЕНКОМ Олегом Васильовичем, старшим офіцером відділу "
         "планування розвідувального управління штабу управління."]
    )
    # З управління — не в управління.
    assert not is_internal_management_move(
        ["4. Капітана ВИГАДАНКА, офіцера відділу управління персоналу - "
         "СТАРШИМ ОФІЦЕРОМ ВІДДІЛУ КАДРІВ."]
    )
    # «Пункт управління» — це командний пункт, а не управління.
    assert not is_internal_management_move(
        ["1. Майора ТЕСТЕНКА, офіцера управління - ОФІЦЕРОМ ПУНКТУ УПРАВЛІННЯ ЗВ'ЯЗКУ."]
    )


def test_biography_is_not_counted():
    assert not is_internal_management_move(
        [
            "1. Майора ТЕСТЕНКА, офіцера відділу кадрів - СТАРШИМ ОФІЦЕРОМ УПРАВЛІННЯ ЗВ'ЯЗКУ.",
            "1985 р.н., освіта: вища, академія державного управління.",
        ]
    )


def test_routing_marks_internal_moves_but_keeps_management_extracts(routes):
    flags = {item["label"]: item["internal_management_move"] for item in routes["item_spans"]}
    assert flags == {"Пункт 1.": True, "Пункт 2.": False, "Пункт 3.": True, "Пункт 4.": False}

    # Витяги до управління — як і раніше, за всіма пунктами управління.
    assert [
        data["items"][0]["label"] for data in routes["management_paragraphs"].values()
    ] == ["Пункт 1.", "Пункт 3.", "Пункт 4."]
    audit = {row["label"]: row["applied_rules"] for row in routes["routing_audit"]}
    assert "внутрішнє переміщення в управлінні" in audit["Пункт 1."]
    assert "внутрішнє переміщення в управлінні" not in audit["Пункт 4."]


def test_message_skips_internal_items_and_keeps_shared_heading(routes):
    generator = _load_generator()
    lines, labels = generator.message_skipped_item_lines(routes)
    assert labels == ["Пункт 1.", "Пункт 3."]

    kept = [line for index, line in enumerate(ORDER_LINES) if index not in lines]
    text = "\n".join(kept)
    assert "ТЕСТЕНКА" not in text and "академія державного управління" not in text
    assert "ЗРАЗКОВА" not in text
    assert "ПРИКЛАДЕНКА" in text and "ВИГАДАНКА" in text
    # Під обома § лишились інші пункти — шапки на місці.
    assert "§ 1" in text and "§ 2" in text


def test_heading_without_remaining_items_is_dropped():
    order_lines = ORDER_LINES[:10] + [
        "§ 2",
        "Відповідно до пункту 2 Положення нижчепойменованих офіцерів ЗВІЛЬНИТИ з "
        "займаних посад і ПРИЗНАЧИТИ:",
        "",
        "3. Майора ЗРАЗКОВА Петра Івановича, начальника відділу управління персоналу - "
        "ЗАСТУПНИКОМ НАЧАЛЬНИКА УПРАВЛІННЯ ПЕРСОНАЛУ.",
        "1980 р.н.",
    ]
    routes = map_military_units(text="\n".join(order_lines), mapping=MAPPING)
    lines, labels = _load_generator().message_skipped_item_lines(routes)
    assert labels == ["Пункт 1.", "Пункт 3."]
    kept = [line for index, line in enumerate(order_lines) if index not in lines]
    assert "§ 1" in kept
    assert "§ 2" not in kept
    assert not any("пункту 2 Положення" in line for line in kept)


def test_nothing_skipped_without_internal_moves():
    # Усе малими — частини «КУДИ» немає, отже й переміщення в управління.
    text = "\n".join(ORDER_LINES[:10]).replace(INTERNAL, INTERNAL.lower())
    routes = map_military_units(text=text, mapping=MAPPING)
    assert _load_generator().message_skipped_item_lines(routes) == (set(), [])
