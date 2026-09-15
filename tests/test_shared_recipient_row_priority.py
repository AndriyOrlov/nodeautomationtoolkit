"""Кілька рядків словника ведуть на одного адресата — хто дає йому назву.

У таблиці трапляється пара рядків з ОДНАКОВОЮ колонкою B: підпорядкована
установа, яка розсилається на свій обласний ТЦК, і сам ОТЦК. Раніше назву для
розрахунку діставав той рядок, у якого колонка C просто не порожня, — і в
«Розрахунку розсилки» зʼявлялася підпорядкована установа замість ОТЦК, хоча в
наказі названо саме ОТЦК.

Назви вигадані, але формою повторюють справжні (разом з опискою «Центер»,
через яку рядок ОТЦК не впізнавався як ТЦК).
"""

import sys
from pathlib import Path

import pytest
from openpyxl import Workbook

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from nodeautomationtoolkit.builtin_nodes.recipient_mapping import (  # noqa: E402
    map_military_units,
    read_recipient_mapping,
)

SUBORDINATE = "1 регіональний центр соціального супроводу"
OBLAST_TCK = "Тестівський Обласний Територіальний Центер Комплектування та Соціальної Підтримки"
SHARED_KEY = "Тестівський ОТЦК та СП"

ITEM = (
    "12. Майора ТЕСТЕНКА Андрія, старшого офіцера відділення обліку мобілізаційної роботи "
    "Тестівського обласного територіального центру комплектування та соціальної підтримки."
)
ORDER = "§ 1\n\nПРИЗНАЧИТИ:\n\n" + ITEM + "\n"


def _mapping_path(tmp_path: Path) -> str:
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["Найменування", "Шифр", "Скорочення", "Корпус", "Кому", "Куди"])
    # Підпорядкований рядок іде ПЕРШИМ — саме його скорочення й перемагало.
    sheet.append([SUBORDINATE, SHARED_KEY, "1 рег. центр с.с.", "",
                  f"Начальнику {SHARED_KEY}", "м. Тестове"])
    sheet.append([OBLAST_TCK, SHARED_KEY, "", "", f"Начальнику {SHARED_KEY}", "м. Тестове"])
    path = tmp_path / "shared.xlsx"
    workbook.save(path)
    return str(path)


def _routes(tmp_path: Path) -> dict:
    mapping = read_recipient_mapping(path=_mapping_path(tmp_path)).get("mapping", {})
    return map_military_units(text=ORDER, mapping=mapping)


def test_calculation_names_the_recipient_not_the_subordinate(tmp_path):
    rows = _routes(tmp_path)["units_table"].rows

    assert rows, "розрахунок порожній"
    assert rows[0][0] == SHARED_KEY, rows[0]
    assert "рег. центр" not in rows[0][0], rows[0]


def test_extract_still_goes_to_the_shared_recipient(tmp_path):
    assert list(_routes(tmp_path)["unit_paragraphs"].keys()) == [SHARED_KEY]


def test_matched_entry_is_the_oblast_row(tmp_path):
    audit = _routes(tmp_path)["routing_audit"]

    row = next(row for row in audit if row.get("label") == "Пункт 12.")
    assert OBLAST_TCK in row.get("matched_entries", ""), row


@pytest.mark.parametrize("reverse", [False, True])
@pytest.mark.parametrize("named", [SUBORDINATE, OBLAST_TCK])
@pytest.mark.parametrize("in_heading", [False, True])
def test_shared_cipher_uses_the_matched_search_row(named, reverse, in_heading):
    entries = [
        (SUBORDINATE, {"cipher": SHARED_KEY, "recipient_to": "SUB_TO", "destination_where": "SUB_WHERE"}),
        (OBLAST_TCK, {"cipher": SHARED_KEY, "recipient_to": "OWN_TO", "destination_where": "OWN_WHERE"}),
    ]
    mapping = dict(reversed(entries) if reverse else entries)
    text = (
        f"§ 1\nПо {named} ПРИЗНАЧИТИ:\n1. Майора ТЕСТЕНКА офіцером."
        if in_heading else f"§ 1\nПРИЗНАЧИТИ:\n1. Майора ТЕСТЕНКА офіцером {named}."
    )
    routes = map_military_units(text=text, mapping=mapping)
    row = routes["unit_paragraphs"][SHARED_KEY]
    expected = mapping[named]
    assert row["recipient_to"] == expected["recipient_to"]
    assert row["destination_where"] == expected["destination_where"]


@pytest.mark.parametrize("same_cipher", [False, True])
def test_identical_addresses_do_not_merge_search_names(same_cipher):
    mapping = {
        "901 окремий тестовий полк": {"cipher": "А9901", "recipient_to": "SAME_TO", "destination_where": "SAME_WHERE"},
        "902 окремий тестовий батальйон": {"cipher": "А9901" if same_cipher else "А9902", "recipient_to": "SAME_TO", "destination_where": "SAME_WHERE"},
    }
    order = "§ 1\nПРИЗНАЧИТИ:\n" + "\n".join(
        f"{i}. Майора ТЕСТЕНКА офіцером {name}." for i, name in enumerate(mapping, 1)
    )
    result = map_military_units(text=order, mapping=mapping)
    assert result["unmatched_items"] == []
    assert [r["matched_entries"] for r in result["routing_audit"]] == list(mapping)
    routes = result["unit_paragraphs"]
    assert len(routes) == (1 if same_cipher else 2)
    assert sum(len(r["items"]) for r in routes.values()) == 2
    assert all(r["recipient_to"] == "SAME_TO" and r["destination_where"] == "SAME_WHERE" for r in routes.values())


def test_explicit_item_row_wins_over_heading_row_for_shared_cipher():
    heading_name = "911 окремий тестовий полк"
    item_name = "912 окремий тестовий батальйон"
    mapping = {
        heading_name: {
            "cipher": "А9911", "abbreviation": "911 отп",
            "recipient_to": "HEADING_TO", "destination_where": "HEADING_WHERE",
        },
        item_name: {
            "cipher": "А9911", "abbreviation": "912 отб",
            "recipient_to": "ITEM_TO", "destination_where": "ITEM_WHERE",
        },
    }
    order = (
        f"§ 1\nПо {heading_name} ПРИЗНАЧИТИ:\n"
        f"1. Майора ТЕСТЕНКА перевести до {item_name}."
    )

    routes = map_military_units(text=order, mapping=mapping)
    assert len(routes["unit_paragraphs"]) == 1
    entry = next(iter(routes["unit_paragraphs"].values()))

    assert entry["recipient_to"] == "ITEM_TO"
    assert entry["destination_where"] == "ITEM_WHERE"
