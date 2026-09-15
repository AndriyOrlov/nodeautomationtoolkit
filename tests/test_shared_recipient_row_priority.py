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
