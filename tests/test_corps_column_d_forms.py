"""Стовпець D може посилатися на корпус ПО-РІЗНОМУ — витяг має бути ОДИН.

У таблиці стовпець D часто містить СКОРОЧЕННЯ корпусу («ОК Захід»), яке
збігається зі стовпцем C його рядка, а не зі стовпцем A і не з шаблоном
«N АК». Рядок корпусу тоді не знаходився: ключ будувався без шифру, і на один
корпус виходило ДВА витяги, причому у другому «Кому»/«Куди» лишалися від
підпорядкованої частини.
"""

import sys
from pathlib import Path

import pytest

SRC_PATH = Path(__file__).resolve().parents[1] / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from nodeautomationtoolkit.builtin_nodes.recipient_mapping import map_military_units

CORPS_NAME = "оперативне командування «Тест»"
CORPS_ABBR = "ОК Тест"
UNIT_NAME = "267 окрема механізована бригада"

ORDER = (
    "§ 1\n"
    "Відповідно до пункту 1 ЗВІЛЬНИТИ:\n"
    "1. Полковника ТЕСТЕНКА Андрія Андрійовича, командира роти, звільнити із "
    "займаної посади і ЗАРАХУВАТИ у розпорядження командувача військ "
    "оперативного командування «Тест». На час перебування у розпорядженні "
    "командувача військ оперативного командування «Тест» залишається на всіх "
    "видах забезпечення та у списках 267 окремої механізованої бригади.\n"
    "\nКомандир військової частини А0001\nполковник   Петро ТЕСТОВИЙ\n"
)


def _mapping(corps_column):
    return {
        CORPS_NAME: {
            "open_name": CORPS_NAME,
            "cipher": "А0777",
            "abbreviation": CORPS_ABBR,
            "corps": "",
            "recipient_to": "Командиру військової частини А0777",
            "destination_where": "м. Корпусне",
        },
        UNIT_NAME: {
            "open_name": UNIT_NAME,
            "cipher": "А0267",
            "abbreviation": "267 омбр",
            "corps": corps_column,
            "recipient_to": "Командиру військової частини А0267",
            "destination_where": "м. Бригадне",
        },
    }


@pytest.mark.parametrize(
    "corps_column",
    [CORPS_ABBR, CORPS_NAME],
    ids=["D=скорочення", "D=повна-назва"],
)
def test_one_extract_for_the_corps_however_column_d_names_it(corps_column):
    routes = map_military_units(text=ORDER, mapping=_mapping(corps_column))
    units = routes["unit_paragraphs"]

    assert len(units) == 1, f"очікувався один витяг, отримано: {sorted(units)}"
    entry = next(iter(units.values()))
    assert entry["recipient_to"] == "Командиру військової частини А0777"
    assert entry["destination_where"] == "м. Корпусне"
