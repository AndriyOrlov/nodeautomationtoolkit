"""Шапка розділу називає КІЛЬКА частин — пункти під нею йдуть до КОЖНОЇ.

Було: пошук частин у шапці зупинявся на першій знайденій (`break`), тож усі
пункти під шапкою «По військовій частині А1111 та військовій частині А2222:»
потрапляли лише в ОДИН витяг. У який саме — залежало від порядку рядків у
словнику, тому на реальному наказі це виглядало як випадковість.

Напрямок КУДИ в тій самій шапці збирав усі збіги й до цієї правки — тут
перевіряються обидва напрямки.
"""

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from nodeautomationtoolkit.builtin_nodes.recipient_mapping import map_military_units

FIRST = "4 окремий тестовий полк"
SECOND = "5 окрема тестова бригада"
OWN = "3 окремий тестовий загін"


def _mapping() -> dict:
    return {
        name: {
            "open_name": name,
            "cipher": cipher,
            "abbreviation": abbreviation,
            "corps": "",
            "recipient_to": f"Командиру військової частини {cipher}",
            "destination_where": "м. Тестове",
        }
        for name, cipher, abbreviation in (
            (OWN, "А0003", "3 отз"),
            (FIRST, "А0004", "4 отп"),
            (SECOND, "А0005", "5 отбр"),
        )
    }


ORDER = f"""§ 1

Відповідно до пункту 1 Положення ПРИЗНАЧИТИ:

1. Капітана ПЕРШОГО Петра, командира роти {OWN}, НАЧАЛЬНИКОМ СЛУЖБИ {OWN}.

По {FIRST} та {SECOND}:

2. Лейтенанта ДРУГОГО Дмитра, командира взводу, КОМАНДИРОМ РОТИ.

3. Сержанта ТРЕТЬОГО Тараса, старшину роти, СТАРШИНОЮ БАТАЛЬЙОНУ.
"""


def _labels(routes: dict, key_part: str) -> list[str]:
    for key, data in routes["unit_paragraphs"].items():
        if key_part in key:
            return [item.get("label") for item in data.get("items", [])]
    return []


def test_items_under_heading_go_to_every_unit_named_in_it():
    routes = map_military_units(text=ORDER, mapping=_mapping())

    for cipher in ("А0004", "А0005"):
        labels = _labels(routes, cipher)
        assert "Пункт 2." in labels, f"{cipher}: пункт 2 не потрапив у витяг ({labels})"
        assert "Пункт 3." in labels, f"{cipher}: пункт 3 не потрапив у витяг ({labels})"


def test_item_with_its_own_unit_is_not_pulled_into_heading_units():
    """Пункт, який сам називає частину, не роздається всім частинам із шапки."""
    routes = map_military_units(text=ORDER, mapping=_mapping())

    assert _labels(routes, "А0003") == ["Пункт 1."]
    for cipher in ("А0004", "А0005"):
        assert "Пункт 1." not in _labels(routes, cipher)


def test_audit_shows_both_units_as_context():
    routes = map_military_units(text=ORDER, mapping=_mapping())

    row = next(row for row in routes["routing_audit"] if row.get("label") == "Пункт 2.")
    context = row.get("context_recipients", "")
    assert "А0004" in context and "А0005" in context, context
