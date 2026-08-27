"""Маршрутизація за частиною, названою в шапці розділу §.

Це постійні регресійні тести для сценаріїв із
``scripts/e2e_extracts/header_routing_check.py``. Усі назви вигадані.
"""

import pytest

from nodeautomationtoolkit.builtin_nodes.recipient_mapping import map_military_units


def _entry(name: str, cipher: str, abbreviation: str, corps: str = "") -> dict:
    return {
        "open_name": name,
        "cipher": cipher,
        "abbreviation": abbreviation,
        "corps": corps,
        "recipient_to": f"Командиру військової частини {cipher}",
        "destination_where": "м. Тестове",
    }


REPAIR = "46 окремий ремонтно-відновлювальний полк"
SIGNAL = "66 полк зв'язку"
CORPS = "55 армійський корпус"
OTHER = "77 окремий тестовий загін"

MAPPING = {
    REPAIR: _entry(REPAIR, "А0046", "46 орвп"),
    SIGNAL: _entry(SIGNAL, "А0066", "66 пз", corps="55 АК"),
    CORPS: _entry(CORPS, "А0055", "55 АК"),
    OTHER: _entry(OTHER, "А0077", "77 отз"),
}


def _routes(body: str) -> dict:
    result = map_military_units(text=body, mapping=MAPPING)
    assert result["unmatched_items"] == []
    return result["unit_paragraphs"]


def test_unit_named_in_section_heading_applies_to_all_unnamed_items():
    routes = _routes(
        "§ 1\n"
        "Відповідно до пунктів 1 та 2 Положення нижчепойменованих осіб "
        "офіцерського складу 46 окремого ремонтно-відновлювального полку "
        "ЗВІЛЬНИТИ з займаних посад і ПРИЗНАЧИТИ до цього самого полку:\n"
        "1. Майора ТЕСТЕНКА призначити начальником служби.\n"
        "2. Капітана ПРИКЛАДЕНКА призначити начальником відділення."
    )

    assert set(routes) == {"46 орвп А0046"}
    assert [item["label"] for item in routes["46 орвп А0046"]["items"]] == [
        "Пункт 1.",
        "Пункт 2.",
    ]


def test_heading_source_and_item_destination_both_receive_extract():
    routes = _routes(
        "§ 2\n"
        "Відповідно до пункту 3 Положення нижчепойменованих осіб "
        "із 66 полку зв'язку 55 армійського корпусу ЗВІЛЬНИТИ з займаних "
        "посад і ПРИЗНАЧИТИ:\n"
        "3. Майора ЗРАЗКОВА призначити офіцером 77 окремого тестового загону."
    )

    assert set(routes) == {"55 АК А0055", "77 отз А0077"}


def test_heading_destination_and_item_source_both_receive_extract():
    routes = _routes(
        "§ 3\n"
        "Відповідно до пункту 4 Положення нижчепойменованих осіб "
        "ЗВІЛЬНИТИ з займаних посад і ПРИЗНАЧИТИ до 66 полку зв'язку "
        "55 армійського корпусу:\n"
        "4. Майора ЗРАЗКОВА, офіцера 77 окремого тестового загону, "
        "призначити офіцером."
    )

    assert set(routes) == {"55 АК А0055", "77 отз А0077"}


def test_unit_name_split_by_soft_break_inside_heading_is_reassembled():
    routes = _routes(
        "§ 4\n"
        "Відповідно до пункту 5 нижчепойменованих осіб офіцерського складу "
        "46 окремого\n"
        "ремонтно-відновлювального полку ЗВІЛЬНИТИ з займаних посад і "
        "ПРИЗНАЧИТИ:\n"
        "5. Майора ТЕСТЕНКА призначити начальником служби."
    )

    assert set(routes) == {"46 орвп А0046"}


def test_multi_soft_break_heading_keeps_complete_source_range():
    routes = _routes(
        "§ 4\n"
        "Відповідно до пункту 5 нижчепойменованих осіб офіцерського складу "
        "46 окремого\n"
        "ремонтно-відновлювального полку звільнити з займаних посад і\n"
        "призначити до цього самого полку:\n"
        "5. Майора ТЕСТЕНКА призначити начальником служби."
    )

    assert set(routes) == {"46 орвп А0046"}
    item = routes["46 орвп А0046"]["items"][0]
    assert item["heading_ranges"] == [(0, 0), (1, 3)]
    assert "ремонтно-відновлювального полку" in item["parent_heading"]
    assert "призначити до цього самого полку" in item["parent_heading"]


def test_routing_searches_complete_heading_when_word_splits_it_into_blocks():
    routes = _routes(
        "§ 4\n"
        "Відповідно до пункту 5 нижчепойменованих осіб офіцерського складу "
        "46 окремого\n"
        "\n"
        "ремонтно-відновлювального полку ЗВІЛЬНИТИ з займаних посад і "
        "ПРИЗНАЧИТИ:\n"
        "5. Майора ТЕСТЕНКА призначити начальником служби."
    )

    assert set(routes) == {"46 орвп А0046"}


def test_routing_searches_physical_heading_paragraphs_not_only_heading_key():
    routes = _routes(
        "§ 4\n"
        "Відповідно до пункту 5 Положення про проходження військової служби\n"
        "Нижчепойменованих осіб офіцерського складу 46 окремого\n"
        "Ремонтно-відновлювального полку ЗВІЛЬНИТИ з займаних посад і\n"
        "ПРИЗНАЧИТИ до цього самого полку:\n"
        "5. Майора ТЕСТЕНКА призначити начальником служби."
    )

    assert set(routes) == {"46 орвп А0046"}


def test_preamble_search_bridges_technical_header_content_boundary():
    routes = _routes(
        "Відповідно\u00a0до пунктів Положення нижчепойменованих осіб "
        "офіцерського складу 46 окремого\n"
        "ремонтно-відновлювального полку ЗВІЛЬНИТИ з займаних посад і "
        "ПРИЗНАЧИТИ до цього самого полку:\n"
        "1. Майора ТЕСТЕНКА призначити начальником служби.\n"
        "2. Капітана ПРИКЛАДЕНКА призначити начальником відділення."
    )

    assert set(routes) == {"46 орвп А0046"}
    assert [item["label"] for item in routes["46 орвп А0046"]["items"]] == [
        "Пункт 1.",
        "Пункт 2.",
    ]


def test_subheading_without_unit_inherits_unit_from_section_heading():
    routes = _routes(
        "§ 5\n"
        "Відповідно до пункту 6 нижчепойменованих осіб офіцерського складу "
        "46 окремого ремонтно-відновлювального полку ЗВІЛЬНИТИ:\n"
        "У ЗАПАС ЗА ПІДПУНКТОМ «А»:\n"
        "6. Майора ТЕСТЕНКА звільнити з військової служби."
    )

    assert set(routes) == {"46 орвп А0046"}


@pytest.mark.parametrize(
    "table_name",
    [
        "8 окремого відновлювального полку",
        "8 окремого-відновлювального полку",
    ],
)
def test_genitive_excel_name_matches_compound_heading_form(table_name):
    """Назва з колонки A у родовому відмінку шукається за стемами слів."""
    mapping = {
        table_name: _entry(table_name, "А0008", "тестовий овп"),
    }
    body = (
        "§ 8\n"
        "Відповідно до пункту 8 Положення нижчепойменованих осіб "
        "офіцерського складу 8 окремого\n"
        "ремонтно-відновлювального полку ЗВІЛЬНИТИ з займаних посад і "
        "ПРИЗНАЧИТИ до цього самого полку:\n"
        "8. Майора ТЕСТЕНКА призначити начальником служби."
    )

    result = map_military_units(text=body, mapping=mapping)

    assert result["unmatched_items"] == []
    assert set(result["unit_paragraphs"]) == {"тестовий овп А0008"}


def test_heading_can_match_by_stable_prefix_of_every_meaningful_word():
    table_name = "208 окремий ремонтно-відновлювальний полк"
    mapping = {
        table_name: _entry(table_name, "А0208", "тестовий полк"),
    }
    body = (
        "§ 9\n"
        "Відповідно до пункту 9 Положення осіб зі складу 208 окремого\n"
        "ремонтно-відновного полку ЗВІЛЬНИТИ і ПРИЗНАЧИТИ:\n"
        "9. Майора ТЕСТЕНКА призначити начальником служби."
    )

    result = map_military_units(text=body, mapping=mapping)

    assert result["unmatched_items"] == []
    assert set(result["unit_paragraphs"]) == {"тестовий полк А0208"}


def test_same_number_without_name_prefixes_is_not_enough():
    table_name = "208 окремий ремонтно-відновлювальний полк"
    mapping = {
        table_name: _entry(table_name, "А0208", "тестовий полк"),
    }
    body = (
        "§ 10\n"
        "Відповідно до пункту 10 Положення осіб зі складу "
        "208 окремого розвідувального полку ПРИЗНАЧИТИ:\n"
        "10. Майора ТЕСТЕНКА призначити начальником служби."
    )

    result = map_military_units(text=body, mapping=mapping)

    assert result["unit_paragraphs"] == {}
    assert len(result["unmatched_items"]) == 1
