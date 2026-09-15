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


def test_same_center_reference_does_not_create_generic_extra_recipients():
    """«Цього самого центру» підтверджує явний центр, а не шукає новий."""
    exact = "901 центр підготовки"
    generic = "центр підготовки"
    shortest = "центр"
    mapping = {
        exact: _entry(exact, "А0901", "901 цп"),
        generic: _entry(generic, "А0002", "загальний цп"),
        shortest: _entry(shortest, "А0003", "загальний центр"),
    }
    body = (
        "§ 9\n"
        "ПРИЗНАЧИТИ:\n"
        "9. Майора ТЕСТЕНКА, офіцера 901 центру підготовки, "
        "начальником служби цього самого центру."
    )

    result = map_military_units(text=body, mapping=mapping)

    assert result["unmatched_items"] == []
    assert set(result["unit_paragraphs"]) == {"901 цп А0901"}


@pytest.mark.parametrize("reverse", [False, True])
def test_numbered_center_wins_over_overlapping_adjectival_alias(reverse):
    """Точний номер у рядку A сильніший за загальну назву того ж фрагмента."""
    numbered = "902 центр підготовки"
    generic = "окремий центр підготовки"
    entries = [
        (numbered, _entry(numbered, "А0902", "902 цп")),
        (generic, _entry(generic, "А0002", "загальний оцп")),
    ]
    mapping = dict(reversed(entries) if reverse else entries)
    body = (
        "§ 10\nПРИЗНАЧИТИ:\n"
        "10. Майора ТЕСТЕНКА, офіцера 902 окремого центру підготовки, "
        "начальником служби."
    )

    result = map_military_units(text=body, mapping=mapping)

    assert set(result["unit_paragraphs"]) == {"902 цп А0902"}


def test_distinct_non_overlapping_generic_unit_mention_is_not_suppressed():
    """Відсікається вкладений збіг, але не окрема явна згадка іншого рядка A."""
    numbered = "903 центр підготовки"
    generic = "центр забезпечення"
    mapping = {
        numbered: _entry(numbered, "А0903", "903 цп"),
        generic: _entry(generic, "А0003", "цз"),
    }
    body = (
        "§ 11\nПРИЗНАЧИТИ:\n"
        "11. Майора ТЕСТЕНКА перевести з 903 центру підготовки "
        "до центру забезпечення."
    )

    result = map_military_units(text=body, mapping=mapping)

    assert set(result["unit_paragraphs"]) == {"903 цп А0903", "цз А0003"}


def test_global_preamble_can_name_more_than_one_source_unit():
    first = "904 окремий тестовий полк"
    second = "905 окремий тестовий батальйон"
    mapping = {
        first: _entry(first, "А0904", "904 отп"),
        second: _entry(second, "А0905", "905 отб"),
    }
    body = (
        f"Нижчепойменованих осіб {first} та {second} ПРИЗНАЧИТИ:\n"
        "1. Майора ТЕСТЕНКА призначити начальником служби."
    )

    result = map_military_units(text=body, mapping=mapping)

    assert set(result["unit_paragraphs"]) == {"904 отп А0904", "905 отб А0905"}


def test_text_cipher_does_not_replace_anaphoric_reference_as_another_unit():
    numbered = "906 центр підготовки"
    generic = "центр"
    mapping = {
        numbered: _entry(numbered, "А0906", "906 цп"),
        generic: _entry(generic, "А0006", "загальний центр"),
    }
    body = (
        "§ 12\nПРИЗНАЧИТИ:\n"
        "12. Майора ТЕСТЕНКА, офіцера 906 центру підготовки, "
        "начальником служби цього самого центру."
    )

    result = map_military_units(text=body, mapping=mapping)

    unit = result["unit_paragraphs"]["906 цп А0906"]
    cipher_text = unit["items"][0]["text_cipher"]
    assert "А0906" in cipher_text
    assert "цього самого центру" in cipher_text
    assert "А0006" not in cipher_text


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


@pytest.mark.parametrize("global_header", ["", f"{OTHER}\n"])
def test_first_section_is_not_the_global_source_for_later_sections(global_header):
    result = map_military_units(
        text=(
            global_header + f"§ 1\nПо {REPAIR} ПРИЗНАЧИТИ:\n"
            "1. Майора ТЕСТЕНКА начальником служби.\n"
            "§ 2\nПРИЗНАЧИТИ:\n"
            f"2. Капітана ПРИКЛАДЕНКА офіцером {SIGNAL}.\n"
            "3. Лейтенанта ЗРАЗКОВА начальником служби.\n"
        ),
        mapping=MAPPING,
    )
    routes = result["unit_paragraphs"]
    assert [i["label"] for i in routes["46 орвп А0046"]["items"]] == ["Пункт 1."]
    assert [i["label"] for i in routes["55 АК А0055"]["items"]] == ["Пункт 2."]
    if global_header:
        assert result["preamble_recipient"] == OTHER
        assert [i["label"] for i in routes["77 отз А0077"]["items"]] == ["Пункт 2.", "Пункт 3."]
    else:
        assert result["preamble_recipient"] == ""
        assert [i["label"] for i in result["unmatched_items"]] == ["Пункт 3."]
