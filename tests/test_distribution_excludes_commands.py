"""ОК і КСВ у розсилку не входять — ні витягів, ні повідомлень (користувач, 04.11.2026).

Усі дані вигадані. Стовпець D частини може посилатися на ОК: тоді частина
отримує витяг сама, а не передає його командуванню.
"""

from nodeautomationtoolkit.builtin_nodes.recipient_mapping import (
    is_distribution_excluded,
    map_military_units,
)
from nodeautomationtoolkit.generator.messages import build_message_recipient_groups

OK = "оперативне командування «Північ»"
KSV = "Командування Сухопутних військ Збройних Сил України"


def _entry(open_name, cipher, abbreviation="", corps="", recipient_to=""):
    return {
        "open_name": open_name,
        "cipher": cipher,
        "abbreviation": abbreviation,
        "corps": corps,
        "recipient_to": recipient_to,
        "destination_where": "м. Тест",
    }


MAPPING = {
    "77 окрема тестова бригада": _entry(
        "77 окрема тестова бригада", "А0077", "77 отбр", OK, "Командиру військової частини"
    ),
    "88 окремий тестовий полк": _entry(
        "88 окремий тестовий полк", "А0088", "88 отп", "", "Командиру військової частини"
    ),
    OK: _entry(OK, "А0001", "ОК «Північ»", KSV, "Командувачу військ"),
    KSV: _entry(KSV, "А0002", "КСВ", "", "Командувачу"),
}

ORDER = """НАКАЗ
§ 1
Відповідно до статті 1 ПРИЗНАЧИТИ:
1. Капітана Тестенка Т.Т., командира роти 77 окремої тестової бригади оперативного командування «Північ» Сухопутних військ Збройних Сил України, – КОМАНДИРОМ РОТИ 88 ОКРЕМОГО ТЕСТОВОГО ПОЛКУ.
1990 р.н.
2. Майора Тестова Т.Т., офіцера відділу оперативного командування «Північ», – СТАРШИМ ОФІЦЕРОМ ВІДДІЛУ КОМАНДУВАННЯ СУХОПУТНИХ ВІЙСЬК ЗБРОЙНИХ СИЛ УКРАЇНИ.
1985 р.н.
"""


def test_commands_are_recognised():
    assert is_distribution_excluded(MAPPING[OK])
    assert is_distribution_excluded(MAPPING[KSV])
    assert is_distribution_excluded("КСВ")
    assert not is_distribution_excluded(MAPPING["77 окрема тестова бригада"])
    assert not is_distribution_excluded("11 армійський корпус")


def test_extracts_go_to_units_not_to_commands():
    routes = map_military_units(ORDER, MAPPING)
    assert sorted(routes["unit_paragraphs"]) == ["77 отбр А0077", "88 отп А0088"]
    # Бригада з D = ОК отримує витяг сама.
    assert [item["label"] for item in routes["unit_paragraphs"]["77 отбр А0077"]["items"]] == ["Пункт 1."]


def test_item_only_about_commands_is_skipped_with_reason_not_lost():
    routes = map_military_units(ORDER, MAPPING)
    assert routes["unmatched_items"] == []
    skipped = {item["label"]: item["reason"] for item in routes["skipped_items"]}
    assert "ОК і КСВ у розсилку не входять" in skipped["Пункт 2."]


def test_messages_have_no_command_recipients():
    routes = map_military_units(ORDER, MAPPING)
    groups = build_message_recipient_groups(MAPPING, routes)
    assert groups["corps"] == []
    assert groups["units"] == ["Командиру військової частини А0077", "Командиру військової частини А0088"]


def test_section_heading_with_command_falls_back_to_preamble_unit():
    """Шапка § називає лише ОК — пункт бере адресата з преамбули наказу."""
    order = """НАКАЗ
Відповідно до статті 1 по 88 окремому тестовому полку
§ 1
Офіцерів оперативного командування «Північ» ПРИЗНАЧИТИ:
1. Капітана Тестенка Т.Т. – КОМАНДИРОМ РОТИ.
1990 р.н.
"""
    routes = map_military_units(order, MAPPING)
    assert "ОК «Північ» А0001" not in routes["unit_paragraphs"]
