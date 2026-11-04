"""Повний ланцюг підпорядкування в повідомленнях: бат → ВЧ → АК/ОК → ОК → КСВ.

ОК і КСВ повідомлення не отримують (03.11.2026, КСВ — тимчасово). Дані вигадані.
"""

from nodeautomationtoolkit.builtin_nodes.message_order import cipher_unit_names
from nodeautomationtoolkit.builtin_nodes.recipient_mapping import _format_full_closed_unit_text
from nodeautomationtoolkit.generator.messages import (
    build_message_recipient_groups,
    is_message_skipped_command,
)


def _row(name, cipher, abbr, superior="", to=""):
    return {
        "open_name": name,
        "cipher": cipher,
        "abbreviation": abbr,
        "corps": superior,
        "recipient_to": to,
    }


MAPPING = {
    "командування сухопутних військ": _row("командування сухопутних військ", "А5555", "КСВ"),
    "1 оперативне командування": _row("1 оперативне командування", "А4444", "ОК Тест", "КСВ"),
    "77 армійський корпус": _row("77 армійський корпус", "А3333", "77 АК", "ОК Тест", "Командиру"),
    "10 окрема механізована бригада": _row(
        "10 окрема механізована бригада", "А1010", "10 омбр", "77 АК", "Командиру"
    ),
    "20 окрема механізована бригада": _row(
        "20 окрема механізована бригада", "А2020", "20 омбр", "ОК Тест", "Командиру"
    ),
    "1 механізований батальйон 10 окремої механізованої бригади": _row(
        "1 механізований батальйон 10 окремої механізованої бригади", "А0101", "1 мб", "10 омбр"
    ),
}


def test_chain_goes_through_corps_up_to_ksv():
    assert _format_full_closed_unit_text(MAPPING["1 механізований батальйон 10 окремої механізованої бригади"], MAPPING, full_chain=True) == (
        "військової частини А0101 військової частини А1010 військової частини А3333 "
        "військової частини А4444 військової частини А5555"
    )


def test_unit_directly_under_ok():
    assert _format_full_closed_unit_text(MAPPING["20 окрема механізована бригада"], MAPPING, full_chain=True) == (
        "військової частини А2020 військової частини А4444 військової частини А5555"
    )


def test_extracts_chain_still_stops_at_corps():
    assert _format_full_closed_unit_text(MAPPING["10 окрема механізована бригада"], MAPPING) == (
        "військової частини А1010 військової частини А3333"
    )


def test_message_text_has_each_cipher_once():
    text = "1. Капітана ТЕСТЕНКА, офіцера 10 окремої механізованої бригади 77 армійського корпусу."
    result = cipher_unit_names(text, MAPPING, full_chain=True)[0]
    assert result == (
        "1. Капітана ТЕСТЕНКА, офіцера військової частини А1010 військової частини А3333 "
        "військової частини А4444 військової частини А5555."
    )


def test_full_chain_takes_only_commands_named_in_the_order():
    from nodeautomationtoolkit.builtin_nodes.message_order import high_commands_in_text

    order = "1 оперативного командування НАКАЗУЮ:\n1. Офіцера 10 окремої механізованої бригади."
    named = high_commands_in_text(order, MAPPING)
    assert named == frozenset({"А4444"})
    text = "1. Офіцера 10 окремої механізованої бригади."
    assert cipher_unit_names(text, MAPPING, full_chain=True, high_commands=named)[0] == (
        "1. Офіцера військової частини А1010 військової частини А3333 військової частини А4444."
    )
    # Ні ОК, ні КСВ у наказі немає — ланцюг до корпусу.
    assert cipher_unit_names(text, MAPPING, full_chain=True, high_commands=frozenset())[0] == (
        "1. Офіцера військової частини А1010 військової частини А3333."
    )
    # Лише КСВ — ОК пропускається, ланцюг іде далі.
    assert cipher_unit_names(text, MAPPING, full_chain=True, high_commands=frozenset({"А5555"}))[0] == (
        "1. Офіцера військової частини А1010 військової частини А3333 військової частини А5555."
    )


def test_normal_version_stops_at_corps():
    text = "1. Капітана ТЕСТЕНКА, офіцера 10 окремої механізованої бригади 77 армійського корпусу."
    assert cipher_unit_names(text, MAPPING)[0] == (
        "1. Капітана ТЕСТЕНКА, офіцера військової частини А1010 військової частини А3333."
    )


def test_checkbox_controls_the_chain():
    from types import SimpleNamespace

    from nodeautomationtoolkit.generator.app_messages import MessagesMixin

    class _Var:
        def __init__(self, value):
            self.value = value

        def get(self):
            return self.value

    assert MessagesMixin._message_full_chain(SimpleNamespace()) is False
    assert MessagesMixin._message_full_chain(SimpleNamespace(message_full_chain=_Var(False))) is False
    assert MessagesMixin._message_full_chain(SimpleNamespace(message_full_chain=_Var(True))) is True



def test_unit_with_command_words_inside_its_name_is_an_ordinary_unit():
    for name, abbr in (
        ("15 окрема бригада оперативного командування Тест", "15 обр"),
        ("навчальний центр Сухопутних військ", "НЦ"),
        ("оперативне командування Тест", ""),
        ("командування Сухопутних військ", ""),
    ):
        entry = _row(name, "А0001", abbr)
        assert is_message_skipped_command(entry) is name.startswith(("оперативне", "командування"))


def test_ok_and_ksv_are_not_recipients():
    assert is_message_skipped_command(MAPPING["1 оперативне командування"])
    assert is_message_skipped_command(MAPPING["командування сухопутних військ"])
    assert not is_message_skipped_command(MAPPING["77 армійський корпус"])

    routes = {"unit_paragraphs": {
        "А1010": {"open_name": "10 окрема механізована бригада", "unit_code": "А1010"},
        "А2020": {"open_name": "20 окрема механізована бригада", "unit_code": "А2020"},
        "А4444": {"open_name": "1 оперативне командування", "unit_code": "А4444"},
    }}
    groups = build_message_recipient_groups(MAPPING, routes)
    everything = " ".join(groups["corps"] + groups["units"])
    assert "А3333" in " ".join(groups["corps"])  # корпус — адресат
    assert "А1010" in everything and "А2020" in everything
    assert "А4444" not in everything and "А5555" not in everything


FIVE_LEVELS = {
    "1 окремий тестовий батальйон": _row("1 окремий тестовий батальйон", "А1111", "1 отб", "2 отбр"),
    "2 окрема тестова бригада": _row("2 окрема тестова бригада", "А2222", "2 отбр", "3 АК"),
    "3 армійський корпус": _row("3 армійський корпус", "А3333", "3 АК", "ОК Тест"),
    "оперативне командування Тест": _row("оперативне командування Тест", "А4444", "ОК Тест", "КСВ"),
    "командування Сухопутних військ": _row("командування Сухопутних військ", "А5555", "КСВ"),
}


def test_five_level_chain_written_in_the_order():
    # Назви підряд: шаблон КСВ не має ковтати «командування Тест» з назви ОК.
    text = (
        "1. Офіцера 1 окремого тестового батальйону 2 окремої тестової бригади 3 армійського "
        "корпусу оперативного командування «Тест» командування Сухопутних військ призначити."
    )
    assert cipher_unit_names(text, FIVE_LEVELS)[0] == (
        "1. Офіцера військової частини А1111 військової частини А2222 військової частини А3333 "
        "військової частини А4444 військової частини А5555 призначити."
    )


def test_five_level_chain_from_column_d_when_ok_and_ksv_are_named_elsewhere():
    from nodeautomationtoolkit.builtin_nodes.message_order import high_commands_in_text

    item = "1. Офіцера 1 окремого тестового батальйону призначити."
    order = item + "\nоперативного командування Тест командування Сухопутних військ"
    named = high_commands_in_text(order, FIVE_LEVELS)
    assert named == frozenset({"А4444", "А5555"})
    assert cipher_unit_names(item, FIVE_LEVELS, full_chain=True, high_commands=named)[0] == (
        "1. Офіцера військової частини А1111 військової частини А2222 військової частини А3333 "
        "військової частини А4444 військової частини А5555 призначити."
    )
    assert cipher_unit_names(item, FIVE_LEVELS)[0] == (
        "1. Офіцера військової частини А1111 військової частини А2222 військової частини А3333 призначити."
    )
