"""Ланцюг підпорядкування в повідомленнях: бат → ВЧ → АК/ОК → ОК (КСВ не шифрується).

Правила користувача: ОК шифрується, КСВ — ні (04.11.2026); ОК і КСВ у розсилку не
входять; галочка «Подовжена версія» додає ОК за стовпцем D, лише якщо його названо в
наказі (03.11.2026). Дані вигадані; ОК у прикладах — «Північ».
"""

from types import SimpleNamespace

from nodeautomationtoolkit.builtin_nodes.message_order import cipher_unit_names, high_commands_in_text
from nodeautomationtoolkit.builtin_nodes.recipient_mapping import _format_full_closed_unit_text
from nodeautomationtoolkit.generator.messages import (
    build_message_recipient_groups,
    is_message_skipped_command,
)

OK = "оперативне командування «Північ»"
KSV = "Командування Сухопутних військ Збройних Сил України"


def _row(name, cipher, abbr, superior="", to=""):
    return {
        "open_name": name,
        "cipher": cipher,
        "abbreviation": abbr,
        "corps": superior,
        "recipient_to": to,
    }


MAPPING = {
    KSV: _row(KSV, "А5555", "КСВ"),
    OK: _row(OK, "А4444", "ОК Північ", "КСВ"),
    "3 армійський корпус": _row("3 армійський корпус", "А3333", "3 АК", "ОК Північ", "Командиру"),
    "2 окрема тестова бригада": _row("2 окрема тестова бригада", "А2222", "2 отбр", "3 АК", "Командиру"),
    "4 окрема тестова бригада": _row("4 окрема тестова бригада", "А4040", "4 отбр", "ОК Північ", "Командиру"),
    "1 окремий тестовий батальйон": _row("1 окремий тестовий батальйон", "А1111", "1 отб", "2 отбр"),
}

BATTALION_ITEM = "1. Офіцера 1 окремого тестового батальйону призначити."


def test_full_chain_goes_up_to_ok_and_never_to_ksv():
    assert _format_full_closed_unit_text(MAPPING["1 окремий тестовий батальйон"], MAPPING, full_chain=True) == (
        "військової частини А1111 військової частини А2222 військової частини А3333 військової частини А4444"
    )
    assert _format_full_closed_unit_text(MAPPING["4 окрема тестова бригада"], MAPPING, full_chain=True) == (
        "військової частини А4040 військової частини А4444"
    )


def test_normal_version_stops_at_corps():
    assert cipher_unit_names(BATTALION_ITEM, MAPPING)[0] == (
        "1. Офіцера військової частини А1111 військової частини А2222 військової частини А3333 призначити."
    )
    assert cipher_unit_names("1. Офіцера 4 окремої тестової бригади призначити.", MAPPING)[0] == (
        "1. Офіцера військової частини А4040 призначити."
    )


def test_extended_version_adds_ok_only_when_it_is_named_in_the_order():
    order = BATTALION_ITEM + "\nу складі оперативного командування «Північ»"
    named = high_commands_in_text(order, MAPPING)
    assert named == frozenset({"А4444"})
    assert cipher_unit_names(BATTALION_ITEM, MAPPING, full_chain=True, high_commands=named)[0] == (
        "1. Офіцера військової частини А1111 військової частини А2222 військової частини А3333 "
        "військової частини А4444 призначити."
    )
    # ОК у наказі не названо — ланцюг до корпусу, як без галочки.
    assert cipher_unit_names(BATTALION_ITEM, MAPPING, full_chain=True, high_commands=frozenset())[0] == (
        "1. Офіцера військової частини А1111 військової частини А2222 військової частини А3333 призначити."
    )


def test_ksv_named_in_the_order_does_not_count_and_stays_open():
    order = "1. Офіцера 1 окремого тестового батальйону командування Сухопутних військ призначити."
    assert high_commands_in_text(order, MAPPING) == frozenset()
    result = cipher_unit_names(order, MAPPING, full_chain=True, high_commands=frozenset())[0]
    assert "А5555" not in result
    assert "командування Сухопутних військ" in result


def test_whole_chain_written_in_the_order():
    # Назви підряд: рядок КСВ не ковтає «командування «Північ»» з назви ОК.
    text = (
        "1. Офіцера 1 окремого тестового батальйону 2 окремої тестової бригади 3 армійського "
        "корпусу оперативного командування «Північ» командування Сухопутних військ призначити."
    )
    assert cipher_unit_names(text, MAPPING)[0] == (
        "1. Офіцера військової частини А1111 військової частини А2222 військової частини А3333 "
        "військової частини А4444 командування Сухопутних військ призначити."
    )


def test_checkbox_controls_the_chain():
    class _Var:
        def __init__(self, value):
            self.value = value

        def get(self):
            return self.value

    from nodeautomationtoolkit.generator.app_messages import MessagesMixin

    assert MessagesMixin._message_full_chain(SimpleNamespace()) is False
    assert MessagesMixin._message_full_chain(SimpleNamespace(message_full_chain=_Var(False))) is False
    assert MessagesMixin._message_full_chain(SimpleNamespace(message_full_chain=_Var(True))) is True


def test_unit_with_command_words_inside_its_name_is_an_ordinary_unit():
    for name, excluded in (
        ("15 окрема бригада оперативного командування «Північ»", False),
        ("навчальний центр Сухопутних військ", False),
        (OK, True),
        ("1 оперативне командування", True),
        (KSV, True),
    ):
        assert is_message_skipped_command(_row(name, "А0001", "")) is excluded, name


def test_ok_and_ksv_are_not_recipients():
    routes = {"unit_paragraphs": {
        "А2222": {"open_name": "2 окрема тестова бригада", "unit_code": "А2222"},
        "А4040": {"open_name": "4 окрема тестова бригада", "unit_code": "А4040"},
        "А4444": {"open_name": OK, "unit_code": "А4444"},
    }}
    groups = build_message_recipient_groups(MAPPING, routes)
    everything = " ".join(groups["corps"] + groups["units"])
    assert "А3333" in " ".join(groups["corps"])  # корпус — адресат
    assert "А2222" in everything and "А4040" in everything
    assert "А4444" not in everything and "А5555" not in everything


def test_troops_commander_of_ok_becomes_unit_commander():
    # 06.11.2026: «командувача військ ОК …» → «командира військової частини …».
    assert cipher_unit_names(
        "1. Відповідно до наказу командувача військ оперативного командування «Північ» від 01.01.2026 № 1.",
        MAPPING,
    )[0] == "1. Відповідно до наказу командира військової частини А4444 від 01.01.2026 № 1."
    assert cipher_unit_names("2. Командувачу військ оперативного командування «Північ» доповісти.", MAPPING)[0] == (
        "2. Командиру військової частини А4444 доповісти."
    )
    assert cipher_unit_names(
        "3. За погодженням із заступником командувача військ оперативного командування «Північ».", MAPPING
    )[0] == "3. За погодженням із заступником командира військової частини А4444."
    # Без шифру після «військ» нічого не міняється.
    text = "4. Командувача військ Тестової групи повідомити."
    assert cipher_unit_names(text, MAPPING)[0] == text


def test_ok_named_without_service_branch_tail_is_still_ciphered():
    full = "Оперативне командування «Північ» Сухопутних військ Збройних Сил України"
    mapping = {key: value for key, value in MAPPING.items() if key != OK}
    mapping[full] = _row(full, "А4444", "ОК Північ", "КСВ")
    assert cipher_unit_names("1. Офіцера оперативного командування «Північ» призначити.", mapping)[0] == (
        "1. Офіцера військової частини А4444 призначити."
    )
    # Хвіст «Сухопутних військ …» — це КСВ: шифр ОК його не ковтає й не прибирає.
    assert cipher_unit_names(
        "1. Офіцера оперативного командування «Північ» Сухопутних військ Збройних Сил України призначити.", mapping
    )[0] == "1. Офіцера військової частини А4444 Сухопутних військ Збройних Сил України призначити."
    # КСВ так само лишається відкритим.
    ksv = "2. Офіцера командування Сухопутних військ Збройних Сил України призначити."
    assert cipher_unit_names(ksv, mapping)[0] == ksv
