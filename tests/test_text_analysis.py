from nodeautomationtoolkit.builtin_nodes.text_analysis import (
    extract_order_fields,
    find_in_text,
    group_items_by_markers,
    replace_in_text,
    split_order_blocks,
)

ORDER_TEXT = """ПАРАГРАФ 1
Причина: для перевірки розподілу
1. Направити матеріали до АЛЬФА.
2. Направити копії до АЛЬФА та БРАВО.
Підстава: наказ від 31.07.2026 № 123/ос. Примірник № 2
Командир військової частини А0000
"""


def test_find_and_replace_text_reports_positions():
    found = find_in_text(ORDER_TEXT, "АЛЬФА")
    replaced = replace_in_text(ORDER_TEXT, "АЛЬФА", "ГРУПА", replace_all=False)

    assert found["count"] == 2
    assert found["matches"][0]["line"] == 3
    assert found["matches"][0]["start"] >= 0
    assert replaced["replacements"] == 1


def test_extracts_order_fields():
    fields = extract_order_fields(ORDER_TEXT)

    assert fields["order_date"] == "31.07.2026"
    assert fields["order_number"] == "123/ос."
    assert fields["copy_number"] == "2"
    assert fields["positions"] == ["Командир військової частини А0000"]


def test_groups_one_item_into_two_documents_with_context():
    split = split_order_blocks(ORDER_TEXT)
    grouped = group_items_by_markers(split["blocks"], "АЛЬФА\nБРАВО")

    assert grouped["names"] == ["АЛЬФА", "БРАВО"]
    assert "1. Направити матеріали" in grouped["groups"]["АЛЬФА"]
    assert "2. Направити копії" in grouped["groups"]["АЛЬФА"]
    assert "2. Направити копії" in grouped["groups"]["БРАВО"]
    assert "ПАРАГРАФ 1" in grouped["groups"]["БРАВО"]
    assert grouped["counts"] == {"АЛЬФА": 2, "БРАВО": 1}
    assert split["signature"] == "Командир військової частини А0000"
    assert "Командир" not in split["body"]


def test_accepts_marker_list_from_mapping_table():
    split = split_order_blocks(ORDER_TEXT)
    grouped = group_items_by_markers(split["blocks"], markers=["БРАВО"])

    assert grouped["names"] == ["БРАВО"]


def test_marker_in_reason_header_applies_to_following_items():
    text = """Відповідно до рішення для АЛЬФА
1. Направити матеріали до БРАВО.
2. Другий пункт без мітки.
Командир частини
"""
    split = split_order_blocks(text)
    grouped = group_items_by_markers(split["blocks"], markers=["АЛЬФА", "БРАВО"])

    assert grouped["counts"] == {"АЛЬФА": 2, "БРАВО": 1}
    assert "2. Другий пункт" in grouped["groups"]["АЛЬФА"]
    assert "1. Направити матеріали" in grouped["groups"]["БРАВО"]


def test_marker_in_action_header_applies_to_following_items():
    text = """ПАРАГРАФ 2
Відповідно до рішення кадрової комісії
ЗВІЛЬНИТИ І ПРИЗНАЧИТИ до АЛЬФА
1. Перший пункт без повторення відправника.
2. Другий пункт без повторення відправника.
ПРИЗНАЧИТИ до БРАВО
3. Третій пункт.
Командир частини
"""
    split = split_order_blocks(text)
    grouped = group_items_by_markers(split["blocks"], markers=["АЛЬФА", "БРАВО"])

    assert split["action_headers"] == [
        "ЗВІЛЬНИТИ І ПРИЗНАЧИТИ до АЛЬФА",
        "ПРИЗНАЧИТИ до БРАВО",
    ]
    assert grouped["counts"] == {"АЛЬФА": 2, "БРАВО": 1}
    assert "ЗВІЛЬНИТИ І ПРИЗНАЧИТИ до АЛЬФА" in grouped["groups"]["АЛЬФА"]
    assert "1. Перший пункт" in grouped["groups"]["АЛЬФА"]
    assert "2. Другий пункт" in grouped["groups"]["АЛЬФА"]
    assert "3. Третій пункт" not in grouped["groups"]["АЛЬФА"]
    assert "ПРИЗНАЧИТИ до БРАВО" in grouped["groups"]["БРАВО"]
    assert "3. Третій пункт" in grouped["groups"]["БРАВО"]
