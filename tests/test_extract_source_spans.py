from nodeautomationtoolkit.builtin_nodes.recipient_mapping import map_military_units


def test_extract_items_keep_source_spans_including_blank_paragraphs():
    text = """НАКАЗ

§ 1

1. 100 окрема бригада направляє військовослужбовця.

Підстава: рапорт.

2. 100 окрема бригада надає уточнення.
"""
    mapping = {"100 окрема бригада": "військова частина А1000"}

    result = map_military_units(text=text, mapping=mapping)
    items = result["unit_paragraphs"]["військова частина А1000"]["items"]

    assert [(item["source_start_line"], item["source_end_line"]) for item in items] == [
        (4, 7),
        (8, 8),
    ]
    assert (items[0]["heading_start_line"], items[0]["heading_end_line"]) == (2, 3)
