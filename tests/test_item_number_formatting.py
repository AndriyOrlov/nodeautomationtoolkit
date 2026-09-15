from nodeautomationtoolkit.builtin_nodes.recipient_mapping import _format_item_numbers_range


def test_formats_simple_consecutive_item_numbers_as_range():
    assert _format_item_numbers_range(["Пункт 1", "Пункт 2", "Пункт 3", "Пункт 4", "Пункт 5"]) == "1-5"


def test_preserves_hierarchical_item_numbers():
    assert _format_item_numbers_range(["Пункт 1.1", "Пункт 1.2", "Пункт 2.1", "Пункт 3"]) == "1.1,1.2,2.1,3"
