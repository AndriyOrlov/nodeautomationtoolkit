"""Відповідність «рядок тексту ↔ абзац Word».

`Content.Text` склеює цілий рядок таблиці в один рядок тексту, тоді як
`doc.Paragraphs` рахує кожну комірку окремим абзацом. Через це після будь-якої
таблиці в тілі наказу пункти зіставлялися не з тими абзацами й губилися.
"""

import importlib.util
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))


def _load_generator_module():
    spec = importlib.util.spec_from_file_location(
        "generate_extracts_alignment_tests", PROJECT_ROOT / "generate_extracts.py"
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


generator = _load_generator_module()
read_document_text = generator.read_document_text

CELL = chr(7)


class _Range:
    def __init__(self, text):
        self.Text = text


class _Paragraph:
    def __init__(self, text):
        self.Range = _Range(text)


class _Paragraphs:
    def __init__(self, texts):
        self._items = [_Paragraph(t) for t in texts]
        self.Count = len(self._items)

    def __call__(self, index):
        return self._items[index - 1]


class _FakeDoc:
    def __init__(self, texts):
        self.Paragraphs = _Paragraphs(texts)


def _line_to_paragraph(doc):
    """Повторює побудову карти з run_extracts."""
    mapping = []
    for index in range(1, doc.Paragraphs.Count + 1):
        raw = doc.Paragraphs(index).Range.Text
        mapping.extend([index] * generator.paragraph_line_count(raw))
    return mapping


ORDER_WITH_TABLE = [
    "§ 1\r",              # 1
    "1. Перший\r",        # 2
    "A" + CELL,           # 3 комірка таблиці
    "B" + CELL,           # 4 комірка таблиці
    "2. Другий\r",        # 5
    "3. Третій\r",        # 6
]


def test_every_line_maps_to_its_own_paragraph():
    doc = _FakeDoc(ORDER_WITH_TABLE)
    lines = read_document_text(doc).splitlines()
    mapping = _line_to_paragraph(doc)

    assert len(lines) == len(mapping) == doc.Paragraphs.Count
    for line_index, line in enumerate(lines):
        paragraph_index = mapping[line_index]
        expected = ORDER_WITH_TABLE[paragraph_index - 1].rstrip("\r" + CELL)
        assert line == expected


def test_items_after_a_table_are_not_replaced_by_cells():
    """Раніше «2. Другий» зіставлявся з коміркою «A»."""
    doc = _FakeDoc(ORDER_WITH_TABLE)
    lines = read_document_text(doc).splitlines()

    assert lines[4] == "2. Другий"
    assert lines[5] == "3. Третій"


def test_table_cells_stay_separate_lines():
    lines = read_document_text(_FakeDoc(ORDER_WITH_TABLE)).splitlines()
    assert lines[2] == "A"
    assert lines[3] == "B"


def test_document_without_tables_is_unchanged():
    doc = _FakeDoc(["§ 1\r", "1. Пункт\r"])
    assert read_document_text(doc).splitlines() == ["§ 1", "1. Пункт"]


def test_empty_paragraph_in_the_middle_keeps_alignment():
    doc = _FakeDoc(["§ 1\r", "\r", "1. Пункт\r"])
    lines = read_document_text(doc).splitlines()

    assert lines == ["§ 1", "", "1. Пункт"]
    assert _line_to_paragraph(doc)[2] == 3


def test_trailing_empty_paragraph_is_harmless():
    """Word завжди має останній порожній абзац — він не дає рядка.

    Це не порушує вирівнювання: усі наявні рядки лишаються на своїх абзацах,
    а в кінцевому порожньому абзаці пунктів не буває.
    """
    doc = _FakeDoc(["§ 1\r", "1. Пункт\r", "\r"])
    lines = read_document_text(doc).splitlines()

    assert lines == ["§ 1", "1. Пункт"]
    mapping = _line_to_paragraph(doc)
    for index, line in enumerate(lines):
        assert line == doc.Paragraphs(mapping[index]).Range.Text.rstrip("\r" + CELL)


def test_empty_document_is_safe():
    assert read_document_text(_FakeDoc([])) == ""


def test_break_at_paragraph_end_does_not_shift_the_map():
    """Розрив сторінки чи мʼякий перенос у КІНЦІ абзацу дає ще один рядок.

    Карта рахувала `raw.splitlines()` — на рядок менше, ніж текст. Після
    такого абзацу межі пунктів зсувались на абзац: пункт губив біографію,
    а сусідній захоплював чужий абзац (дубль у витягу).
    """
    doc = _FakeDoc(
        [
            "§ 1\r",
            "1. Перший\r",
            "біографія першого\x0c\r",  # розрив сторінки в кінці абзацу
            "2. Другий\r",
            "рядок\x0bбіографії\x0b\r",  # мʼякий перенос у кінці абзацу
            "3. Третій\r",
            "\x0c\r",  # абзац лише з розривом
            "4. Четвертий\r",
        ]
    )
    lines = read_document_text(doc).splitlines()
    mapping = _line_to_paragraph(doc)

    assert len(mapping) >= len(lines)
    for number, word in ((1, "Перший"), (2, "Другий"), (3, "Третій"), (4, "Четвертий")):
        line_index = lines.index(f"{number}. {word}")
        paragraph = doc.Paragraphs(mapping[line_index]).Range.Text
        assert paragraph.startswith(f"{number}. {word}")


def test_page_break_only_paragraph_is_told_apart():
    assert generator.is_page_break_only_text("\x0c\r")
    assert generator.is_page_break_only_text(" \x0c \r")
    assert not generator.is_page_break_only_text("біографія\x0c\r")
    assert not generator.is_page_break_only_text("біографія\r")
