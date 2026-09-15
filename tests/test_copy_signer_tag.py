"""Підписант як окремий тег заготовки примірника.

Якщо у заготовці є `{{підписант}}`, блок підписанта підставляється в нього,
а `{{зміст}}` завершується перед підписантом — щоб він не дублювався.
Без цього тегу підписант лишається частиною змісту (старі заготовки).
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
        "generate_extracts_signer_tag_tests", PROJECT_ROOT / "generate_extracts.py"
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


generator = _load_generator_module()
CELL_MARK = chr(7)

SIGNER_LINE = "майор              Петро ПЕТРЕНКО"


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


class _Content:
    def __init__(self, text):
        self.Text = text


class _FakeDoc:
    def __init__(self, texts):
        self.Paragraphs = _Paragraphs(texts)
        self.Content = _Content("".join(texts))


ORDER = [
    "§ 2\r",                             # 1
    "11. Пункт наказу.\r",               # 2 ← кінець змісту, коли є тег
    "\r",
    "Тимчасово виконуючий обов'язки\r",  # 4 ← початок підписанта
    "командувача військ\r",              # 5
    SIGNER_LINE + "\r",                  # 6 ← звання + ПІБ
    "Комірка звороту" + CELL_MARK,       # 7 ← таблиця, ігнорується
]


def _context(signer_as_tag: bool) -> dict:
    app = generator.App.__new__(generator.App)
    return generator.App._order_body_context(app, _FakeDoc(ORDER), signer_as_tag=signer_as_tag)


def test_without_tag_signer_stays_inside_content():
    result = _context(signer_as_tag=False)

    assert result["span"] == (1, 6)
    assert "{{підписант}}" not in result["values"]
    assert result["signature_line"] == ""


def test_with_tag_content_ends_before_signer():
    result = _context(signer_as_tag=True)

    assert result["span"] == (1, 2)


def test_with_tag_signer_block_goes_into_the_tag():
    values = _context(signer_as_tag=True)["values"]

    block = values["{{підписант}}"]
    assert block.startswith("Тимчасово виконуючий обов'язки")
    assert block.endswith(SIGNER_LINE)
    assert block.count("\r") == 2  # три рядки блоку


def test_signer_spacing_is_preserved():
    """Звання і ПІБ вирівняні пробілами — переносимо як є."""
    block = _context(signer_as_tag=True)["values"]["{{підписант}}"]
    assert SIGNER_LINE in block


def test_signature_line_is_reported_for_formatting():
    assert _context(signer_as_tag=True)["signature_line"] == SIGNER_LINE


def test_back_page_table_never_enters_the_signer_block():
    block = _context(signer_as_tag=True)["values"]["{{підписант}}"]
    assert "Комірка звороту" not in block
