"""Вікно порівняння: у зміненому рядку фон мають лише неправильні слова.

Шифр іншої частини — червоним (помилка), решта розбіжностей — жовтим, а весь
рядок фону не отримує. Рядки без розмітки слів (пропущені, зайві) — як раніше.
"""

import os
import sys
from pathlib import Path

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
pytest.importorskip("PySide6")

SRC_PATH = Path(__file__).resolve().parents[1] / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from PySide6.QtCore import Qt  # noqa: E402
from PySide6.QtGui import QColor, QTextCursor  # noqa: E402
from PySide6.QtWidgets import QApplication, QTextEdit  # noqa: E402

from nodeautomationtoolkit.builtin_nodes.compare_documents import word_diff  # noqa: E402
from nodeautomationtoolkit.generator_qt import theme  # noqa: E402
from nodeautomationtoolkit.generator_qt.compare_window import CompareWindow  # noqa: E402


@pytest.fixture(scope="module")
def qt_app():
    return QApplication.instance() or QApplication([])


def _background_at(view: QTextEdit, block_number: int, column: int) -> str:
    block = view.document().findBlockByNumber(block_number)
    cursor = QTextCursor(block)
    cursor.setPosition(block.position() + column + 1)  # формат символу ЛІВОРУЧ від курсора
    brush = cursor.charFormat().background()
    return brush.color().name() if brush.style() != Qt.BrushStyle.NoBrush else ""


def test_only_wrong_words_are_highlighted(qt_app):
    ref = "Командиру військової частини А1111 надіслати копію наказу."
    gen = "Командиру військової частини А2222 надіслати копію накзу."
    ref_spans, gen_spans = word_diff(ref, gen)
    rows = [
        {"status": "EQUAL", "ref_line": "§ 1", "gen_line": "§ 1"},
        {"status": "MODIFIED", "ref_line": ref, "gen_line": gen, "ref_spans": ref_spans, "gen_spans": gen_spans},
        {"status": "DELETED", "ref_line": "пропущений абзац", "gen_line": ""},
    ]
    view = QTextEdit()
    CompareWindow._fill(view, rows, "gen_line")

    changed = QColor(theme.WORD_TONES["changed"][0]).name()
    unit = QColor(theme.WORD_TONES["unit"][0]).name()
    assert view.document().findBlockByNumber(1).text() == gen
    assert _background_at(view, 1, gen.index("А2222")) == unit
    assert _background_at(view, 1, gen.index("накзу")) == changed
    assert _background_at(view, 1, gen.index("Командиру")) == ""
    # Змінений рядок із розміткою не зафарбовується цілком.
    assert view.document().findBlockByNumber(1).blockFormat().background().style() == Qt.BrushStyle.NoBrush
    # Пропущений абзац — як раніше, фоном усього рядка.
    assert view.document().findBlockByNumber(2).blockFormat().background().style() != Qt.BrushStyle.NoBrush
