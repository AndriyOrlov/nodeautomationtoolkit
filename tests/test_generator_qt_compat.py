"""Міст Tk→Qt: логіка генератора бачить ту саму поверхню, що й у Tk.

Логіка `generate_extracts.App` написана під tkinter. Qt-оболонка підміняє
лише цю поверхню, тож тут перевіряється, що підміна поводиться як Tk саме
в тих місцях, на які логіка спирається.
"""

import importlib.util
import os
import sys
from pathlib import Path

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
pytest.importorskip("PySide6")

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from PySide6.QtCore import Qt  # noqa: E402
from PySide6.QtGui import QGuiApplication  # noqa: E402
from PySide6.QtWidgets import QApplication, QFileDialog, QMessageBox, QTabWidget, QTreeWidget, QWidget  # noqa: E402

from nodeautomationtoolkit.generator_qt import compat  # noqa: E402


@pytest.fixture(scope="module")
def qt_app():
    return QApplication.instance() or QApplication([])


def _load_generator():
    spec = importlib.util.spec_from_file_location(
        "generate_extracts_qt_compat_tests", PROJECT_ROOT / "generate_extracts.py"
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_variables_coerce_like_tk():
    assert compat.StringVar().get() == ""
    assert compat.StringVar(value=5).get() == "5"
    assert compat.BooleanVar(value="0").get() is False
    assert compat.BooleanVar(value=1).get() is True
    assert compat.DoubleVar(value="1.5").get() == 1.5


def test_variable_notifies_only_on_change_and_drops_deleted_widgets():
    variable = compat.StringVar(value="a")
    seen = []
    variable.subscribe(seen.append)

    def deleted_widget(_value):
        raise RuntimeError("Internal C++ object (QLineEdit) already deleted.")

    variable.subscribe(deleted_widget)
    variable.set("a")
    assert seen == []
    variable.set("b")
    variable.set("c")
    assert seen == ["b", "c"]

    def broken_logic(_value):
        raise RuntimeError("справжня помилка")

    variable.subscribe(broken_logic)
    with pytest.raises(RuntimeError, match="справжня"):
        variable.set("d")


def test_tree_bridge_keeps_tk_marks_as_real_checkboxes(qt_app):
    widget = QTreeWidget()
    widget.setColumnCount(2)
    tree = compat.TreeBridge(widget)
    toggled = []
    tree.mark_callbacks.append(lambda: toggled.append(True))

    hint = tree.insert("", compat.END, values=("—", "Накази не обрано"))
    order = tree.insert("", compat.END, iid="order_0", values=(compat.MARK_ON, "C:/Накази/№ 1.docx"))
    assert tree.get_children() == (hint, "order_0")
    assert tree.item("order_0", "values") == (compat.MARK_ON, "C:/Накази/№ 1.docx")
    assert widget.topLevelItem(1).checkState(0) == Qt.CheckState.Checked
    assert widget.topLevelItem(0).data(0, Qt.ItemDataRole.CheckStateRole) is None

    # Користувач знімає галочку в таблиці — логіка бачить «☐».
    widget.topLevelItem(1).setCheckState(0, Qt.CheckState.Unchecked)
    assert tree.item(order, "values")[0] == compat.MARK_OFF
    assert toggled == [True]

    tree.item(order, values=(compat.MARK_ON, "C:/Накази/№ 2.docx"))
    assert tree.item(order, "values") == (compat.MARK_ON, "C:/Накази/№ 2.docx")

    for child in tree.get_children():
        tree.delete(child)
    assert tree.get_children() == ()
    assert widget.topLevelItemCount() == 0

    with pytest.raises(ValueError):
        tree.insert("", compat.END, iid="same", values=("x",))
        tree.insert("", compat.END, iid="same", values=("y",))


def test_real_marked_paths_logic_runs_on_the_bridge(qt_app, tmp_path):
    generator = _load_generator()
    present = tmp_path / "Наказ № 1 від 01.09.2026.docx"
    present.write_bytes(b"")
    widget = QTreeWidget()
    tree = compat.TreeBridge(widget)
    generator.App._fill_orders_tree(None, tree, [str(present), str(tmp_path / "немає.docx")], "підказка")
    assert generator.App._marked_paths(tree) == [str(present)]

    widget.topLevelItem(0).setCheckState(0, Qt.CheckState.Unchecked)
    assert generator.App._marked_paths(tree) == []


def test_selection_api(qt_app):
    widget = QTreeWidget()
    widget.setSelectionMode(QTreeWidget.SelectionMode.ExtendedSelection)
    tree = compat.TreeBridge(widget)
    first = tree.insert("", compat.END, values=("1", "a"))
    second = tree.insert("", compat.END, values=("2", "b"))
    for child in tree.get_children():
        tree.selection_add(child)
    assert tree.selection() == (first, second)
    tree.selection_remove(first)
    assert tree.selection() == (second,)


def test_columns_heading_and_width(qt_app):
    widget = QTreeWidget()
    tree = compat.TreeBridge(widget)
    tree["columns"] = ["Пункт", "Причина"]
    assert tree["columns"] == ("Пункт", "Причина")
    assert widget.headerItem().text(1) == "Причина"
    tree.heading("Причина", text="Чому")
    tree.column("Пункт", width=222)
    assert widget.headerItem().text(1) == "Чому"
    assert widget.columnWidth(0) == 222


def test_notebook_bridge_index_of_selected_page(qt_app):
    tabs = QTabWidget()
    pages = [QWidget(), QWidget(), QWidget()]
    for index, page in enumerate(pages):
        tabs.addTab(page, str(index))
    notebook = compat.NotebookBridge(tabs)
    notebook.select(pages[2])
    assert notebook.index(notebook.select()) == 2
    notebook.select(0)
    assert notebook.index(notebook.select()) == 0


def test_button_state_combines_logic_and_busy_lock(qt_app):
    widget = compat.TkButton("Дія")
    widget.config(state=compat.DISABLED)
    assert not widget.isEnabled()
    widget.set_busy_locked(True)
    widget.config(state=compat.NORMAL)
    assert not widget.isEnabled()
    widget.set_busy_locked(False)
    assert widget.isEnabled()


def test_root_bridge_has_no_ttk_style_and_uses_clipboard(qt_app):
    window = QWidget()
    root = compat.RootBridge(window)
    assert not hasattr(root, "style")
    root.title("Генератор")
    root.geometry("640x480")
    assert window.windowTitle() == "Генератор"
    assert (window.width(), window.height()) == (640, 480)
    root.clipboard_clear()
    root.clipboard_append("журнал")
    assert QGuiApplication.clipboard().text() == "журнал"


def test_dialog_bridges_return_tk_shapes_on_cancel(qt_app, monkeypatch):
    filters = []
    monkeypatch.setattr(
        QFileDialog, "getOpenFileName", staticmethod(lambda *args: (filters.append(args[3]) or ("", "")))
    )
    monkeypatch.setattr(QFileDialog, "getOpenFileNames", staticmethod(lambda *args: ([], "")))
    monkeypatch.setattr(QFileDialog, "getExistingDirectory", staticmethod(lambda *args: ""))
    monkeypatch.setattr(QFileDialog, "getSaveFileName", staticmethod(lambda *args: ("C:/звіт", "")))

    assert compat.FileDialogBridge.askopenfilename(filetypes=[("Word files", "*.docx *.doc")]) == ""
    assert filters == ["Word files (*.docx *.doc)"]
    assert compat.FileDialogBridge.askopenfilenames() == ()
    assert compat.FileDialogBridge.askdirectory() == ""
    assert compat.FileDialogBridge.asksaveasfilename(defaultextension=".md") == "C:/звіт.md"

    shown = []
    monkeypatch.setattr(QMessageBox, "warning", staticmethod(lambda *args: shown.append(args[1:])))
    compat.MessageBoxBridge.showwarning("Помилка", "Оберіть наказ")
    assert shown == [("Помилка", "Оберіть наказ")]


@pytest.mark.parametrize(
    ("message", "level"),
    [
        ("  ПОМИЛКА (наказ.docx): Word не відповідає", "error"),
        ("УВАГА: 2 пункт(ів) без адресата.", "warning"),
        ("⚠️ Пунктів без адресата (пропущено): 2", "warning"),
        ("  ✓ 55 АК А0055: 3 стор., 6.8 с.", "success"),
        ("✅ Пунктів без адресата (пропущено): 0", "success"),
        ("Готовий до роботи.", "info"),
        ("Індексуємо абзаци оригіналу наказу...", "info"),
    ],
)
def test_log_levels(message, level):
    assert compat.classify_log_level(message) == level


def test_log_copy_text_has_time_but_no_level_tags():
    model = compat.LogModel(clock=lambda: "12:00:00")
    model.append("УВАГА: пункт без адресата")
    model.append("Збережено файл")
    assert model.plain_text() == "[12:00:00] УВАГА: пункт без адресата\n[12:00:00] Збережено файл"
    assert "[ГОТОВО]" not in model.plain_text()

    text = compat.TextBridge(model)
    text.delete("1.0", compat.END)
    assert text.get("1.0", compat.END) == ""
    text.insert(compat.END, "рядок\n")
    assert model.entries[-1].message == "рядок"
