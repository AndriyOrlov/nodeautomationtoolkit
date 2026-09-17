"""Міст між логікою генератора (написаною під Tk) і Qt-оболонкою.

Логіка в `generate_extracts.App` звертається до інтерфейсу через вузьку
поверхню Tk: змінні з `.get()/.set()`, `messagebox`, `filedialog`,
Treeview-подібні списки, текст журналу та `config(state=...)` кнопок.
Тут ця сама поверхня реалізована поверх Qt. Генерація працює без жодної
зміни, а нова змінна в `App.__init__` автоматично з'являється й у Qt.
"""

from __future__ import annotations

import itertools
import os
import re
import time
import types
from typing import Any, Callable, Iterable, NamedTuple

from PySide6.QtCore import Qt
from PySide6.QtGui import QGuiApplication
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QMessageBox,
    QPushButton,
    QTabWidget,
    QTreeWidget,
    QTreeWidgetItem,
    QWidget,
)

END = "end"
INSERT = "insert"
WORD = "word"
SEL = "sel"
DISABLED = "disabled"
NORMAL = "normal"
MARK_ON = "☑"
MARK_OFF = "☐"
_MARKS = (MARK_ON, MARK_OFF)


def pump_events() -> None:
    """Дає Qt перемалювати вікно посеред довгої синхронної роботи з Word."""
    app = QApplication.instance()
    if app is not None:
        app.processEvents()


def sleep_responsive(seconds: float) -> None:
    """Пауза, під час якої вікно лишається живим (режим превʼю)."""
    deadline = time.monotonic() + max(0.0, float(seconds))
    while (remaining := deadline - time.monotonic()) > 0:
        time.sleep(min(0.05, remaining))
        pump_events()


# ─────────────────────────────────────────────────────────────────────────────
# Змінні з API tkinter
# ─────────────────────────────────────────────────────────────────────────────
class _Variable:
    _default: Any = ""

    def __init__(self, master=None, value=None, name=None):
        self._value = self._coerce(self._default if value is None else value)
        self._listeners: list[Callable[[Any], None]] = []

    def _coerce(self, value):
        return value

    def get(self):
        return self._value

    def set(self, value) -> None:
        value = self._coerce(value)
        if value == self._value:
            return
        self._value = value
        for listener in tuple(self._listeners):
            try:
                listener(value)
            except RuntimeError as error:
                # Прив'язаний віджет уже знищено (закрите вікно зразків) —
                # такого слухача просто прибираємо. Інші помилки не ковтаємо.
                if "deleted" not in str(error):
                    raise
                self.unsubscribe(listener)

    def subscribe(self, listener: Callable[[Any], None]) -> Callable[[Any], None]:
        self._listeners.append(listener)
        return listener

    def unsubscribe(self, listener: Callable[[Any], None]) -> None:
        if listener in self._listeners:
            self._listeners.remove(listener)

    def trace_add(self, _mode, callback) -> str:
        self.subscribe(lambda _value: callback("", "", "write"))
        return f"trace{id(callback)}"


class StringVar(_Variable):
    _default = ""

    def _coerce(self, value):
        return "" if value is None else str(value)


class BooleanVar(_Variable):
    _default = False

    def _coerce(self, value):
        if isinstance(value, str):
            return value.strip().casefold() in ("1", "true", "yes", "on")
        return bool(value)


class DoubleVar(_Variable):
    _default = 0.0

    def _coerce(self, value):
        return float(value)


class IntVar(_Variable):
    _default = 0

    def _coerce(self, value):
        return int(float(value))


def make_tk_namespace() -> types.SimpleNamespace:
    """Те, що логіка генератора бере з модуля `tkinter` під час роботи."""
    return types.SimpleNamespace(
        StringVar=StringVar,
        BooleanVar=BooleanVar,
        DoubleVar=DoubleVar,
        IntVar=IntVar,
        END=END,
        INSERT=INSERT,
        WORD=WORD,
        SEL=SEL,
    )


# ─────────────────────────────────────────────────────────────────────────────
# Діалоги
# ─────────────────────────────────────────────────────────────────────────────
_dialog_parent: QWidget | None = None


def set_dialog_parent(widget: QWidget | None) -> None:
    """Вікно, над яким з'являються діалоги, коли активного вікна немає."""
    global _dialog_parent
    _dialog_parent = widget


def _parent() -> QWidget | None:
    app = QApplication.instance()
    active = app.activeWindow() if app is not None else None
    return active or _dialog_parent


def filter_from_filetypes(filetypes: Iterable | None) -> str:
    """`[("Word files", "*.docx *.doc")]` → `"Word files (*.docx *.doc)"`."""
    parts = []
    for label, patterns in filetypes or ():
        if isinstance(patterns, (list, tuple)):
            patterns = " ".join(patterns)
        parts.append(f"{label} ({patterns})")
    return ";;".join(parts)


class MessageBoxBridge:
    """`tkinter.messagebox` поверх QMessageBox."""

    @staticmethod
    def showinfo(title="", message="", **_options):
        QMessageBox.information(_parent(), str(title), str(message))
        return "ok"

    @staticmethod
    def showwarning(title="", message="", **_options):
        QMessageBox.warning(_parent(), str(title), str(message))
        return "ok"

    @staticmethod
    def showerror(title="", message="", **_options):
        QMessageBox.critical(_parent(), str(title), str(message))
        return "ok"

    @staticmethod
    def askyesno(title="", message="", **_options) -> bool:
        answer = QMessageBox.question(_parent(), str(title), str(message))
        return answer == QMessageBox.StandardButton.Yes

    @staticmethod
    def askokcancel(title="", message="", **_options) -> bool:
        answer = QMessageBox.question(
            _parent(),
            str(title),
            str(message),
            QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel,
        )
        return answer == QMessageBox.StandardButton.Ok


class FileDialogBridge:
    """`tkinter.filedialog` поверх QFileDialog. Скасування — порожній результат, як у Tk."""

    @staticmethod
    def askopenfilename(title="Оберіть файл", filetypes=None, initialdir="", **_options) -> str:
        path, _selected = QFileDialog.getOpenFileName(
            _parent(), str(title), str(initialdir or ""), filter_from_filetypes(filetypes)
        )
        return path or ""

    @staticmethod
    def askopenfilenames(title="Оберіть файли", filetypes=None, initialdir="", **_options) -> tuple:
        paths, _selected = QFileDialog.getOpenFileNames(
            _parent(), str(title), str(initialdir or ""), filter_from_filetypes(filetypes)
        )
        return tuple(paths or ())

    @staticmethod
    def askdirectory(title="Оберіть теку", initialdir="", **_options) -> str:
        return QFileDialog.getExistingDirectory(_parent(), str(title), str(initialdir or "")) or ""

    @staticmethod
    def asksaveasfilename(
        title="Зберегти як", filetypes=None, defaultextension="", initialfile="", initialdir="", **_options
    ) -> str:
        start = os.path.join(str(initialdir or ""), str(initialfile or "")) if initialfile else str(initialdir or "")
        path, _selected = QFileDialog.getSaveFileName(
            _parent(), str(title), start, filter_from_filetypes(filetypes)
        )
        if path and defaultextension and not os.path.splitext(path)[1]:
            path += defaultextension
        return path or ""


# ─────────────────────────────────────────────────────────────────────────────
# Вікно, вкладки, кнопки
# ─────────────────────────────────────────────────────────────────────────────
class RootBridge:
    """Замінник `tk.Tk` для логіки генератора."""

    def __init__(self, window: QWidget):
        self.window = window

    def title(self, text: str | None = None):
        if text is None:
            return self.window.windowTitle()
        self.window.setWindowTitle(str(text))
        return None

    def geometry(self, spec: str | None = None):
        if spec is None:
            return f"{self.window.width()}x{self.window.height()}"
        match = re.match(r"^\s*(\d+)x(\d+)", str(spec))
        if match:
            self.window.resize(int(match.group(1)), int(match.group(2)))
        return None

    def minsize(self, width=None, height=None):
        if width is not None and height is not None:
            self.window.setMinimumSize(int(width), int(height))

    def update(self) -> None:
        pump_events()

    update_idletasks = update

    @staticmethod
    def clipboard_clear() -> None:
        QGuiApplication.clipboard().setText("")

    @staticmethod
    def clipboard_append(text: str) -> None:
        clipboard = QGuiApplication.clipboard()
        clipboard.setText(clipboard.text() + str(text))


class NotebookBridge:
    """`ttk.Notebook` поверх QTabWidget: `select()`, `select(сторінка)`, `index(...)`."""

    def __init__(self, tabs: QTabWidget):
        self.widget = tabs

    def select(self, tab_id=None):
        if tab_id is None:
            return self.widget.currentWidget()
        if isinstance(tab_id, int):
            self.widget.setCurrentIndex(tab_id)
        elif isinstance(tab_id, QWidget):
            self.widget.setCurrentWidget(tab_id)
        return None

    def index(self, tab_id) -> int:
        if isinstance(tab_id, int):
            return tab_id
        if tab_id in (None, "current"):
            return self.widget.currentIndex()
        return self.widget.indexOf(tab_id)

    def tabs(self) -> list[QWidget]:
        return [self.widget.widget(index) for index in range(self.widget.count())]


class TkButton(QPushButton):
    """Кнопка, яку логіка вмикає/вимикає через `config(state=...)`.

    Стан від логіки й блокування на час пакета — два незалежні прапорці.
    Інакше внутрішній прогін повного циклу, завершуючись, «вмикав» би
    кнопки, поки зовнішній пакет ще працює.
    """

    def __init__(self, text: str = "", variant: str = "secondary", parent: QWidget | None = None):
        super().__init__(text, parent)
        self.setProperty("variant", variant)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self._logic_enabled = True
        self._busy_locked = False

    def config(self, **options) -> None:
        if "state" in options:
            self._logic_enabled = str(options["state"]) != DISABLED
            self._apply_enabled()
        if "text" in options:
            self.setText(str(options["text"]))

    configure = config

    def set_busy_locked(self, locked: bool) -> None:
        self._busy_locked = bool(locked)
        self._apply_enabled()

    def _apply_enabled(self) -> None:
        super().setEnabled(self._logic_enabled and not self._busy_locked)


# ─────────────────────────────────────────────────────────────────────────────
# Журнал
# ─────────────────────────────────────────────────────────────────────────────
_ERROR_MARKERS = ("помилка", "❌", "traceback", "exception")
_WARNING_MARKERS = ("увага", "⚠", "не вдалося", "перервано")
_SUCCESS_MARKERS = ("✓", "✅", "успішно", "збережено", "завершено", "сформовано")


def classify_log_level(message: str) -> str:
    """Рівень рядка журналу — лише для кольору; сам текст не змінюється."""
    text = str(message).casefold()
    for level, markers in (
        ("error", _ERROR_MARKERS),
        ("warning", _WARNING_MARKERS),
        ("success", _SUCCESS_MARKERS),
    ):
        if any(marker in text for marker in markers):
            return level
    return "info"


class LogEntry(NamedTuple):
    stamp: str
    level: str
    message: str


class LogModel:
    """Записи журналу й усі його відображення (одна модель — кілька вкладок)."""

    def __init__(self, clock: Callable[[], str] | None = None):
        self.entries: list[LogEntry] = []
        self._views: list = []
        self._clock = clock or (lambda: time.strftime("%H:%M:%S"))

    def attach(self, view) -> None:
        self._views.append(view)
        view.show_entries(self.entries)
        view.destroyed.connect(lambda *_args, gone=view: self._detach(gone))

    def _detach(self, view) -> None:
        if view in self._views:
            self._views.remove(view)

    def append(self, message) -> LogEntry:
        entry = LogEntry(self._clock(), classify_log_level(message), str(message))
        self.entries.append(entry)
        for view in tuple(self._views):
            view.add_entry(entry)
        return entry

    def clear(self) -> None:
        self.entries.clear()
        for view in tuple(self._views):
            view.show_entries([])

    def plain_text(self) -> str:
        """Текст для копіювання: час і повідомлення, без візуальних міток рівня.

        Мітки на кшталт «УВАГА» знеособлення сприйняло б як назви й замінило б
        на ярлики, тож у буфер іде рівно те, що писала логіка.
        """
        return "\n".join(f"[{entry.stamp}] {entry.message}" for entry in self.entries)


class TextBridge:
    """`tk.Text` журналу для логіки: `insert/see/delete/get`."""

    def __init__(self, model: LogModel):
        self.model = model

    def insert(self, _index, text) -> None:
        text = str(text)
        if text.endswith("\n"):
            text = text[:-1]
        self.model.append(text)

    def see(self, _index) -> None:
        return None

    def delete(self, *_indices) -> None:
        self.model.clear()

    def get(self, *_indices) -> str:
        return self.model.plain_text()

    def tag_add(self, *_args) -> None:
        return None

    def event_generate(self, *_args) -> None:
        return None


# ─────────────────────────────────────────────────────────────────────────────
# Treeview
# ─────────────────────────────────────────────────────────────────────────────
Renderer = Callable[["TreeBridge", QTreeWidgetItem, tuple], None]


def plain_renderer(_bridge: "TreeBridge", item: QTreeWidgetItem, values: tuple) -> None:
    for column, value in enumerate(values):
        item.setText(column, "-" if value in (None, "") else str(value))


def is_placeholder_row(values: Iterable) -> bool:
    """Рядок-підказка («— Очікується …»), а не справжні дані."""
    values = tuple(values)
    return bool(values) and (
        values[0] == "—" or any(isinstance(value, str) and value.startswith("— ") for value in values)
    )


def _flatten(ids: Iterable) -> list[str]:
    flat: list[str] = []
    for value in ids:
        if isinstance(value, (list, tuple)):
            flat.extend(str(inner) for inner in value)
        else:
            flat.append(str(value))
    return flat


class TreeBridge:
    """Tk-подібний Treeview над QTreeWidget.

    Логіка зберігає в першій колонці позначку «☑/☐» і читає її назад через
    `item(iid, "values")`. У Qt ця позначка — справжня галочка в рядку, а
    `values()` повертає «☑/☐» за її поточним станом. Тому
    `App._marked_paths` і `_fill_orders_tree` працюють без змін.
    """

    def __init__(self, widget: QTreeWidget, renderer: Renderer | None = None):
        self.widget = widget
        self._renderer = renderer or plain_renderer
        self._items: dict[str, QTreeWidgetItem] = {}
        self._values: dict[str, tuple] = {}
        self._ids = itertools.count(1)
        self._columns: list[str] = []
        self._rendering = False
        self.changed_callbacks: list[Callable[[], None]] = []
        self.mark_callbacks: list[Callable[[], None]] = []
        widget.itemChanged.connect(self._on_item_changed)

    # ── API ttk.Treeview, який використовує логіка ────────────────────────
    def get_children(self, item: str = "") -> tuple[str, ...]:
        if item:
            return ()
        return tuple(
            self.widget.topLevelItem(row).data(0, Qt.ItemDataRole.UserRole)
            for row in range(self.widget.topLevelItemCount())
        )

    def insert(self, parent, index, iid=None, values=(), **_options) -> str:
        iid = f"I{next(self._ids):03d}" if iid is None else str(iid)
        if iid in self._items:
            raise ValueError(f"Рядок {iid!r} уже існує")
        item = QTreeWidgetItem()
        item.setData(0, Qt.ItemDataRole.UserRole, iid)
        self._items[iid] = item
        self._values[iid] = tuple(values)
        # Малюємо ДО вставки: поки рядка немає в таблиці, зміна галочки не
        # породжує сигнал itemChanged.
        self._render(iid)
        count = self.widget.topLevelItemCount()
        position = count if index in (END, None) else max(0, min(int(index), count))
        self.widget.insertTopLevelItem(position, item)
        self._notify(self.changed_callbacks)
        return iid

    def delete(self, *iids) -> None:
        for iid in _flatten(iids):
            item = self._items.pop(iid, None)
            self._values.pop(iid, None)
            if item is None:
                continue
            row = self.widget.indexOfTopLevelItem(item)
            if row >= 0:
                self.widget.takeTopLevelItem(row)
        self._notify(self.changed_callbacks)

    def item(self, iid, option=None, **options):
        iid = str(iid)
        if iid not in self._items:
            raise KeyError(iid)
        if "values" in options:
            self._values[iid] = tuple(options["values"])
            self._render(iid)
            self._notify(self.changed_callbacks)
        if option == "values":
            return self.values(iid)
        if option is None and not options:
            return {"text": "", "values": list(self.values(iid))}
        return None

    def values(self, iid: str) -> tuple:
        values = list(self._values[str(iid)])
        if values and values[0] in _MARKS:
            checked = self._items[str(iid)].checkState(0) == Qt.CheckState.Checked
            values[0] = MARK_ON if checked else MARK_OFF
        return tuple(values)

    def selection(self) -> tuple[str, ...]:
        return tuple(iid for iid in self.get_children() if self._items[iid].isSelected())

    def selection_add(self, *iids) -> None:
        for iid in _flatten(iids):
            if iid in self._items:
                self._items[iid].setSelected(True)

    def selection_remove(self, *iids) -> None:
        for iid in _flatten(iids):
            if iid in self._items:
                self._items[iid].setSelected(False)

    def selection_set(self, *iids) -> None:
        self.widget.clearSelection()
        self.selection_add(*iids)

    def see(self, iid) -> None:
        item = self._items.get(str(iid))
        if item is not None:
            self.widget.scrollToItem(item)

    def identify_row(self, y) -> str:
        item = self.widget.itemAt(0, int(y))
        return "" if item is None else str(item.data(0, Qt.ItemDataRole.UserRole))

    def bind(self, *_args, **_options) -> None:
        return None

    def configure(self, **_options) -> None:
        return None

    config = configure

    def __setitem__(self, key, value) -> None:
        if key == "columns":
            self._columns = [str(column) for column in value]
            self.widget.setColumnCount(max(1, len(self._columns)))
            self.widget.setHeaderLabels(self._columns)

    def __getitem__(self, key):
        if key == "columns":
            return tuple(self._columns)
        raise KeyError(key)

    def heading(self, column, text=None, **_options) -> None:
        index = self._column_index(column)
        if index is not None and text is not None:
            self.widget.headerItem().setText(index, str(text))

    def column(self, column, width=None, **_options) -> None:
        index = self._column_index(column)
        if index is not None and width:
            self.widget.setColumnWidth(index, int(width))

    # ── для оболонки ──────────────────────────────────────────────────────
    def rows(self) -> list[tuple]:
        return [self.values(iid) for iid in self.get_children()]

    def _column_index(self, column) -> int | None:
        if isinstance(column, int):
            return column
        if column in self._columns:
            return self._columns.index(column)
        return None

    def _render(self, iid: str) -> None:
        item = self._items[iid]
        values = self._values[iid]
        self._rendering = True
        try:
            if values and values[0] in _MARKS:
                item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                item.setCheckState(
                    0, Qt.CheckState.Checked if values[0] == MARK_ON else Qt.CheckState.Unchecked
                )
            else:
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsUserCheckable)
                item.setData(0, Qt.ItemDataRole.CheckStateRole, None)
            self._renderer(self, item, values)
        finally:
            self._rendering = False

    def _on_item_changed(self, item: QTreeWidgetItem, column: int) -> None:
        if self._rendering or column != 0:
            return
        iid = item.data(0, Qt.ItemDataRole.UserRole)
        values = self._values.get(iid, ())
        if values and values[0] in _MARKS:
            self._notify(self.mark_callbacks)
            self._notify(self.changed_callbacks)

    @staticmethod
    def _notify(callbacks: list[Callable[[], None]]) -> None:
        for callback in tuple(callbacks):
            callback()
