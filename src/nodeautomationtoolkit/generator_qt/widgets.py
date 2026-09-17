"""Будівельні блоки інтерфейсу «Obsidian»: картки, бейджі, таблиці, журнал, рядок стану."""

from __future__ import annotations

import html
import os
import re
from typing import Callable, Iterable

import PySide6
from PySide6.QtCore import QRectF, Qt, qVersion
from PySide6.QtGui import QColor, QFont, QPainter, QPen
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QPlainTextEdit,
    QProgressBar,
    QStyle,
    QStyledItemDelegate,
    QStyleOptionViewItem,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from . import theme
from .compat import LogEntry, LogModel, TkButton, TreeBridge, is_placeholder_row

PILL_ROLE = Qt.ItemDataRole.UserRole + 1


# ─────────────────────────────────────────────────────────────────────────────
# Прив'язка віджетів до змінних логіки
# ─────────────────────────────────────────────────────────────────────────────
def bind_line_edit(edit: QLineEdit, variable) -> QLineEdit:
    """Поле й змінна логіки завжди показують одне значення."""

    def show_value(value) -> None:
        text = str(value)
        if edit.text() != text:
            edit.setText(text)
            edit.setCursorPosition(0)
        edit.setToolTip(text)

    show_value(variable.get())
    listener = variable.subscribe(show_value)
    edit.textEdited.connect(variable.set)
    edit.destroyed.connect(lambda *_args: variable.unsubscribe(listener))
    return edit


def bind_check_box(box: QCheckBox, variable) -> QCheckBox:
    box.setChecked(bool(variable.get()))

    def show_value(value) -> None:
        if box.isChecked() != bool(value):
            box.setChecked(bool(value))

    listener = variable.subscribe(show_value)
    box.toggled.connect(variable.set)
    box.destroyed.connect(lambda *_args: variable.unsubscribe(listener))
    return box


def bind_label(label: QLabel, variable, formatter: Callable[[object], str] = str) -> QLabel:
    label.setText(formatter(variable.get()))
    listener = variable.subscribe(lambda value: label.setText(formatter(value)))
    label.destroyed.connect(lambda *_args: variable.unsubscribe(listener))
    return label


def line_edit(variable, placeholder: str = "") -> QLineEdit:
    edit = QLineEdit()
    if placeholder:
        edit.setPlaceholderText(placeholder)
    return bind_line_edit(edit, variable)


def check_box(text: str, variable) -> QCheckBox:
    return bind_check_box(QCheckBox(text), variable)


def button(text: str, variant: str = "secondary", slot: Callable | None = None, tooltip: str = "") -> TkButton:
    widget = TkButton(text, variant)
    if slot is not None:
        widget.clicked.connect(lambda _checked=False: slot())
    if tooltip:
        widget.setToolTip(tooltip)
    return widget


def label(text: str = "", name: str = "") -> QLabel:
    widget = QLabel(text)
    if name:
        widget.setObjectName(name)
    return widget


def divider() -> QLabel:
    return label("|", "Divider")


# ─────────────────────────────────────────────────────────────────────────────
# Бейджі, картки, смуги
# ─────────────────────────────────────────────────────────────────────────────
class Badge(QLabel):
    def __init__(self, text: str = "", tone: str = "slate", parent: QWidget | None = None):
        super().__init__(text, parent)
        self.setProperty("tone", tone)

    def set_tone(self, tone: str) -> None:
        if self.property("tone") != tone:
            self.setProperty("tone", tone)
            theme.repolish(self)


class SectionCard(QFrame):
    """Картка розділу: крапка, заголовок великими, бейджі, дії праворуч, тіло."""

    def __init__(self, title: str, accent: str = "indigo", variant: str = "default", parent: QWidget | None = None):
        super().__init__(parent)
        self.setObjectName("SectionCard")
        self.setProperty("variant", variant)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(10, 8, 10, 10)
        outer.setSpacing(6)

        header = QHBoxLayout()
        header.setSpacing(8)
        dot = label("", "SectionDot")
        dot.setFixedSize(7, 7)
        dot.setProperty("accent", accent)
        self.title = label(title.upper(), "SectionTitle")
        self.title.setProperty("accent", accent)
        header.addWidget(dot, 0, Qt.AlignmentFlag.AlignVCenter)
        header.addWidget(self.title)
        self.badges = QHBoxLayout()
        self.badges.setSpacing(6)
        header.addLayout(self.badges)
        header.addStretch(1)
        self.actions = QHBoxLayout()
        self.actions.setSpacing(6)
        header.addLayout(self.actions)
        outer.addLayout(header)

        rule = QFrame()
        rule.setObjectName("SectionRule")
        rule.setFixedHeight(1)
        outer.addWidget(rule)

        self.body = QVBoxLayout()
        self.body.setSpacing(6)
        outer.addLayout(self.body, 1)

    def add_badge(self, badge: QWidget) -> QWidget:
        self.badges.addWidget(badge)
        return badge

    def add_action(self, widget: QWidget) -> QWidget:
        self.actions.addWidget(widget)
        return widget

    def add_divider(self) -> None:
        self.actions.addWidget(divider())


class ActionStrip(QFrame):
    """Смуга головних дій вкладки зі стислим станом пакета праворуч."""

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setObjectName("ActionStrip")
        self._layout = QHBoxLayout(self)
        self._layout.setContentsMargins(8, 6, 10, 6)
        self._layout.setSpacing(6)
        self.status = label("", "StripStatus")
        self.status.setTextFormat(Qt.TextFormat.RichText)
        self._finished = False

    def add(self, widget: QWidget) -> QWidget:
        self._layout.addWidget(widget)
        return widget

    def finish(self) -> "ActionStrip":
        if not self._finished:
            self._layout.addStretch(1)
            self._layout.addWidget(self.status)
            self._finished = True
        return self


def hint_bar(note: str = "") -> QFrame:
    bar = QFrame()
    bar.setObjectName("HintBar")
    layout = QHBoxLayout(bar)
    layout.setContentsMargins(8, 3, 6, 3)
    layout.setSpacing(6)
    layout.addWidget(label("⇪  Сюди можна перетягнути накази або папку (Drag & Drop)", "HintText"))
    layout.addStretch(1)
    if note:
        layout.addWidget(Badge(note, "amber"))
    return bar


ORDER_EXTENSIONS = (".docx", ".doc")


def dropped_order_paths(urls) -> list[str]:
    """Локальні файли наказів із перетягування: DOCX/DOC, без тимчасових «~$…» Word."""
    paths = []
    for url in urls:
        if not url.isLocalFile():
            continue
        path = url.toLocalFile()
        name = os.path.basename(path)
        if name.lower().endswith(ORDER_EXTENSIONS) and not name.startswith("~$") and os.path.isfile(path):
            paths.append(path)
    return paths


class OrderDropZone(QFrame):
    """Рамка, у яку перетягують наказ: `on_drop(шлях)` викликається з першим наказом.

    Перетягування, які рамка приймає, не доходять до обробника всього вікна —
    тож наказ у рамці не підміняє зразки й інші налаштування вкладки.
    """

    def __init__(self, text: str, on_drop: Callable[[str], None], parent: QWidget | None = None):
        super().__init__(parent)
        self.setObjectName("OrderDropZone")
        self.setProperty("active", False)
        self.setAcceptDrops(True)
        self.setMinimumHeight(64)
        self._on_drop = on_drop
        layout = QHBoxLayout(self)
        layout.setContentsMargins(12, 10, 12, 10)
        caption = label(text, "DropZoneText")
        caption.setAlignment(Qt.AlignmentFlag.AlignCenter)
        caption.setWordWrap(True)
        layout.addWidget(caption, 1)

    def _set_active(self, active: bool) -> None:
        self.setProperty("active", active)
        self.style().unpolish(self)
        self.style().polish(self)

    def dragEnterEvent(self, event) -> None:  # noqa: N802 - назва з Qt
        if dropped_order_paths(event.mimeData().urls()):
            event.acceptProposedAction()
            self._set_active(True)
        else:
            event.ignore()

    def dragMoveEvent(self, event) -> None:  # noqa: N802
        if dropped_order_paths(event.mimeData().urls()):
            event.acceptProposedAction()

    def dragLeaveEvent(self, event) -> None:  # noqa: N802
        self._set_active(False)
        super().dragLeaveEvent(event)

    def dropEvent(self, event) -> None:  # noqa: N802
        self._set_active(False)
        paths = dropped_order_paths(event.mimeData().urls())
        if not paths:
            event.ignore()
            return
        event.acceptProposedAction()
        self._on_drop(paths[0])


def run_params_card(
    title: str,
    *,
    executor_var,
    folder_var,
    on_pick_folder: Callable[[], None],
    executor_hint: str = "",
    people: Iterable[tuple] = (),
    options: Callable[[QHBoxLayout], None] | None = None,
) -> SectionCard:
    """Реквізити, що змінюються від наказу до наказу: виконавець, папка, особи."""
    card = SectionCard(title, accent="slate")
    grid = QGridLayout()
    grid.setHorizontalSpacing(8)
    grid.setVerticalSpacing(6)
    grid.addWidget(label("Виконавець:", "FieldLabel"), 0, 0)
    grid.addWidget(line_edit(executor_var, executor_hint), 0, 1)
    grid.addWidget(label("Папка результату:", "FieldLabel"), 0, 2)
    grid.addWidget(line_edit(folder_var), 0, 3)
    grid.addWidget(button("📂", "secondary", on_pick_folder, "Обрати папку результату"), 0, 4)
    grid.setColumnStretch(1, 2)
    grid.setColumnStretch(3, 3)
    card.body.addLayout(grid)

    for person_title, position_var, rank_var, name_var in people:
        card.body.addWidget(label(person_title, "PersonTitle"))
        row = QGridLayout()
        row.setHorizontalSpacing(8)
        row.addWidget(label("Посада:", "FieldLabel"), 0, 0)
        row.addWidget(line_edit(position_var), 0, 1)
        row.addWidget(label("Звання:", "FieldLabel"), 0, 2)
        row.addWidget(line_edit(rank_var), 0, 3)
        row.addWidget(label("ПІБ:", "FieldLabel"), 0, 4)
        row.addWidget(line_edit(name_var), 0, 5)
        row.setColumnStretch(1, 3)
        row.setColumnStretch(3, 1)
        row.setColumnStretch(5, 2)
        card.body.addLayout(row)

    if options is not None:
        options_row = QHBoxLayout()
        options_row.setSpacing(10)
        options(options_row)
        card.body.addLayout(options_row)
    return card


# ─────────────────────────────────────────────────────────────────────────────
# Таблиці
# ─────────────────────────────────────────────────────────────────────────────
_SELECTION = {
    "none": QAbstractItemView.SelectionMode.NoSelection,
    "single": QAbstractItemView.SelectionMode.SingleSelection,
    "extended": QAbstractItemView.SelectionMode.ExtendedSelection,
}


def make_table(
    columns: list[str],
    *,
    selection: str = "none",
    header_tone: str = "slate",
    stretch_column: int | None = None,
    widths: dict[int, int] | None = None,
    min_height: int = 110,
) -> QTreeWidget:
    table = QTreeWidget()
    table.setObjectName("DataTable")
    table.setProperty("headerTone", header_tone)
    table.setColumnCount(len(columns))
    table.setHeaderLabels(columns)
    table.setRootIsDecorated(False)
    table.setUniformRowHeights(True)
    table.setAlternatingRowColors(True)
    table.setAllColumnsShowFocus(True)
    table.setSelectionMode(_SELECTION[selection])
    table.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
    table.setTextElideMode(Qt.TextElideMode.ElideMiddle)
    table.setMinimumHeight(min_height)
    header = table.header()
    header.setSectionsMovable(False)
    header.setMinimumSectionSize(28)
    header.setDefaultAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
    header.setStretchLastSection(stretch_column is None)
    for index, width in (widths or {}).items():
        table.setColumnWidth(index, width)
    if stretch_column is not None:
        header.setSectionResizeMode(stretch_column, QHeaderView.ResizeMode.Stretch)
    return table


class PillDelegate(QStyledItemDelegate):
    """Малює значення комірки пігулкою, якщо в ній задано тон (`PILL_ROLE`)."""

    def paint(self, painter: QPainter, option: QStyleOptionViewItem, index) -> None:
        tone = index.data(PILL_ROLE)
        text = str(index.data(Qt.ItemDataRole.DisplayRole) or "")
        if not tone or not text:
            super().paint(painter, option, index)
            return
        panel = QStyleOptionViewItem(option)
        self.initStyleOption(panel, index)
        panel.text = ""
        widget = option.widget
        style = widget.style() if widget is not None else QApplication.style()
        style.drawControl(QStyle.ControlElement.CE_ItemViewItem, panel, painter, widget)

        background, border, foreground = theme.PILL_TONES.get(tone, theme.PILL_TONES["slate"])
        painter.save()
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        font = QFont(option.font)
        font.setPointSizeF(max(7.0, font.pointSizeF() - 0.5))
        painter.setFont(font)
        metrics = painter.fontMetrics()
        width = metrics.horizontalAdvance(text) + 14
        height = metrics.height() + 4
        rect = option.rect
        pill = QRectF(
            rect.x() + max(4.0, (rect.width() - width) / 2),
            rect.center().y() - height / 2 + 0.5,
            min(width, rect.width() - 8),
            height,
        )
        painter.setPen(QPen(QColor(border), 1))
        painter.setBrush(QColor(background))
        painter.drawRoundedRect(pill, 3, 3)
        painter.setPen(QColor(foreground))
        painter.drawText(pill, Qt.AlignmentFlag.AlignCenter, text)
        painter.restore()


def _mono_font(size: float = 8.5, bold: bool = False) -> QFont:
    font = QFont(theme.MONO)
    font.setStyleHint(QFont.StyleHint.Monospace)
    font.setPointSizeF(size)
    font.setBold(bold)
    return font


def _reset_item(item: QTreeWidgetItem, columns: int) -> None:
    for column in range(columns):
        item.setText(column, "")
        item.setData(column, Qt.ItemDataRole.FontRole, None)
        item.setData(column, Qt.ItemDataRole.ForegroundRole, None)
        item.setData(column, Qt.ItemDataRole.BackgroundRole, None)
        item.setData(column, PILL_ROLE, None)
        item.setToolTip(column, "")


def _mute_item(item: QTreeWidgetItem, columns: int) -> None:
    for column in range(columns):
        font = QFont(item.font(column))
        font.setItalic(True)
        item.setFont(column, font)
        item.setForeground(column, QColor(theme.C["text_faint"]))


def format_size(path: str) -> str:
    try:
        size = os.path.getsize(path)
    except OSError:
        return "—"
    return f"{max(1, round(size / 1024))} KB"


ORDER_COLUMNS = ["", "ТИП", "ФАЙЛ НАКАЗУ", "ДАТА / НОМЕР", "РОЗМІР", "СТАТУС"]
ORDER_WIDTHS = {0: 34, 1: 50, 3: 190, 4: 76, 5: 124}


def make_orders_renderer(metadata_from_filename: Callable[[str], tuple[str, str]]):
    """Рядок наказу: галочка, тип, назва, дата/номер, розмір, статус.

    Номер і дату беремо ЛИШЕ з назви файлу — як і вся логіка генератора
    (PROJECT_RULES 3.1), щоб таблиця не показувала інших реквізитів, ніж
    потраплять у документи.
    """
    columns = len(ORDER_COLUMNS)

    def render(_bridge: TreeBridge, item: QTreeWidgetItem, values: tuple) -> None:
        _reset_item(item, columns)
        if is_placeholder_row(values):
            item.setText(2, str(values[1]) if len(values) > 1 else "")
            _mute_item(item, columns)
            return
        path = str(values[1]) if len(values) > 1 else ""
        name = os.path.basename(path)
        stem, extension = os.path.splitext(name)
        number, date = metadata_from_filename(name)
        exists = os.path.isfile(path)

        item.setText(1, extension.lstrip(".").upper() or "—")
        item.setFont(1, _mono_font(8, bold=True))
        item.setForeground(1, QColor(theme.C["blue"]))

        item.setText(2, stem or name)
        item.setToolTip(2, path)
        item.setForeground(2, QColor(theme.C["text_strong"]))

        item.setText(3, " / ".join(part for part in (date, f"№ {number}" if number else "") if part) or "—")
        item.setFont(3, _mono_font())

        item.setText(4, format_size(path) if exists else "—")
        item.setFont(4, _mono_font())
        item.setForeground(4, QColor(theme.C["text_muted"]))
        item.setTextAlignment(4, Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)

        item.setText(5, "✓ Готовий" if exists else "⚠ Не знайдено")
        item.setData(5, PILL_ROLE, "emerald" if exists else "rose")

    return render


COPY_RESULT_COLUMNS = ["№", "НАЗВА ФАЙЛУ", "НОМЕР НАКАЗУ", "ДАТА НАКАЗУ", "СТОРІНОК", "АРК. ДЛЯ ДРУКУ", "ПОВНИЙ ШЛЯХ"]
COPY_RESULT_WIDTHS = {0: 42, 2: 118, 3: 100, 4: 92, 5: 118, 6: 340}


def render_copy_result(_bridge: TreeBridge, item: QTreeWidgetItem, values: tuple) -> None:
    """Сформований примірник: номер, назва, реквізити, сторінки, аркуші, шлях."""
    columns = len(COPY_RESULT_COLUMNS)
    _reset_item(item, columns)
    if is_placeholder_row(values):
        for column, value in enumerate(values[:columns]):
            item.setText(column, str(value))
        _mute_item(item, columns)
        return
    index, name, number, date, pages, sheets, path = (list(values) + [""] * columns)[:columns]

    item.setText(0, str(index))
    item.setFont(0, _mono_font(8.5, bold=True))
    item.setForeground(0, QColor(theme.C["teal_soft"]))
    item.setTextAlignment(0, Qt.AlignmentFlag.AlignCenter)

    item.setText(1, str(name))
    item.setForeground(1, QColor(theme.C["text_strong"]))
    item.setToolTip(1, str(path))

    item.setText(2, f"№ {number}" if str(number) not in ("", "—") else "—")
    item.setFont(2, _mono_font())
    item.setText(3, str(date) or "—")
    item.setFont(3, _mono_font())

    if str(pages) not in ("", "—"):
        item.setText(4, f"{pages} стор.")
        item.setData(4, PILL_ROLE, "slate")
    if str(sheets) not in ("", "—"):
        item.setText(5, f"{sheets} арк.")
        item.setData(5, PILL_ROLE, "indigo")

    item.setText(6, str(path))
    item.setFont(6, _mono_font(8))
    item.setForeground(6, QColor(theme.C["text_faint"]))
    item.setToolTip(6, str(path))


def render_analysis_row(_bridge: TreeBridge, item: QTreeWidgetItem, values: tuple) -> None:
    columns = max(1, item.treeWidget().columnCount() if item.treeWidget() else len(values))
    _reset_item(item, max(columns, len(values)))
    for column, value in enumerate(values):
        text = "-" if value in (None, "") else str(value)
        item.setText(column, text)
        item.setToolTip(column, text)
    if is_placeholder_row(values):
        _mute_item(item, len(values))
    elif values and str(values[0]).startswith("Управління ·"):
        for column in range(len(values)):
            item.setBackground(column, QColor("#052e2b"))
            item.setForeground(column, QColor(theme.C["emerald_text"]))
            font = QFont(item.font(column))
            font.setBold(True)
            item.setFont(column, font)


# ─────────────────────────────────────────────────────────────────────────────
# Журнал і рядок стану
# ─────────────────────────────────────────────────────────────────────────────
class LogConsole(QPlainTextEdit):
    """Консоль журналу: час, кольорова мітка рівня, текст повідомлення."""

    def __init__(self, model: LogModel, parent: QWidget | None = None):
        super().__init__(parent)
        self.setObjectName("LogConsole")
        self.setReadOnly(True)
        self.setMaximumBlockCount(4000)
        self.setLineWrapMode(QPlainTextEdit.LineWrapMode.WidgetWidth)
        self.setMinimumHeight(96)
        self.setFont(_mono_font(8.5))
        model.attach(self)

    def show_entries(self, entries: Iterable[LogEntry]) -> None:
        self.clear()
        for entry in entries:
            self.add_entry(entry)

    def add_entry(self, entry: LogEntry) -> None:
        tag, tag_color, text_color = theme.LOG_TONES.get(entry.level, theme.LOG_TONES["info"])
        message = html.escape(entry.message).replace("\n", "<br/>")
        # Відступи на початку рядків (статистика друку) зберігаємо.
        message = re.sub(
            r"(^|<br/>)( +)", lambda match: match.group(1) + "&nbsp;" * len(match.group(2)), message
        )
        self.appendHtml(
            f'<span style="color:{theme.C["text_dim"]}">[{entry.stamp}]</span> '
            f'<span style="color:{tag_color}; font-weight:600">[{tag}]</span> '
            f'<span style="color:{text_color}">{message}</span>'
        )
        bar = self.verticalScrollBar()
        bar.setValue(bar.maximum())


class StatusStrip(QFrame):
    """Нижній рядок стану: готовність, вибір, зразок, прогрес, версії."""

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setObjectName("StatusStrip")
        layout = QHBoxLayout(self)
        layout.setContentsMargins(10, 3, 10, 3)
        layout.setSpacing(8)

        self.dot = label("", "StatusDot")
        self.dot.setFixedSize(7, 7)
        self.state = label("Готово", "StatusState")
        self.selected = label("")
        self.template = label("")
        self.template.setTextFormat(Qt.TextFormat.RichText)
        layout.addWidget(self.dot, 0, Qt.AlignmentFlag.AlignVCenter)
        layout.addWidget(self.state)
        layout.addWidget(divider())
        layout.addWidget(self.selected)
        layout.addWidget(divider())
        layout.addWidget(self.template)
        layout.addStretch(1)

        self.caption = label("")
        self.progress = QProgressBar()
        self.progress.setObjectName("SlimProgress")
        self.progress.setRange(0, 1000)
        self.progress.setTextVisible(False)
        self.progress.setFixedWidth(180)
        self.progress.setFixedHeight(6)
        layout.addWidget(self.caption)
        layout.addWidget(self.progress, 0, Qt.AlignmentFlag.AlignVCenter)
        layout.addWidget(divider())
        layout.addWidget(label(f"PySide6 {PySide6.__version__} · Qt {qVersion()}  |  UTF-8", "StatusTech"))

    def set_busy(self, busy: bool) -> None:
        state = "busy" if busy else "ready"
        self.state.setText("Виконується…" if busy else "Готово")
        for widget in (self.dot, self.state):
            if widget.property("state") != state:
                widget.setProperty("state", state)
                theme.repolish(widget)

    def set_progress(self, caption: str | None = None, percent: float | None = None) -> None:
        if caption is not None:
            self.caption.setText(str(caption))
        if percent is not None:
            self.progress.setValue(int(max(0.0, min(100.0, float(percent))) * 10))


def plural_files(count: int) -> str:
    tail10, tail100 = count % 10, count % 100
    if tail10 == 1 and tail100 != 11:
        word = "файл"
    elif 2 <= tail10 <= 4 and not 12 <= tail100 <= 14:
        word = "файли"
    else:
        word = "файлів"
    return f"{count} {word}"
