"""Вікно індексатора посад (PySide6, тема «Obsidian» генератора).

Обирається папка з наказами і папка, куди зберігати індекс. Індексування йде
у фоновому потоці; після нього — список усіх знайдених посад із фільтрами і
перелік наказів, де трапляється вибрана посада.
"""

from __future__ import annotations

import csv
import os
import sys
from pathlib import Path

from PySide6.QtCore import QSettings, Qt, QThread, Signal
from PySide6.QtGui import QGuiApplication
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLineEdit,
    QMainWindow,
    QSplitter,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from ..generator_qt import compat, theme
from ..generator_qt.widgets import (
    ActionStrip,
    Badge,
    LogConsole,
    SectionCard,
    StatusStrip,
    button,
    label,
    make_table,
)
from . import store
from .extractor import ROLE_CURRENT, ROLE_TARGET
from .position_dictionary import ALL_POSITIONS_FILE, load_mo317, load_vos_names, strip_brackets
from .unit_split import MARKERS_FILE

SORT_ROLE = Qt.ItemDataRole.UserRole + 7
ID_ROLE = Qt.ItemDataRole.UserRole + 8

POSITION_COLUMNS = [
    "Посада", "Довідник", "Згадок", "Наказів", "Частин", "Займана", "Призначення", "Перша дата", "Остання дата",
    "Код посади", "ВОС",
]
POSITION_WIDTHS = {1: 76, 2: 62, 3: 62, 4: 58, 5: 66, 6: 86, 7: 84, 8: 90, 9: 80, 10: 140}
OCCURRENCE_COLUMNS = ["Дата", "№ наказу", "§", "Роль", "Частина", "Файл"]
OCCURRENCE_WIDTHS = {0: 84, 1: 80, 2: 40, 3: 92, 4: 260}

ROLE_FILTERS = (("Усі ролі", None), ("Лише займані", ROLE_CURRENT), ("Лише призначення", ROLE_TARGET))
KNOWN_FILTERS = (("Усі посади", None), ("Лише з довідника", 1), ("Не знайдені в довіднику", 0))


def _mo317_columns(info) -> tuple[str, str, str]:
    """(коди посад, ВОС коротко, підказка) з переліку МО № 317 для рядка списку."""
    if info is None:
        return "", "", ""
    codes = ", ".join(sorted(code for code in info.codes if code))
    if info.all_vos and not info.vos:
        vos = "усі ВОС"
    elif len(info.vos) > 6:
        vos = f"{len(info.vos)} ВОС" + (" (+ усі)" if info.all_vos else "")
    else:
        vos = ", ".join(sorted(info.vos)) + (" + усі" if info.all_vos else "")
    names = load_vos_names()
    lines = [f"Наказ МО № 317 · звання: {'; '.join(sorted(info.ranks)) or '—'}"]
    for row in info.rows:
        lines.append(
            f"код {row.get('Код посади', '')}, розряд {row.get('Тарифний розряд', '') or '—'}: "
            f"{row.get('ВОС (як у переліку)', '')}"
        )
    for number in sorted(info.vos)[:25]:
        lines.append(f"  {number} — {names.get(number, '(назви в переліку ВОС немає)')}")
    if len(info.vos) > 25:
        lines.append(f"  … ще {len(info.vos) - 25}")
    return codes, vos, "\n".join(lines)


def _ua_date(iso: str) -> str:
    return f"{iso[8:10]}.{iso[5:7]}.{iso[0:4]}" if len(iso) == 10 else ""


class SortableItem(QTreeWidgetItem):
    """Числа й дати сортуються за значенням, а не як текст."""

    def __lt__(self, other: QTreeWidgetItem) -> bool:
        column = self.treeWidget().sortColumn() if self.treeWidget() else 0
        mine, theirs = self.data(column, SORT_ROLE), other.data(column, SORT_ROLE)
        if mine is not None and theirs is not None:
            return mine < theirs
        return self.text(column).casefold() < other.text(column).casefold()


class IndexWorker(QThread):
    logged = Signal(str)
    progressed = Signal(int, int, str)
    finished_with = Signal(object, str)

    def __init__(self, orders_folder: str, index_folder: str, recursive: bool, force: bool):
        super().__init__()
        self.orders_folder = orders_folder
        self.index_folder = index_folder
        self.recursive = recursive
        self.force = force
        self._cancel = False

    def cancel(self) -> None:
        self._cancel = True

    def run(self) -> None:
        try:
            summary = store.build_index(
                self.orders_folder,
                self.index_folder,
                recursive=self.recursive,
                force=self.force,
                log=self.logged.emit,
                progress=self.progressed.emit,
                cancelled=lambda: self._cancel,
            )
            self.finished_with.emit(summary, "")
        except Exception as error:
            self.finished_with.emit(None, f"{type(error).__name__}: {error}")


class IndexerWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Індексатор посад у наказах — тестова версія")
        self.resize(1320, 860)
        self.settings = QSettings("NodeAutomationToolkit", "OrderPositionIndexer")
        self.log_model = compat.LogModel()
        self.worker: IndexWorker | None = None
        self.rows: list[store.PositionRow] = []

        central = QWidget()
        central.setObjectName("AppRoot")
        root = QVBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)
        root.addWidget(self._build_header())

        page = QWidget()
        page.setObjectName("TabPage")
        body = QVBoxLayout(page)
        body.setContentsMargins(10, 10, 10, 10)
        body.setSpacing(8)
        body.addWidget(self._build_folders_card())
        body.addWidget(self._build_actions())

        results = QSplitter(Qt.Orientation.Horizontal)
        results.addWidget(self._build_positions_card())
        results.addWidget(self._build_occurrences_card())
        results.setStretchFactor(0, 3)
        results.setStretchFactor(1, 2)

        vertical = QSplitter(Qt.Orientation.Vertical)
        vertical.addWidget(results)
        vertical.addWidget(self._build_log_card())
        vertical.setStretchFactor(0, 4)
        vertical.setStretchFactor(1, 1)
        body.addWidget(vertical, 1)
        root.addWidget(page, 1)

        self.status = StatusStrip()
        root.addWidget(self.status)
        self.setCentralWidget(central)

        self._restore_settings()
        self.reload_positions()

    # ── Побудова ─────────────────────────────────────────────────────────
    def _build_header(self) -> QFrame:
        bar = QFrame()
        bar.setObjectName("TitleBar")
        layout = QHBoxLayout(bar)
        layout.setContentsMargins(12, 6, 12, 6)
        layout.setSpacing(8)
        logo = label("🗂", "AppLogo")
        logo.setFixedSize(20, 20)
        logo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(logo)
        layout.addWidget(label("Node Automation Toolkit", "AppName"))
        layout.addWidget(label("|", "Divider"))
        layout.addWidget(label("Індексатор посад у наказах по особовому складу", "AppSubtitle"))
        layout.addWidget(Badge("тестова версія", "amber"))
        layout.addStretch(1)
        layout.addWidget(Badge("лише локально · без мережі", "teal"))
        return bar

    def _build_folders_card(self) -> SectionCard:
        card = SectionCard("Папки", accent="slate")
        grid = QGridLayout()
        grid.setHorizontalSpacing(8)
        grid.setVerticalSpacing(6)

        self.orders_edit = QLineEdit()
        self.orders_edit.setPlaceholderText("Папка, де лежать накази (.docx, .doc, .rtf)")
        self.index_edit = QLineEdit()
        self.index_edit.setPlaceholderText("Папка, куди зберегти індекс")
        self.index_edit.editingFinished.connect(self.reload_positions)

        grid.addWidget(label("Папка з наказами:", "FieldLabel"), 0, 0)
        grid.addWidget(self.orders_edit, 0, 1)
        grid.addWidget(button("📂", "secondary", self.pick_orders_folder, "Обрати папку з наказами"), 0, 2)
        grid.addWidget(label("Папка індексу:", "FieldLabel"), 1, 0)
        grid.addWidget(self.index_edit, 1, 1)
        grid.addWidget(button("📂", "secondary", self.pick_index_folder, "Обрати папку для індексу"), 1, 2)
        grid.addWidget(
            button("Відкрити", "ghost", self.open_index_folder, "Відкрити папку індексу в Провіднику"), 1, 3
        )
        grid.setColumnStretch(1, 1)
        card.body.addLayout(grid)

        options = QHBoxLayout()
        self.recursive_box = QCheckBox("Разом із підпапками")
        self.recursive_box.setChecked(True)
        options.addWidget(self.recursive_box)
        options.addStretch(1)
        options.addWidget(
            label(f"Індекс — один файл «{store.INDEX_FILENAME}». Накази лише читаються.", "MutedLabel")
        )
        card.body.addLayout(options)
        return card

    def _build_actions(self) -> ActionStrip:
        strip = ActionStrip()
        self.run_button = strip.add(button("▶  Індексувати", "run", lambda: self.start_indexing(False),
                                           "Додати нові й змінені накази; незмінені пропускаються"))
        self.force_button = strip.add(button("↻  Переіндексувати все", "secondary", lambda: self.start_indexing(True),
                                             "Очистити індекс і прочитати всі накази заново"))
        self.stop_button = strip.add(button("⏹  Зупинити", "rose", self.stop_indexing))
        self.stop_button.setEnabled(False)
        strip.add(label("|", "Divider"))
        strip.add(button("✎  Довідник посад", "amber", self.open_positions_extra,
                         "Словник усіх посад (CSV: називний, родовий, давальний, орудний, коди МО 317). "
                         "Свій рядок позначте в «Джерело» словом «вручну»; після правки — «Переіндексувати все»"))
        strip.add(button("✎  Ознаки частин", "amber", self.open_unit_markers,
                         "Список, з чого в тексті посади починається частина; після правки — «Переіндексувати все»"))
        strip.add(button("⧉  Копіювати список", "ghost", self.copy_positions, "Скопіювати видимі посади"))
        strip.add(button("⭳  Експорт CSV", "ghost", self.export_csv, "Зберегти видимі посади в папку індексу"))
        strip.finish()
        self.strip = strip
        return strip

    def _build_positions_card(self) -> SectionCard:
        card = SectionCard("Проіндексовані посади", accent="indigo", variant="results")
        self.count_badge = card.add_badge(Badge("0", "indigo"))

        filters = QHBoxLayout()
        filters.setSpacing(6)
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("🔍  Пошук за частиною назви посади…")
        self.search_edit.setClearButtonEnabled(True)
        self.search_edit.textChanged.connect(self.apply_filter)
        self.role_combo = QComboBox()
        for caption, _role in ROLE_FILTERS:
            self.role_combo.addItem(caption)
        self.role_combo.currentIndexChanged.connect(self.apply_filter)
        self.known_combo = QComboBox()
        for caption, _known in KNOWN_FILTERS:
            self.known_combo.addItem(caption)
        self.known_combo.setToolTip(
            "Посади, яких немає в довіднику, варто дописати в «Довідник посад» — тоді вони зведуться в один рядок"
        )
        self.known_combo.currentIndexChanged.connect(self.apply_filter)
        filters.addWidget(self.search_edit, 1)
        filters.addWidget(self.role_combo)
        filters.addWidget(self.known_combo)
        card.body.addLayout(filters)

        self.positions_table = make_table(
            POSITION_COLUMNS, selection="single", header_tone="indigo", stretch_column=0,
            widths=POSITION_WIDTHS, min_height=240,
        )
        self.positions_table.setTextElideMode(Qt.TextElideMode.ElideRight)
        self.positions_table.setSortingEnabled(True)
        self.positions_table.sortByColumn(2, Qt.SortOrder.DescendingOrder)
        self.positions_table.itemSelectionChanged.connect(self.show_occurrences)
        card.body.addWidget(self.positions_table, 1)
        return card

    def _build_occurrences_card(self) -> SectionCard:
        card = SectionCard("Де трапляється", accent="teal", variant="results")
        self.occurrence_badge = card.add_badge(Badge("—", "teal"))
        self.selected_label = label("Оберіть посаду зліва", "MutedLabel")
        self.selected_label.setWordWrap(True)
        card.body.addWidget(self.selected_label)
        self.occurrences_table = make_table(
            OCCURRENCE_COLUMNS, selection="single", header_tone="teal", stretch_column=5,
            widths=OCCURRENCE_WIDTHS, min_height=240,
        )
        self.occurrences_table.setSortingEnabled(True)
        self.occurrences_table.itemDoubleClicked.connect(self.open_occurrence_file)
        card.body.addWidget(self.occurrences_table, 1)
        card.body.addWidget(
            label("Наведіть на частину — повний текст із наказу. Подвійний клік — відкрити наказ.", "MutedLabel")
        )
        return card

    def _build_log_card(self) -> SectionCard:
        card = SectionCard("Журнал", accent="slate", variant="log")
        card.add_action(button("Очистити", "rose", self.log_model.clear))
        card.body.addWidget(LogConsole(self.log_model))
        return card

    # ── Папки й налаштування ─────────────────────────────────────────────
    def _restore_settings(self) -> None:
        self.orders_edit.setText(str(self.settings.value("orders_folder", "")))
        self.index_edit.setText(str(self.settings.value("index_folder", "")))
        self.recursive_box.setChecked(str(self.settings.value("recursive", "true")).lower() == "true")

    def _save_settings(self) -> None:
        self.settings.setValue("orders_folder", self.orders_edit.text().strip())
        self.settings.setValue("index_folder", self.index_edit.text().strip())
        self.settings.setValue("recursive", "true" if self.recursive_box.isChecked() else "false")

    def pick_orders_folder(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "Папка з наказами", self.orders_edit.text().strip())
        if folder:
            self.orders_edit.setText(os.path.normpath(folder))
            self._save_settings()

    def pick_index_folder(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "Папка для збереження індексу", self.index_edit.text().strip())
        if folder:
            self.index_edit.setText(os.path.normpath(folder))
            self._save_settings()
            self.reload_positions()

    def open_index_folder(self) -> None:
        folder = self.index_edit.text().strip()
        if folder and os.path.isdir(folder):
            os.startfile(folder)

    # ── Індексування ─────────────────────────────────────────────────────
    def start_indexing(self, force: bool) -> None:
        orders = self.orders_edit.text().strip()
        index = self.index_edit.text().strip()
        if not orders or not os.path.isdir(orders):
            self.log_model.append("❌ Помилка: оберіть наявну папку з наказами")
            return
        if not index:
            self.log_model.append("❌ Помилка: оберіть папку, куди зберегти індекс")
            return
        self._save_settings()
        self.log_model.append(f"Індексування: {orders}" + (" (повністю заново)" if force else ""))
        self.worker = IndexWorker(orders, index, self.recursive_box.isChecked(), force)
        self.worker.logged.connect(self.log_model.append)
        self.worker.progressed.connect(self._on_progress)
        self.worker.finished_with.connect(self._on_finished)
        self._set_busy(True)
        self.worker.start()

    def stop_indexing(self) -> None:
        if self.worker is not None:
            self.worker.cancel()
            self.stop_button.setEnabled(False)

    def _on_progress(self, done: int, total: int, name: str) -> None:
        percent = 100.0 * done / total if total else 100.0
        self.status.set_progress(f"{done}/{total}  {name}"[:90], percent)

    def _on_finished(self, summary: store.IndexSummary | None, error: str) -> None:
        self._set_busy(False)
        if self.worker is not None:
            self.worker.wait()
            self.worker = None
        if summary is None:
            self.log_model.append(f"❌ Помилка індексування: {error}")
            return
        self.log_model.append(
            f"✅ Завершено: переглянуто {summary.scanned}, прочитано {summary.indexed}, "
            f"без змін {summary.unchanged}, прибрано зниклих {summary.removed}, "
            f"помилок {summary.failed}; нових згадок посад {summary.hits}"
        )
        if summary.skipped_legacy:
            self.log_model.append(
                f"⚠ Увага: {summary.skipped_legacy} файлів .doc/.rtf пропущено — Word недоступний"
            )
        self.reload_positions()

    def _set_busy(self, busy: bool) -> None:
        self.status.set_busy(busy)
        for widget in (self.run_button, self.force_button, self.orders_edit, self.index_edit, self.recursive_box):
            widget.setEnabled(not busy)
        self.stop_button.setEnabled(busy)
        if not busy:
            self.status.set_progress("", 0)

    # ── Результати ───────────────────────────────────────────────────────
    def reload_positions(self) -> None:
        index = self.index_edit.text().strip()
        self.rows = store.load_positions(index) if index else []
        stats = store.load_stats(index) if index else None
        if stats is None:
            self.status.template.setText("Індекс ще не створено")
            self.strip.status.setText("")
        else:
            self.status.template.setText(f"Індекс: {store.index_path(index)}")
            extra = f" · <span style='color:{theme.C['amber_text']}'>.doc пропущено: {stats.legacy}</span>" if stats.legacy else ""
            errors = f" · <span style='color:{theme.C['rose_text']}'>помилок: {stats.errors}</span>" if stats.errors else ""
            self.strip.status.setText(
                f"Наказів у індексі: <b>{stats.ok}</b> · посад: <b>{stats.positions}</b> · "
                f"згадок: <b>{stats.mentions}</b> · пунктів: <b>{stats.items}</b> · "
                f"осіб: <b>{stats.people}</b>{extra}{errors}"
            )
        self.apply_filter()

    def _visible_rows(self) -> list[store.PositionRow]:
        needle = " ".join(self.search_edit.text().casefold().split())
        role = ROLE_FILTERS[self.role_combo.currentIndex()][1]
        known = KNOWN_FILTERS[self.known_combo.currentIndex()][1]
        visible = []
        for row in self.rows:
            if needle and needle not in " ".join(row.display.casefold().split()):
                continue
            if role == ROLE_CURRENT and not row.current or role == ROLE_TARGET and not row.target:
                continue
            if known is not None and row.known != known:
                continue
            visible.append(row)
        return visible

    def apply_filter(self, *_args) -> None:
        table = self.positions_table
        visible = self._visible_rows()
        table.setUpdatesEnabled(False)
        table.setSortingEnabled(False)
        table.clear()
        items = []
        mo317 = load_mo317()
        for row in visible:
            codes, vos, details = _mo317_columns(mo317.get(strip_brackets(row.display)))
            values = [
                (row.display, None), ("✓" if row.known else "—", row.known), (str(row.mentions), row.mentions),
                (str(row.orders), row.orders), (str(row.units or ""), row.units),
                (str(row.current or ""), row.current),
                (str(row.target or ""), row.target), (_ua_date(row.first_date), row.first_date),
                (_ua_date(row.last_date), row.last_date),
                (codes, None), (vos, None),
            ]
            item = SortableItem()
            for column, (text, sort_value) in enumerate(values):
                item.setText(column, text)
                if sort_value is not None:
                    item.setData(column, SORT_ROLE, sort_value)
            item.setToolTip(0, row.display + (f"\n\n{details}" if details else ""))
            if details:
                item.setToolTip(9, details)
                item.setToolTip(10, details)
            item.setData(0, ID_ROLE, row.id)
            items.append(item)
        table.addTopLevelItems(items)
        table.setSortingEnabled(True)
        table.setUpdatesEnabled(True)
        self.count_badge.setText(f"{len(visible)} з {len(self.rows)}" if len(visible) != len(self.rows) else str(len(self.rows)))
        self.status.selected.setText(f"Показано посад: {len(visible)}")
        self.show_occurrences()

    def show_occurrences(self) -> None:
        table = self.occurrences_table
        table.setSortingEnabled(False)
        table.clear()
        selected = self.positions_table.selectedItems()
        if not selected:
            self.selected_label.setText("Оберіть посаду зліва")
            self.occurrence_badge.setText("—")
            table.setSortingEnabled(True)
            return
        item = selected[0]
        codes, vos = item.text(9), item.text(10)
        self.selected_label.setText(item.text(0) + (f"   ·   код посади {codes}, ВОС: {vos}" if codes else ""))
        self.selected_label.setToolTip(item.toolTip(9))
        occurrences = store.load_occurrences(self.index_edit.text().strip(), int(item.data(0, ID_ROLE)))
        rows = []
        for occurrence in occurrences:
            row = SortableItem()
            row.setText(0, _ua_date(occurrence.date))
            row.setData(0, SORT_ROLE, occurrence.date)
            row.setText(1, occurrence.number)
            row.setText(2, occurrence.section.replace("§", "").strip())
            row.setText(3, occurrence.role)
            row.setText(4, occurrence.unit or "—")
            row.setToolTip(4, occurrence.full_text)
            row.setText(5, Path(occurrence.path).name)
            row.setToolTip(5, occurrence.path)
            row.setData(5, ID_ROLE, occurrence.path)
            rows.append(row)
        table.addTopLevelItems(rows)
        table.setSortingEnabled(True)
        table.sortByColumn(0, Qt.SortOrder.DescendingOrder)
        orders = len({occurrence.path for occurrence in occurrences})
        self.occurrence_badge.setText(f"{len(occurrences)} у {orders} нак.")

    def open_occurrence_file(self, item: QTreeWidgetItem, _column: int) -> None:
        path = item.data(5, ID_ROLE)
        if path and os.path.isfile(path):
            os.startfile(path)
        else:
            self.log_model.append("⚠ Увага: файл наказу не знайдено — переіндексуйте папку")

    def open_positions_extra(self) -> None:
        os.startfile(str(ALL_POSITIONS_FILE))

    def open_unit_markers(self) -> None:
        os.startfile(str(MARKERS_FILE))

    def copy_positions(self) -> None:
        rows = self._visible_rows()
        QGuiApplication.clipboard().setText("\n".join(row.display for row in rows))
        self.log_model.append(f"✓ Скопійовано посад: {len(rows)}")

    def export_csv(self) -> None:
        index = self.index_edit.text().strip()
        if not index or not os.path.isdir(index):
            self.log_model.append("❌ Помилка: спершу оберіть папку індексу")
            return
        target = Path(index) / "positions.csv"
        with open(target, "w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.writer(handle, delimiter=";")
            writer.writerow(POSITION_COLUMNS)
            for row in self._visible_rows():
                writer.writerow([row.display, "так" if row.known else "ні", row.mentions, row.orders, row.units, row.current,
                                 row.target, _ua_date(row.first_date), _ua_date(row.last_date),
                                 *_mo317_columns(load_mo317().get(strip_brackets(row.display)))[:2]])
        self.log_model.append(f"✓ Збережено: {target}")

    def closeEvent(self, event) -> None:
        self._save_settings()
        if self.worker is not None:
            self.worker.cancel()
            self.worker.wait(5000)
        super().closeEvent(event)


def launch(argv: list[str] | None = None) -> int:
    app = QApplication.instance() or QApplication(list(argv if argv is not None else sys.argv))
    app.setApplicationName("Індексатор посад")
    theme.apply_theme(app)
    window = IndexerWindow()
    window.show()
    theme.use_dark_title_bar(window)
    return app.exec()
