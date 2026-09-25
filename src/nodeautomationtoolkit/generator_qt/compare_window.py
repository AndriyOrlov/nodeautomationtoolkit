"""Режим порівняння DOCX (еталон ↔ згенерований) у Qt-оболонці.

Та сама логіка, що й у Tk-вікні (`compare_documents.compare_docx_documents`),
лише інтерфейс інший: змішувати цикли подій Tk і Qt в одному процесі не можна.
"""

from __future__ import annotations

import os

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QColor, QFont, QGuiApplication, QTextBlockFormat, QTextCharFormat, QTextCursor
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QFileDialog,
    QGridLayout,
    QHBoxLayout,
    QLineEdit,
    QMessageBox,
    QSplitter,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from nodeautomationtoolkit.builtin_nodes.compare_documents import CompareResult, compare_docx_documents

from . import theme
from .widgets import Badge, SectionCard, button, label

MODES = ("витяги", "примірник_2", "повідомлення_зміст", "повідомлення_супровід")


class CompareWindow(QWidget):
    def __init__(
        self,
        reference_path: str = "",
        generated_path: str = "",
        mode: str = "витяги",
        parent: QWidget | None = None,
    ):
        super().__init__(parent, Qt.WindowType.Window)
        self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)
        self.setObjectName("AppRoot")
        self.setWindowTitle("Порівняння з еталоном — Side-by-Side Diff")
        self.resize(1280, 800)
        self.last_result: CompareResult | None = None
        self._syncing_scroll = False

        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(8)

        files = SectionCard("Файли для порівняння", accent="amber")
        grid = QGridLayout()
        grid.setHorizontalSpacing(8)
        self.reference_edit = QLineEdit(reference_path or "")
        self.generated_edit = QLineEdit(generated_path or "")
        grid.addWidget(label("Еталон (ручний DOCX):", "FieldLabel"), 0, 0)
        grid.addWidget(self.reference_edit, 0, 1)
        grid.addWidget(button("📂 Обрати", "secondary", lambda: self._pick(self.reference_edit, "еталонний")), 0, 2)
        grid.addWidget(label("Згенерований DOCX:", "FieldLabel"), 1, 0)
        grid.addWidget(self.generated_edit, 1, 1)
        grid.addWidget(button("📂 Обрати", "secondary", lambda: self._pick(self.generated_edit, "згенерований")), 1, 2)
        grid.setColumnStretch(1, 1)
        files.body.addLayout(grid)

        controls = QHBoxLayout()
        controls.addWidget(label("Режим:", "FieldLabel"))
        self.mode_combo = QComboBox()
        self.mode_combo.addItems(MODES)
        if mode in MODES:
            self.mode_combo.setCurrentText(mode)
        controls.addWidget(self.mode_combo)
        self.ignore_blank_box = QCheckBox("Ігнорувати порожні абзаци")
        self.ignore_blank_box.setChecked(True)
        self.ignore_blank_box.setToolTip(
            "Зайвий чи відсутній Enter не вважається розбіжністю; "
            "правила порожніх рядків перед пунктом і підписантом не перевіряються."
        )
        controls.addWidget(self.ignore_blank_box)
        controls.addWidget(button("⇄ Порівняти файли", "run", self.run_comparison))
        controls.addStretch(1)
        controls.addWidget(label("Легенда:", "MutedLabel"))
        controls.addWidget(Badge("Пропущено / видалено", "rose"))
        controls.addWidget(Badge("Зайве / додано", "emerald"))
        controls.addWidget(Badge("Змінено / стиль", "amber"))
        files.body.addLayout(controls)
        layout.addWidget(files)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        self.reference_view = self._pane(splitter, "Еталонний документ (очікувано)")
        self.generated_view = self._pane(splitter, "Згенерований документ (отримано)")
        splitter.setSizes([640, 640])
        layout.addWidget(splitter, 1)
        self._link_scrollbars()

        bottom = QHBoxLayout()
        self.summary = label("— Оберіть файли та натисніть «Порівняти» —", "SummaryLabel")
        self.summary.setWordWrap(True)
        bottom.addWidget(self.summary, 1)
        self.save_button = button("💾 Зберегти звіт (.md)", "secondary", self.save_report_file)
        self.copy_button = button("📋 Скопіювати опис для чату AI", "amber", self.copy_chat_report)
        self.save_button.setEnabled(False)
        self.copy_button.setEnabled(False)
        bottom.addWidget(self.save_button)
        bottom.addWidget(self.copy_button)
        layout.addLayout(bottom)

        if self.reference_edit.text() and self.generated_edit.text():
            QTimer.singleShot(0, self.run_comparison)

    # ── побудова ─────────────────────────────────────────────────────────
    @staticmethod
    def _pane(splitter: QSplitter, title: str) -> QTextEdit:
        card = SectionCard(title, accent="slate", variant="log")
        view = QTextEdit()
        view.setObjectName("DiffPane")
        view.setReadOnly(True)
        view.setLineWrapMode(QTextEdit.LineWrapMode.NoWrap)
        font = QFont(theme.MONO)
        font.setStyleHint(QFont.StyleHint.Monospace)
        font.setPointSizeF(9.5)
        view.setFont(font)
        card.body.addWidget(view, 1)
        splitter.addWidget(card)
        return view

    def _link_scrollbars(self) -> None:
        left = self.reference_view.verticalScrollBar()
        right = self.generated_view.verticalScrollBar()

        def follow(source, target):
            def sync(value: int) -> None:
                if self._syncing_scroll:
                    return
                self._syncing_scroll = True
                try:
                    if source.maximum() > 0:
                        target.setValue(round(value * target.maximum() / source.maximum()))
                finally:
                    self._syncing_scroll = False

            return sync

        left.valueChanged.connect(follow(left, right))
        right.valueChanged.connect(follow(right, left))

    def _pick(self, edit: QLineEdit, which: str) -> None:
        path, _selected = QFileDialog.getOpenFileName(
            self, f"Оберіть {which} файл DOCX", "", "Word Documents (*.docx);;All Files (*.*)"
        )
        if path:
            edit.setText(path)

    # ── дії ──────────────────────────────────────────────────────────────
    def run_comparison(self) -> None:
        reference = self.reference_edit.text().strip()
        generated = self.generated_edit.text().strip()
        if not (os.path.isfile(reference) and os.path.isfile(generated)):
            QMessageBox.warning(self, "Помилка вибору", "Вкажіть дійсні шляхи до обох DOCX файлів.")
            return
        try:
            result = compare_docx_documents(
                reference,
                generated,
                mode=self.mode_combo.currentText(),
                ignore_blank_paragraphs=self.ignore_blank_box.isChecked(),
            )
        except Exception as error:
            QMessageBox.critical(self, "Помилка порівняння", f"Не вдалося порівняти файли:\n{error}")
            return
        self.last_result = result
        self._fill(self.reference_view, result.side_by_side_rows, "ref_line")
        self._fill(self.generated_view, result.side_by_side_rows, "gen_line")
        self.summary.setText(result.summary_text)
        self.save_button.setEnabled(True)
        self.copy_button.setEnabled(True)

    @staticmethod
    def _fill(view: QTextEdit, rows: list[dict], side: str) -> None:
        view.clear()
        cursor = view.textCursor()
        cursor.beginEditBlock()
        for position, row in enumerate(rows):
            background, foreground, bold = theme.DIFF_TONES.get(row.get("status"), theme.DIFF_TONES["EQUAL"])
            block = QTextBlockFormat()
            if background:
                block.setBackground(QColor(background))
            characters = QTextCharFormat()
            characters.setForeground(QColor(foreground))
            if bold:
                characters.setFontWeight(QFont.Weight.Bold)
            if position == 0:
                cursor.setBlockFormat(block)
            else:
                cursor.insertBlock(block)
            cursor.insertText(row.get(side) or "[—]", characters)
        cursor.endEditBlock()
        view.moveCursor(QTextCursor.MoveOperation.Start)

    def copy_chat_report(self) -> None:
        if not self.last_result:
            return
        QGuiApplication.clipboard().setText(self.last_result.ai_chat_report)
        QMessageBox.information(
            self,
            "Скопійовано",
            "Структурований опис розбіжностей скопійовано в буфер обміну.\n"
            "Вставте його (Ctrl+V) у чат AI.",
        )

    def save_report_file(self) -> None:
        if not self.last_result:
            return
        path, _selected = QFileDialog.getSaveFileName(
            self, "Зберегти звіт порівняння", "", "Markdown (*.md);;Text Files (*.txt)"
        )
        if not path:
            return
        if not os.path.splitext(path)[1]:
            path += ".md"
        with open(path, "w", encoding="utf-8") as handle:
            handle.write(self.last_result.ai_chat_report)
        QMessageBox.information(self, "Збережено", f"Звіт збережено:\n{path}")
