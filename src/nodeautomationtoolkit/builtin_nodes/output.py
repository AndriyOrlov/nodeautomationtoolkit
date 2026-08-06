from __future__ import annotations

from dataclasses import asdict, is_dataclass
from pprint import pformat
from typing import Any

from nodeautomationtoolkit.core.definition import node
from nodeautomationtoolkit.core.table_types import DataTable


def format_result(value: Any) -> str:
    if is_dataclass(value) and not isinstance(value, type):
        value = asdict(value)
    if isinstance(value, str):
        return value
    return pformat(value, width=100, sort_dicts=False)


def _show_result_dialog(title: str, text: str) -> None:
    from PySide6.QtGui import QGuiApplication
    from PySide6.QtWidgets import (
        QApplication,
        QDialog,
        QDialogButtonBox,
        QPlainTextEdit,
        QPushButton,
        QVBoxLayout,
    )

    if QApplication.instance() is None:
        raise RuntimeError("Перегляд результату доступний лише у вікні програми")

    dialog = QDialog()
    dialog.setWindowTitle(title or "Результат")
    dialog.resize(760, 520)
    layout = QVBoxLayout(dialog)
    preview = QPlainTextEdit(text)
    preview.setReadOnly(True)
    layout.addWidget(preview)

    buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
    copy_button = QPushButton("Копіювати")
    copy_button.clicked.connect(lambda: QGuiApplication.clipboard().setText(text))
    buttons.addButton(copy_button, QDialogButtonBox.ButtonRole.ActionRole)
    buttons.rejected.connect(dialog.reject)
    layout.addWidget(buttons)
    dialog.exec()


def _show_table_dialog(table: DataTable) -> None:
    from PySide6.QtWidgets import (
        QApplication,
        QDialog,
        QDialogButtonBox,
        QTableWidget,
        QTableWidgetItem,
        QVBoxLayout,
    )

    if QApplication.instance() is None:
        raise RuntimeError("Перегляд таблиці доступний лише у вікні програми")
    dialog = QDialog()
    dialog.setWindowTitle(table.title)
    dialog.resize(920, 560)
    layout = QVBoxLayout(dialog)
    widget = QTableWidget(len(table.rows), len(table.columns))
    widget.setHorizontalHeaderLabels(list(table.columns))
    for row_index, row in enumerate(table.rows):
        for column_index, value in enumerate(row):
            widget.setItem(row_index, column_index, QTableWidgetItem(str(value)))
    widget.resizeColumnsToContents()
    layout.addWidget(widget)
    buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
    buttons.rejected.connect(dialog.reject)
    layout.addWidget(buttons)
    dialog.exec()


@node(
    name="Показати результат",
    category="Результат",
    description="Фінальна тестова нода: показує отримане значення в окремому вікні.",
    type_id="builtin.output.show_result",
    execution_inputs=("exec",),
    preview_policy="never",
)
def show_result(value: Any, title: str = "Результат") -> Any:
    _show_result_dialog(title, format_result(value))
    return value


@node(
    name="Показати таблицю",
    category="Результат",
    description="Фінальна тестова нода: відкриває табличний звіт у зручному вікні.",
    type_id="builtin.output.show_table",
    execution_inputs=("exec",),
    preview_policy="never",
)
def show_table(table: DataTable) -> DataTable:
    _show_table_dialog(table)
    return table


@node(
    name="Зберегти таблицю CSV",
    category="Результат",
    description="Зберігає локальний табличний звіт у CSV, який відкривається в Excel.",
    type_id="builtin.output.save_table_csv",
    outputs={"path": "str", "rows": "int", "message": "str"},
    execution_inputs=("exec",),
    execution_outputs=("then",),
    preview_policy="never",
)
def save_table_csv(table: DataTable, path: str = "", delimiter: str = ";") -> dict:
    from pathlib import Path

    if not path.strip():
        raise ValueError("Не вказано шлях CSV")
    target = Path(path).expanduser()
    if target.suffix.casefold() != ".csv":
        target = target.with_suffix(".csv")
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text("\ufeff" + table.to_csv(delimiter), encoding="utf-8")
    return {
        "path": str(target.resolve()),
        "rows": len(table.rows),
        "message": f"Збережено рядків: {len(table.rows)}",
    }


def _show_image_dialog(title: str, image_path: str) -> None:
    from pathlib import Path
    from PySide6.QtGui import QPixmap
    from PySide6.QtWidgets import (
        QApplication,
        QDialog,
        QDialogButtonBox,
        QLabel,
        QScrollArea,
        QVBoxLayout,
    )

    if QApplication.instance() is None:
        raise RuntimeError("Перегляд зображення доступний лише у вікні програми")

    path = Path(image_path).expanduser()
    if not path.is_file():
        raise FileNotFoundError(f"Зображення не знайдено: {image_path}")

    dialog = QDialog()
    dialog.setWindowTitle(title or "Прев'ю зображення")
    dialog.resize(900, 680)
    layout = QVBoxLayout(dialog)

    scroll = QScrollArea()
    pixmap = QPixmap(str(path))
    label = QLabel()
    label.setPixmap(pixmap)
    scroll.setWidget(label)
    scroll.setWidgetResizable(True)
    layout.addWidget(scroll)

    buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
    buttons.rejected.connect(dialog.reject)
    layout.addWidget(buttons)
    dialog.exec()


@node(
    name="Показати зображення",
    category="Результат",
    description="Фінальна нода: відкриває кольорове зображення або схему в окремому вікні.",
    type_id="builtin.output.show_image",
    execution_inputs=("exec",),
    preview_policy="never",
)
def show_image(image_path: str, title: str = "Схема макету наказу") -> str:
    _show_image_dialog(title, image_path)
    return image_path

