from __future__ import annotations

from dataclasses import asdict, is_dataclass
from pprint import pformat
from typing import Any

from nodeautomationtoolkit.core.definition import node


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


@node(
    name="Показати результат",
    category="Результат",
    description="Фінальна тестова нода: показує отримане значення в окремому вікні.",
    type_id="builtin.output.show_result",
    execution_inputs=("exec",),
)
def show_result(value: Any, title: str = "Результат") -> Any:
    _show_result_dialog(title, format_result(value))
    return value
