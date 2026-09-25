"""Вікно «Інструкція»: показує `instruction.md` поруч із цим модулем.

Текст лежить окремим файлом, щоб його можна було правити без зміни коду. У зібраному
exe файл додається в `build_generator_qt.bat` (`--add-data`).
"""

from __future__ import annotations

from pathlib import Path

from PySide6.QtWidgets import QDialog, QHBoxLayout, QTextBrowser, QVBoxLayout, QWidget

from .widgets import button

INSTRUCTION_FILE = Path(__file__).with_name("instruction.md")


def load_instruction_text() -> str:
    try:
        return INSTRUCTION_FILE.read_text(encoding="utf-8")
    except OSError:
        return (
            "# Інструкція\n\nФайл інструкції не знайдено поруч із програмою "
            f"(`{INSTRUCTION_FILE.name}`)."
        )


class InstructionDialog(QDialog):
    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setObjectName("InstructionDialog")
        self.setWindowTitle("Інструкція")
        self.resize(900, 760)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(10)

        self.viewer = QTextBrowser()
        self.viewer.setObjectName("InstructionText")
        self.viewer.setOpenExternalLinks(False)
        self.viewer.setMarkdown(load_instruction_text())
        layout.addWidget(self.viewer, 1)

        bottom = QHBoxLayout()
        bottom.addStretch(1)
        bottom.addWidget(button("Закрити", "secondary", self.close))
        layout.addLayout(bottom)
