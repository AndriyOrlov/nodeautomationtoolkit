"""Пауза між кроками друку: Пробіл або Enter — далі, Esc — зупинити.

Після кожного кроку надруковане треба відкласти (а перед задньою сторінкою —
перекласти верхній аркуш у ручну подачу), тож програма чекає натискання.
"""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QDialog, QHBoxLayout, QVBoxLayout, QWidget

from .fluent import BodyLabel, CaptionLabel, PrimaryPushButton, PushButton, StrongBodyLabel, SubtitleLabel


class PrintPauseDialog(QDialog):
    def __init__(self, done: str, after: str, upcoming: str, parent: QWidget | None = None):
        super().__init__(parent)
        self.setObjectName("PrintPauseDialog")
        self.setWindowTitle("Друк — пауза")
        self.setMinimumWidth(520)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(12)

        layout.addWidget(SubtitleLabel("Надруковано"))
        done_label = BodyLabel(done)
        done_label.setWordWrap(True)
        layout.addWidget(done_label)
        if after:
            action = StrongBodyLabel(after)
            action.setWordWrap(True)
            layout.addWidget(action)
        layout.addWidget(BodyLabel("Відкладіть надруковане."))
        upcoming_label = CaptionLabel(f"Далі: {upcoming}" if upcoming else "Це був останній крок.")
        upcoming_label.setWordWrap(True)
        layout.addWidget(upcoming_label)

        buttons = QHBoxLayout()
        buttons.addStretch(1)
        self.stop_button = PushButton("Зупинити (Esc)")
        self.continue_button = PrimaryPushButton("Продовжити (Пробіл / Enter)" if upcoming else "Готово (Enter)")
        self.stop_button.clicked.connect(self.reject)
        self.continue_button.clicked.connect(self.accept)
        buttons.addWidget(self.stop_button)
        buttons.addWidget(self.continue_button)
        layout.addLayout(buttons)
        self.continue_button.setFocus()

    def keyPressEvent(self, event) -> None:  # noqa: N802 - назва з Qt
        if event.key() in (Qt.Key.Key_Space, Qt.Key.Key_Return, Qt.Key.Key_Enter):
            self.accept()
            return
        if event.key() == Qt.Key.Key_Escape:
            self.reject()
            return
        super().keyPressEvent(event)
