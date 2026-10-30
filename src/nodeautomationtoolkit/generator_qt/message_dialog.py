"""Повідомлення програми окремим вікном (замість затемнення Fluent усередині вікна).

Діалог Fluent (`MessageBox`) — це не окреме вікно, а напівпрозора накладка
ВСЕРЕДИНІ головного вікна. Коли головне вікно перекрите Word'ом чи згорнуте, або
текст довгий (перелік файлів) і кнопки виходять за межі вікна, користувач до
нього не діставався, а програма чекала відповіді — виглядало як зависання
(29.10.2026). Тут — звичайне вікно верхнього рівня поверх усіх: текст
прокручується, кнопки видно завжди, Enter — так/гаразд, Esc — ні/скасувати.
"""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QDialog, QHBoxLayout, QScrollArea, QVBoxLayout, QWidget

from . import theme
from .fluent import BodyLabel, PrimaryPushButton, PushButton, SubtitleLabel

#: Скільки місця дає текст, перш ніж з'являється прокрутка.
_MAX_TEXT_HEIGHT = 420


class MessageDialog(QDialog):
    def __init__(
        self,
        title: str,
        message: str,
        yes_text: str = "Гаразд",
        no_text: str | None = None,
        parent: QWidget | None = None,
    ):
        super().__init__(parent)
        self.setObjectName("MessageDialog")
        self.setWindowTitle(str(title) or "Повідомлення")
        self.setWindowFlags(
            Qt.WindowType.Dialog
            | Qt.WindowType.WindowTitleHint
            | Qt.WindowType.WindowCloseButtonHint
            | Qt.WindowType.WindowStaysOnTopHint
        )
        self.setMinimumWidth(460)
        self.setMaximumWidth(760)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 20, 24, 18)
        layout.setSpacing(12)
        heading = SubtitleLabel(str(title))
        heading.setWordWrap(True)
        layout.addWidget(heading)

        text = BodyLabel(str(message))
        text.setWordWrap(True)
        text.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        text.setObjectName("MessageText")
        holder = QWidget()
        holder.setObjectName("TabPage")
        holder_layout = QVBoxLayout(holder)
        holder_layout.setContentsMargins(0, 0, 8, 0)
        holder_layout.addWidget(text)
        holder_layout.addStretch(1)
        scroll = QScrollArea()
        scroll.setObjectName("TabScroll")
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QScrollArea.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll.setWidget(holder)
        holder.adjustSize()
        scroll.setMinimumHeight(min(holder.sizeHint().height() + 8, _MAX_TEXT_HEIGHT))
        scroll.setMaximumHeight(_MAX_TEXT_HEIGHT)
        layout.addWidget(scroll, 1)

        buttons = QHBoxLayout()
        buttons.addStretch(1)
        self.no_button = None
        if no_text is not None:
            self.no_button = PushButton(no_text)
            self.no_button.clicked.connect(self.reject)
            buttons.addWidget(self.no_button)
        self.yes_button = PrimaryPushButton(yes_text)
        self.yes_button.clicked.connect(self.accept)
        self.yes_button.setDefault(True)
        buttons.addWidget(self.yes_button)
        layout.addLayout(buttons)
        self.yes_button.setFocus()

    def keyPressEvent(self, event) -> None:  # noqa: N802 - назва з Qt
        if event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            self.accept()
            return
        if event.key() == Qt.Key.Key_Escape:
            self.reject()
            return
        super().keyPressEvent(event)

    def showEvent(self, event) -> None:  # noqa: N802
        super().showEvent(event)
        theme.use_dark_title_bar(self)
        self.raise_()
        self.activateWindow()


def ask(title: str, message: str, yes_text: str = "Гаразд", no_text: str | None = None, parent=None) -> bool:
    """Показати діалог і дочекатися відповіді: True — так/гаразд, False — ні/скасувати/закрито."""
    dialog = MessageDialog(title, message, yes_text, no_text, parent)
    try:
        return dialog.exec() == QDialog.DialogCode.Accepted
    finally:
        dialog.deleteLater()
