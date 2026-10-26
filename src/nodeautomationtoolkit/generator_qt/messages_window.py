"""Окрема програма: лише повідомлення про прийняття кадрових рішень.

Це та сама сторінка «Повідомлення», що й у генераторі (`QtShellMixin`),
без примірників, витягів і генератора наказів. Логіка — та сама
`generate_extracts.App`, тож результат не відрізняється від генератора.
Свій config: `config_messages.json` поруч із EXE (або з `generate_extracts.py`
під час запуску з коду), щоб не чіпати налаштувань генератора.
"""

from __future__ import annotations

import os
import sys

from PySide6.QtCore import Qt
from PySide6.QtGui import QIcon, QPixmap
from PySide6.QtWidgets import QApplication, QFrame, QHBoxLayout, QLabel, QVBoxLayout, QWidget

from . import compat, theme
from .main_window import QtShellMixin, _package_version, _short_folder, install_qt_bridge
from .samples_dialog import SamplesDialog
from .widgets import Badge, StatusStrip, button, label

APP_TITLE = "Повідомлення про прийняття кадрових рішень"
CONFIG_NAME = "config_messages.json"
# Лого (щит із булавою, мечем і хрестом): SVG — джерело, PNG — для вікна,
# ICO — для самого EXE (`build_messages_qt.bat`).
LOGO_PNG = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icons", "messages_logo.png")
# `App.handle_drag_and_drop` питає номер активної вкладки; у генераторі
# повідомлення — третя (індекс 2), і саме за ним обробник їх розпізнає.
MESSAGES_TAB = 2


class _MessagesOnlyNotebook:
    """Замінник вкладок для `handle_drag_and_drop`: вкладка одна — повідомлення."""

    def select(self):
        return MESSAGES_TAB

    def index(self, _tab):
        return MESSAGES_TAB


def config_folder(legacy_file: str) -> str:
    """Тека для config_messages.json. Однофайловий EXE розпаковується в тимчасову
    теку, яка зникає після закриття, тож конфіг лежить поруч із самим EXE."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(legacy_file))


class MessagesShellMixin(QtShellMixin):
    """Вікно з однією сторінкою «Повідомлення»."""

    def __init__(self):
        self._strip_status = {}
        super().__init__()
        self.main_window.setWindowTitle(f"{APP_TITLE} — {self.legacy.APP_VERSION}")
        self.main_window.setWindowIcon(QIcon(LOGO_PNG))
        self.main_window.resize(1100, 820)
        self.main_window.setMinimumSize(860, 620)

    def _config_path(self) -> str:
        return os.path.join(config_folder(self.legacy.__file__), CONFIG_NAME)

    def create_widgets(self):
        self.log_text = compat.TextBridge(self.log_model)
        self.p2_log_text = compat.TextBridge(self.p2_log_model)
        self.notebook = _MessagesOnlyNotebook()

        central = QWidget()
        central.setObjectName("AppRoot")
        root = QVBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)
        root.addWidget(self._build_header())
        self.tab_messages_page = self._scroll_page(self._build_messages_tab())
        self.tab_messages = self.tab_messages_page
        root.addWidget(self.tab_messages_page, 1)
        self.status_strip = StatusStrip()
        root.addWidget(self.status_strip)
        self.main_window.setCentralWidget(central)

        refresh = lambda *_args: self._refresh_badges()  # noqa: E731
        for variable in (
            self.doc_path,
            self.message_source_summary,
            self.message_content_template_path,
        ):
            variable.subscribe(refresh)
        self.progress_text.subscribe(lambda value: self.status_strip.set_progress(caption=value))
        self.progress_value.subscribe(lambda value: self.status_strip.set_progress(percent=value))

        self._refresh_source_summary()
        self.log("Готовий до роботи. Оберіть наказ і натисніть «Створити 2 повідомлення».")

    def _build_header(self) -> QFrame:
        bar = QFrame()
        bar.setObjectName("TitleBar")
        layout = QHBoxLayout(bar)
        layout.setContentsMargins(12, 6, 12, 6)
        layout.setSpacing(8)
        logo = QLabel()
        logo.setPixmap(
            QPixmap(LOGO_PNG).scaled(
                24,
                24,
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation,
            )
        )
        layout.addWidget(logo)
        layout.addWidget(label(APP_TITLE, "AppName"))
        version = _package_version()
        if version:
            layout.addWidget(Badge(f"v{version}", "teal"))
        layout.addStretch(1)
        layout.addWidget(
            self._lock(button("⚙  Зразки та словник", "amber", self.open_samples_window))
        )
        return bar

    def _refresh_orders_view(self):
        # Списку наказів тут немає — лише підсумковий рядок про обраний наказ.
        self._refresh_source_summary()

    def _refresh_badges(self, *_args) -> None:
        if not getattr(self, "_ui_ready", False):
            return
        order = self.doc_path.get()
        count = 1 if order and os.path.isfile(order) else 0
        self._strip_status[MESSAGES_TAB].setText(self._strip_text(count, self._busy_depth > 0))
        self.status_strip.selected.setText(f"Наказ: {os.path.basename(order) if count else '—'}")
        template = self.message_content_template_path.get()
        c = theme.C
        self.status_strip.template.setText(
            f'Зразок: <b style="color:{c["text_soft"]}; font-family:{theme.MONO}">'
            f"{os.path.basename(template) if template else 'не обрано'}</b>"
            f'  <span style="color:{c["text_dim"]}">•</span>  '
            f'Тека: <span style="color:{c["text_soft"]}">'
            f"{_short_folder(os.path.dirname(order)) if count else '—'}</span>"
        )

    def open_samples_window(self):
        if self._samples_dialog is None:
            dialog = SamplesDialog(self, self.main_window, messages_only=True)
            dialog.finished.connect(lambda *_args: setattr(self, "_samples_dialog", None))
            self._samples_dialog = dialog
        self._samples_dialog.show()
        self._samples_dialog.raise_()
        self._samples_dialog.activateWindow()


def create_messages_app_class(legacy):
    return type("QtMessagesApp", (MessagesShellMixin, legacy.App), {"legacy": legacy})


def launch_messages(legacy, argv: list[str] | None = None) -> int:
    qt_app = QApplication.instance() or QApplication(list(argv if argv is not None else sys.argv))
    qt_app.setApplicationName(APP_TITLE)
    qt_app.setWindowIcon(QIcon(LOGO_PNG))
    theme.apply_theme(qt_app)
    install_qt_bridge(legacy)
    shell = create_messages_app_class(legacy)()
    shell.main_window.show()
    theme.use_dark_title_bar(shell.main_window)
    return qt_app.exec()
