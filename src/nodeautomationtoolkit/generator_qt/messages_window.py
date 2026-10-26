"""Окрема програма «Повідомлення»: лише два шифровані повідомлення за наказом.

Логіка та сама, що й у вкладці повідомлень генератора
(`generate_extracts.App.run_generate_messages`). Тут — мінімальне вікно:
наказ, таблиця, два зразки, папка результату й виконавець. Налаштування
лежать в окремому файлі поруч із програмою, щоб не змішуватися з конфігом
повного генератора.
"""

from __future__ import annotations

import os
import sys

from PySide6.QtCore import Qt
from PySide6.QtGui import QCloseEvent, QDragEnterEvent, QDragMoveEvent, QDropEvent
from PySide6.QtWidgets import QApplication, QFrame, QGridLayout, QHBoxLayout, QMainWindow, QVBoxLayout, QWidget

from . import compat, theme
from .main_window import _package_version, install_qt_bridge
from .widgets import ActionStrip, Badge, LogConsole, SectionCard, bind_label, button, label, line_edit

CONFIG_NAME = "messages_config.json"


def config_folder() -> str:
    """Поруч з EXE, а під час розробки — у корені проєкту."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))


class MessagesWindow(QMainWindow):
    def __init__(self, shell):
        super().__init__()
        self._shell = shell
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event: QDragMoveEvent) -> None:
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent) -> None:
        paths = [url.toLocalFile() for url in event.mimeData().urls() if url.isLocalFile()]
        if paths:
            event.acceptProposedAction()
            self._shell.accept_dropped_paths(paths)

    def closeEvent(self, event: QCloseEvent) -> None:
        self._shell.save_config()
        super().closeEvent(event)


class MessagesShellMixin:
    legacy = None  # модуль generate_extracts; задається в create_messages_app_class
    # Як у повному генераторі: нові частини дописуються в таблицю заготовками.
    ADD_NEW_UNITS_TO_TABLE = True

    def __init__(self):
        self.main_window = MessagesWindow(self)
        compat.set_dialog_parent(self.main_window)
        self.log_model = compat.LogModel()
        self._busy = False
        super().__init__(compat.RootBridge(self.main_window))
        version = _package_version()
        self.main_window.setWindowTitle(f"Повідомлення за наказом{f' v{version}' if version else ''}")
        self.main_window.resize(900, 640)
        self.main_window.setMinimumSize(720, 520)
        self._refresh_summary()

    def load_config(self):
        self.config_file = os.path.join(config_folder(), CONFIG_NAME)
        super().load_config()

    # ── Інтерфейс ────────────────────────────────────────────────────────
    def create_widgets(self):
        self.log_text = compat.TextBridge(self.log_model)

        central = QWidget()
        central.setObjectName("AppRoot")
        root = QVBoxLayout(central)
        root.setContentsMargins(12, 10, 12, 12)
        root.setSpacing(8)

        header = QHBoxLayout()
        header.addWidget(label("Повідомлення за наказом", "AppName"))
        version = _package_version()
        if version:
            header.addWidget(Badge(f"v{version}", "teal"))
        header.addStretch(1)
        root.addLayout(header)

        files = SectionCard("Файли", accent="indigo")
        grid = QGridLayout()
        grid.setHorizontalSpacing(8)
        grid.setVerticalSpacing(6)
        rows = (
            ("Наказ:", self.doc_path, self.select_message_order),
            ("Таблиця (Excel):", self.excel_path, self.select_excel),
            ("Зразок супроводу:", self.message_cover_template_path, self.select_message_cover_template),
            ("Зразок зі змістом:", self.message_content_template_path, self.select_message_content_template),
            ("Папка результату:", self.message_out_folder, self.select_message_output_folder),
        )
        for row, (caption, variable, pick) in enumerate(rows):
            grid.addWidget(label(caption, "FieldLabel"), row, 0)
            grid.addWidget(line_edit(variable), row, 1)
            grid.addWidget(button("📂 Обрати", "secondary", pick), row, 2)
        grid.addWidget(label("Виконавець:", "FieldLabel"), len(rows), 0)
        grid.addWidget(line_edit(self.message_executor), len(rows), 1, 1, 2)
        grid.setColumnStretch(1, 1)
        files.body.addLayout(grid)

        # Засвідчувач «Згідно з оригіналом» — теги {{засвідчувач_*}} у зразках.
        files.body.addWidget(label("Засвідчувач («Згідно з оригіналом»)", "PersonTitle"))
        certifier = QGridLayout()
        certifier.setHorizontalSpacing(8)
        certifier.addWidget(label("Посада:", "FieldLabel"), 0, 0)
        certifier.addWidget(line_edit(self.certifier_position), 0, 1)
        certifier.addWidget(label("Звання:", "FieldLabel"), 0, 2)
        certifier.addWidget(line_edit(self.certifier_rank), 0, 3)
        certifier.addWidget(label("ПІБ:", "FieldLabel"), 0, 4)
        certifier.addWidget(line_edit(self.certifier_name), 0, 5)
        certifier.setColumnStretch(1, 3)
        certifier.setColumnStretch(3, 1)
        certifier.setColumnStretch(5, 2)
        files.body.addLayout(certifier)
        hint = label(
            "Порожня папка результату — «Messages_Output» поруч із наказом. "
            "Файли можна перетягнути у вікно: .xlsx — таблиця, тека — папка результату, .docx — наказ.",
            "MutedLabel",
        )
        hint.setWordWrap(True)
        files.body.addWidget(hint)
        files.body.addWidget(bind_label(label("", "SummaryLabel"), self.message_source_summary))
        root.addWidget(files)

        strip = ActionStrip()
        self.btn_generate_messages = button("✉ Створити 2 повідомлення", "run", self.run_generate_messages)
        strip.add(self.btn_generate_messages)
        strip.add(button("📂 Відкрити папку результату", "secondary", self.open_message_output_folder))
        strip.add(button("ⓘ Теги", "sky", self.show_message_tags))
        strip.finish()
        root.addWidget(strip)

        log_card = SectionCard("Журнал", accent="slate", variant="log")
        log_card.add_action(button("Очистити", "rose", self.log_model.clear))
        log_card.body.addWidget(LogConsole(self.log_model))
        root.addWidget(log_card, 1)

        self.main_window.setCentralWidget(central)
        self.doc_path.subscribe(lambda *_args: self._refresh_summary())
        self.log("Оберіть наказ, таблицю й два зразки, потім натисніть «Створити 2 повідомлення».")

    def _refresh_summary(self) -> None:
        path = self.doc_path.get()
        self.message_source_summary.set(
            "Наказ: " + os.path.basename(path) if path and os.path.isfile(path) else "Наказ не обрано"
        )

    # ── Дії ──────────────────────────────────────────────────────────────
    def select_message_order(self):
        path = self.legacy.filedialog.askopenfilename(filetypes=[("Word files", "*.docx *.doc")])
        if path:
            self._set_order(path)

    def _set_order(self, path: str) -> None:
        self.doc_path.set(os.path.abspath(path))
        self.save_config()
        self.log(f"Наказ: {os.path.basename(path)}")

    def accept_dropped_paths(self, paths: list[str]) -> None:
        if self._busy:
            self.log("Зачекайте завершення поточного запуску.")
            return
        for path in paths:
            name = os.path.basename(path).lower()
            if os.path.isdir(path):
                self.message_out_folder.set(path)
                self.log(f"Папка результату: {path}")
            elif name.endswith((".xlsx", ".xls")):
                self.excel_path.set(path)
                self.log(f"Таблиця: {os.path.basename(path)}")
            elif name.endswith((".docx", ".doc")) and not name.startswith("~$"):
                self._set_order(path)
        self.save_config()

    def run_generate_messages(self):
        if self._busy:
            return None
        self._busy = True
        QApplication.setOverrideCursor(Qt.CursorShape.BusyCursor)
        try:
            return super().run_generate_messages()
        finally:
            QApplication.restoreOverrideCursor()
            self._busy = False

    def open_message_output_folder(self):
        folder = self.message_out_folder.get()
        if not folder and self.doc_path.get():
            folder = os.path.join(os.path.dirname(os.path.abspath(self.doc_path.get())), "Messages_Output")
        if folder and os.path.isdir(folder):
            os.startfile(folder)  # noqa: S606 - лише Windows
        else:
            self.legacy.messagebox.showwarning(
                "Папка результату", "Папка ще не створена. Вона з'явиться після першого запуску."
            )

    def log(self, message: str):
        self.log_model.append(message)
        compat.pump_events()

    def clear_log(self):
        self.log_model.clear()

    # Повний генератор оновлює тут свої списки й поля підписанта; у цьому
    # вікні їх немає, а підписанта повідомлення читають із наказу самі.
    def _refresh_orders_view(self):
        return None

    def _refresh_order_signer(self, text=None):
        return "", {"position": "", "rank": "", "name": ""}


def create_messages_app_class(legacy):
    return type("MessagesApp", (MessagesShellMixin, legacy.App), {"legacy": legacy})


def launch(legacy, argv: list[str] | None = None) -> int:
    qt_app = QApplication.instance() or QApplication(list(argv if argv is not None else sys.argv))
    qt_app.setApplicationName("Повідомлення за наказом")
    theme.apply_theme(qt_app)
    install_qt_bridge(legacy)
    shell = create_messages_app_class(legacy)()
    shell.main_window.show()
    theme.use_dark_title_bar(shell.main_window)
    return qt_app.exec()
