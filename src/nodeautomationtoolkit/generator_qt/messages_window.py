"""Окрема програма «Повідомлення»: лише два шифровані повідомлення за наказом.

Логіка та сама, що й у вкладці повідомлень генератора
(`generate_extracts.App.run_generate_messages`). Тут — мінімальне вікно:
у головному вікні лише наказ; таблиця, зразки, папка результату, виконавець
і засвідчувач — у вікні «Налаштування». Налаштування
лежать в окремому файлі поруч із програмою, щоб не змішуватися з конфігом
повного генератора.
"""

from __future__ import annotations

import os
import sys

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication, QDialog, QFrame, QGridLayout, QHBoxLayout, QScrollArea, QVBoxLayout, QWidget

from nodeautomationtoolkit.generator.errors import explain_error
from nodeautomationtoolkit.generator.order_text import text_before_order_signer

from . import compat, theme
from .fluent import FluentIcon, NavigationItemPosition, TitleLabel
from .main_window import DropFluentWindow, _package_version, install_qt_bridge
from .widgets import (
    ActionStrip,
    LogConsole,
    SectionCard,
    bind_label,
    button,
    check_box,
    label,
    line_edit,
    run_params_card,
)

CONFIG_NAME = "messages_config.json"


def config_folder() -> str:
    """Поруч з EXE, а під час розробки — у корені проєкту."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))


class MessagesWindow(DropFluentWindow):
    """Вікно Fluent з однією сторінкою; перетягування й закриття — як у генераторі."""

    def __init__(self, shell):
        super().__init__(shell)
        self.navigationInterface.setExpandWidth(210)


class MessagesShellMixin:
    legacy = None  # модуль generate_extracts; задається в create_messages_app_class
    # Як у повному генераторі: нові частини дописуються в таблицю заготовками.
    ADD_NEW_UNITS_TO_TABLE = True

    def on_window_close(self) -> None:
        self.save_config()

    def __init__(self):
        self.main_window = MessagesWindow(self)
        compat.set_dialog_parent(self.main_window)
        self.log_model = compat.LogModel()
        self._busy = False
        super().__init__(compat.RootBridge(self.main_window))
        version = _package_version()
        self.main_window.setWindowTitle(f"Повідомлення за наказом{f' v{version}' if version else ''}")
        self.main_window.resize(1040, 820)
        self.main_window.setMinimumSize(720, 520)
        self._refresh_summary()

    def load_config(self):
        self.config_file = os.path.join(config_folder(), CONFIG_NAME)
        super().load_config()

    # ── Інтерфейс ────────────────────────────────────────────────────────
    def create_widgets(self):
        self.log_text = compat.TextBridge(self.log_model)

        central = QWidget()
        central.setObjectName("TabPage")
        root = QVBoxLayout(central)
        root.setContentsMargins(28, 8, 28, 16)
        root.setSpacing(12)

        root.addWidget(TitleLabel("Повідомлення за наказом"))

        order = SectionCard("Наказ", accent="indigo")
        order.add_action(button("⚙ Налаштування", "secondary", self.open_settings))
        row = QHBoxLayout()
        row.setSpacing(8)
        row.addWidget(line_edit(self.doc_path, "Оберіть наказ або перетягніть його у вікно"), 1)
        row.addWidget(button("📂 Обрати", "secondary", self.select_message_order))
        order.body.addLayout(row)
        order.body.addWidget(bind_label(label("", "SummaryLabel"), self.message_source_summary))
        root.addWidget(order)

        # Ті самі поля, що на сторінці повідомлень великої програми (вимога 26.10.2026):
        # виконавець, папка, підписант оригіналу наказу, засвідчувач, галочка управління.
        def message_options(row: QHBoxLayout) -> None:
            skip_box = check_box(
                "Не додавати внутрішнє переміщення в управлінні",
                self.message_skip_internal_management,
            )
            skip_box.setToolTip(
                "Пункт, де «управління» є і в посаді «звідки», і в посаді «КУДИ», "
                "не переноситься у шифрований зміст."
            )
            row.addWidget(skip_box)
            chain_box = check_box("Подовжена версія", self.message_full_chain)
            chain_box.setToolTip(
                "Увімкнено — ланцюг за стовпцем D іде вище корпусу до ОК, якщо ОК "
                "названо в наказі. Вимкнено — як раніше, до корпусу. КСВ не шифрується."
            )
            row.addWidget(chain_box)
            general_box = check_box("Звичайні витяги", self.message_general_extracts)
            general_box.setToolTip(
                "Увімкнено — після повідомлень у ту саму папку створюються адресні витяги, "
                "як кнопка «Створити витяги» у великій програмі (потрібен зразок витягу)."
            )
            row.addWidget(general_box)
            management_box = check_box("Витяги до управління", self.message_management_extracts)
            management_box.setToolTip(
                "Увімкнено — після повідомлень у ту саму папку створюються витяги до "
                "управління (потрібен зразок витягу в «Налаштуваннях»)."
            )
            row.addWidget(management_box)
            row.addStretch(1)

        root.addWidget(
            run_params_card(
                "Реквізити цього прогону",
                executor_var=self.message_executor,
                folder_var=self.message_out_folder,
                on_pick_folder=self.select_message_output_folder,
                executor_hint="",
                people=(
                    (
                        "Підписант оригіналу наказу (зчитується з наказу, можна виправити)",
                        self.order_signer_position,
                        self.order_signer_rank,
                        self.order_signer_name,
                    ),
                    (
                        "Засвідчувач («Згідно з оригіналом»)",
                        self.certifier_position,
                        self.certifier_rank,
                        self.certifier_name,
                    ),
                ),
                options=message_options,
            )
        )

        strip = ActionStrip()
        self.btn_generate_messages = button("✉ Створити повідомлення", "run", self.run_generate_messages)
        strip.add(self.btn_generate_messages)
        strip.add(button("📂 Відкрити папку результату", "secondary", self.open_message_output_folder))
        strip.finish()
        root.addWidget(strip)

        log_card = SectionCard("Журнал", accent="slate", variant="log")
        log_card.add_action(button("🗑 Очистити", "ghost", self.log_model.clear))
        log_card.body.addWidget(LogConsole(self.log_model))
        root.addWidget(log_card, 1)

        page = QScrollArea()
        page.setObjectName("TabScroll")
        page.setWidgetResizable(True)
        page.setFrameShape(QFrame.Shape.NoFrame)
        page.setWidget(central)
        page.setObjectName("messagesPage")
        self.main_window.addSubInterface(page, FluentIcon.MAIL, "Повідомлення")
        nav = self.main_window.navigationInterface
        bottom = NavigationItemPosition.BOTTOM
        self._nav_items = {
            "settings": nav.addItem(
                "settings", FluentIcon.SETTING, "Налаштування",
                onClick=self.open_settings, selectable=False, position=bottom,
            ),
            "theme": nav.addItem(
                "theme", FluentIcon.CONSTRACT, "Світла / темна тема",
                onClick=self.toggle_theme, selectable=False, position=bottom,
            ),
        }
        self.doc_path.subscribe(lambda *_args: self._refresh_summary())
        self.log("Оберіть наказ і натисніть «Створити повідомлення». Таблиця й зразки — у «Налаштування».")

    def toggle_theme(self) -> None:
        mode = theme.LIGHT if theme.is_dark() else theme.DARK
        self.ui_theme.set(mode)
        theme.apply_theme(QApplication.instance(), mode)
        self.save_config()

    def _refresh_summary(self) -> None:
        path = self.doc_path.get()
        self.message_source_summary.set(
            "Наказ: " + os.path.basename(path) if path and os.path.isfile(path) else "Наказ не обрано"
        )

    # ── Дії ──────────────────────────────────────────────────────────────
    # Діалог вибору відкривається там, де лежить уже вибраний файл (або папка),
    # а вибране одразу пишеться в налаштування.
    @staticmethod
    def _start_folder(*values: str) -> str:
        for value in values:
            if not value:
                continue
            folder = value if os.path.isdir(value) else os.path.dirname(value)
            if os.path.isdir(folder):
                return folder
        return ""

    def _pick_file(self, variable, filetypes, *fallbacks: str) -> str:
        path = self.legacy.filedialog.askopenfilename(
            filetypes=filetypes, initialdir=self._start_folder(variable.get(), *fallbacks)
        )
        if path:
            variable.set(os.path.abspath(path))
            self.save_config()
        return path

    def select_excel(self):
        self._pick_file(self.excel_path, [("Excel files", "*.xlsx *.xls")])

    def select_message_cover_template(self):
        self._pick_file(
            self.message_cover_template_path, [("Word files", "*.docx *.doc")],
            self.message_content_template_path.get(),
        )

    def select_message_content_template(self):
        self._pick_file(
            self.message_content_template_path, [("Word files", "*.docx *.doc")],
            self.message_cover_template_path.get(),
        )

    def select_message_calc_template(self):
        self._pick_file(
            self.message_calc_template_path, [("Word files", "*.docx *.doc")],
            self.message_cover_template_path.get(),
        )

    def select_template(self):
        self._pick_file(
            self.template_path, [("Word files", "*.docx *.doc")],
            self.message_cover_template_path.get(),
        )

    def select_message_output_folder(self):
        path = self.legacy.filedialog.askdirectory(
            initialdir=self._start_folder(self.message_out_folder.get(), self.doc_path.get())
        )
        if path:
            self.message_out_folder.set(os.path.abspath(path))
            self.save_config()

    def select_message_order(self):
        path = self.legacy.filedialog.askopenfilename(
            filetypes=[("Word files", "*.docx *.doc")], initialdir=self._start_folder(self.doc_path.get())
        )
        if path:
            self._set_order(path)

    def _set_order(self, path: str) -> None:
        self.doc_path.set(os.path.abspath(path))
        self.save_config()
        self.log(f"Наказ: {os.path.basename(path)}")
        self._read_order_signer()

    def _read_order_signer(self) -> None:
        """Підписант із наказу в поля — як у великій програмі; поля можна виправити.

        Наступний наказ читається заново, тож підписант попереднього не «прилипає».
        """
        for variable in (self.order_signer_position, self.order_signer_rank, self.order_signer_name):
            variable.set("")
        QApplication.setOverrideCursor(Qt.CursorShape.BusyCursor)
        try:
            super()._refresh_order_signer()
        except Exception as error:
            self.log(f"УВАГА: підписанта не зчитано ({error}) — впишіть його вручну або залиште порожнім.")
        finally:
            QApplication.restoreOverrideCursor()

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

    def open_settings(self):
        dialog = SettingsDialog(self, self.main_window)
        theme.use_dark_title_bar(dialog)
        dialog.exec()

    def _settings_missing(self) -> list[str]:
        return [
            caption
            for caption, variable in (
                ("таблиця", self.excel_path),
                ("зразок супроводу", self.message_cover_template_path),
                ("зразок зі змістом", self.message_content_template_path),
            )
            if not variable.get()
        ]

    def run_generate_messages(self):
        if self._busy:
            return None
        missing = self._settings_missing()
        if missing:
            self.log(f"Не вибрано в налаштуваннях: {', '.join(missing)}.")
            self.open_settings()
            if self._settings_missing():
                return None
        self._busy = True
        QApplication.setOverrideCursor(Qt.CursorShape.BusyCursor)
        try:
            result = super().run_generate_messages()
            if getattr(self, "last_messages_created", False):
                self._run_extracts_after_messages()
            return result
        finally:
            QApplication.restoreOverrideCursor()
            self._busy = False

    def _extracts_scope(self) -> str:
        """Які витяги робити після повідомлень: обидві галочки → усі, одна — лише її."""
        general = self.message_general_extracts.get()
        management = self.message_management_extracts.get()
        if general and management:
            return "all"
        if general:
            return "general"
        if management:
            return "management"
        return ""

    def _run_extracts_after_messages(self) -> None:
        """Витяги — у ту саму папку, що й повідомлення (05.11.2026).

        Логіка та сама, що й кнопки «Створити витяги» / «Витяги до управління»
        великої програми; лише папка — папка повідомлень, а виконавець —
        виконавець повідомлень.
        """
        scope = self._extracts_scope()
        if not scope:
            return
        template = self.template_path.get()
        if not template or not os.path.isfile(template):
            self.log("Витяги не створено: у «Налаштуваннях» не вибрано зразок витягу.")
            return
        saved = (self.out_folder.get(), self.executor.get())
        self.out_folder.set(self.message_out_folder.get())
        if not self.executor.get().strip():
            self.executor.set(self.message_executor.get())
        try:
            self.run_extracts(scope=scope, notify_empty=False)
        except Exception as error:
            explanation = explain_error(error)
            self.log(f"ПОМИЛКА витягів:\n{explanation}")
            self.legacy.messagebox.showerror("Витяги не створені", explanation)
        finally:
            self.out_folder.set(saved[0])
            self.executor.set(saved[1])
            self.legacy.clear_table_caches()

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
        # Витяги до управління передають текст наказу: обрізаємо його по
        # підписанту, але поля підписанта (їх уже зчитано й, можливо,
        # виправлено) не чіпаємо.
        if text is None:
            return "", {"position": "", "rank": "", "name": ""}
        return text_before_order_signer(text)

    # Таблиць результатів витягів у цьому вікні немає — підсумок у журналі.
    def _populate_result_tab(self, key, columns, rows):
        return None

    def show_layout_warnings(self, warnings):
        return None


class SettingsDialog(QDialog):
    """Усе, що не змінюється від наказу до наказу: таблиця, зразки, папка, люди."""

    def __init__(self, shell, parent: QWidget | None = None):
        super().__init__(parent)
        self._shell = shell
        self.setObjectName("SamplesDialog")
        self.setWindowTitle("Налаштування")
        self.resize(860, 380)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(8)

        grid = QGridLayout()
        grid.setHorizontalSpacing(8)
        grid.setVerticalSpacing(6)
        rows = (
            ("Таблиця (Excel):", shell.excel_path, shell.select_excel),
            ("Зразок супроводу:", shell.message_cover_template_path, shell.select_message_cover_template),
            ("Зразок зі змістом:", shell.message_content_template_path, shell.select_message_content_template),
            (
                "Розрахунок розсилки:",
                shell.message_calc_template_path,
                shell.select_message_calc_template,
            ),
            ("Зразок витягу:", shell.template_path, shell.select_template),
        )
        for row, (caption, variable, pick) in enumerate(rows):
            grid.addWidget(label(caption, "FieldLabel"), row, 0)
            grid.addWidget(line_edit(variable), row, 1)
            grid.addWidget(button("📂 Обрати", "secondary", pick), row, 2)
        grid.setColumnStretch(1, 1)
        layout.addLayout(grid)

        hint = label(
            "Тут — лише таблиця й зразки. Виконавець, папка результату, підписант і засвідчувач — "
            "у головному вікні, у картці «Реквізити цього прогону». Зразок витягу потрібен для "
            "витягів до управління: їх створено поруч із повідомленнями, якщо в наказі є такі пункти.",
            "MutedLabel",
        )
        hint.setWordWrap(True)
        layout.addWidget(hint)
        layout.addStretch(1)

        bottom = QHBoxLayout()
        bottom.addWidget(button("ⓘ Теги", "ghost", shell.show_message_tags))
        bottom.addStretch(1)
        bottom.addWidget(button("💾 Зберегти та закрити", "run", self.accept))
        layout.addLayout(bottom)

    def done(self, result: int) -> None:
        # Esc, хрестик і кнопка ведуть сюди — налаштування зберігаються завжди.
        self._shell.save_config()
        super().done(result)


def create_messages_app_class(legacy):
    return type("MessagesApp", (MessagesShellMixin, legacy.App), {"legacy": legacy})


def launch(legacy, argv: list[str] | None = None, splash=None) -> int:
    qt_app = QApplication.instance() or QApplication(list(argv if argv is not None else sys.argv))
    qt_app.setApplicationName("Повідомлення за наказом")
    theme.apply_theme(qt_app)
    install_qt_bridge(legacy)
    shell = create_messages_app_class(legacy)()
    theme.apply_theme(qt_app, shell.ui_theme.get())
    shell.main_window.show()
    if splash is not None:
        splash.finish(shell.main_window)
    return qt_app.exec()
