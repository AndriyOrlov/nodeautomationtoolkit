"""Вікно «Накази» — генератор наказів по особовому складу.

Чотири кроки в тому самому порядку, що й у пакеті `order_generator`:

    ① Індексація — папка зі старими наказами → локальний індекс;
    ② Генерація — план переміщення / документ / індекс → проєкт наказу;
    ③ Перевірка — що бракує і що не сходиться, простою мовою;
    ④ Збірка — .docx за правилами верстки витягів.

Логіка лежить у `order_generator`; тут лише інтерфейс і виклики. Поле
«Шаблон наказу» зберігається в конфіг разом з рештою полів; поки шаблону
немає, збірка робить чистий аркуш за геометрією додатка 53.

Генератор ще в роботі, тому в головному вікні він не вкладкою, а окремим
вікном за кнопкою «🔒 Накази» і тимчасовим паролем (рішення користувача
18.09.2026): у зібраній програмі ним поки не користуються. Це замок від
випадкового запуску, а не захист даних.
"""

from __future__ import annotations

import os
from pathlib import Path

from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFrame,
    QGridLayout,
    QInputDialog,
    QLineEdit,
    QMainWindow,
    QScrollArea,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from . import compat
from .widgets import (
    ActionStrip,
    Badge,
    OrderDropZone,
    SectionCard,
    bind_label,
    button,
    check_box,
    hint_bar,
    label,
    line_edit,
    make_table,
    run_params_card,
)

ITEM_COLUMNS = ["№", "Звання", "ПІБ", "З якої посади", "На яку посаду", "Джерела", "Зауваження"]
ITEM_WIDTHS = {0: 42, 1: 110, 2: 210, 3: 260, 4: 260, 5: 120}
PROBLEM_COLUMNS = ["", "Де", "Особа", "Що не так", "Що зробити"]
PROBLEM_WIDTHS = {0: 28, 1: 90, 2: 200, 3: 320}

PLAN_FILETYPES = [("Книга Excel", "*.xlsx *.xlsm"), ("Усі файли", "*.*")]
DOCUMENT_FILETYPES = [
    ("Кадрові документи", "*.docx *.doc *.rtf *.txt *.xlsx *.pdf *.jpg *.png"),
    ("Усі файли", "*.*"),
]
TEMPLATE_FILETYPES = [("Документ Word", "*.docx"), ("Усі файли", "*.*")]

#: Тимчасовий пароль до вікна наказів (генератор у роботі). Це не захист
#: даних, а замок від випадкового запуску в зібраній програмі.
ORDERS_PASSWORD = "2281488"

APPOINTMENT = "призначення"
DISMISSAL = "звільнення"
RANK = "присвоєння звання"

#: Поля ручного вводу: ключ запису, підпис, підказка, для якого виду наказу.
MANUAL_FIELDS = (
    ("rank", "Звання", "полковник", ""),
    ("surname", "Прізвище", "у називному: ІВАНЕНКО", ""),
    ("name", "Ім'я", "Олексій", ""),
    ("patronymic", "По батькові", "Вікторович", ""),
    ("ipn", "РНОКПП", "10 цифр", ""),
    ("birth", "Народження", "11.08.1976 або 1976", ""),
    ("education", "Освіта", "НАОУ (оср) у 2003 р.", APPOINTMENT),
    ("service_since", "У ЗС із", "08.1996", APPOINTMENT),
    ("current.text", "Посада, з якої", "у називному", ""),
    ("current.shpk", "шпк старої посади", "майор", APPOINTMENT),
    ("target.text", "Посада, на яку", "у називному", APPOINTMENT),
    ("target.shpk", "шпк нової посади", "підполковник", APPOINTMENT),
    ("target.vos", "ВОС", "0210003", APPOINTMENT),
    ("basis", "Підстава", "рапорт від 01.09.2026", ""),
    ("dismissal", "Підпункт звільнення", "1.а або а", DISMISSAL),
    ("destination", "Куди", "у запас / у відставку", DISMISSAL),
    ("service_calendar", "Вислуга календарна", "29 років 3 місяці", DISMISSAL),
    ("service_privileged", "Вислуга пільгова", "29 років 11 місяців або немає", DISMISSAL),
    ("registration", "На облік до", "Слобідського ОРТЦК та СП м. Харкова", DISMISSAL),
    ("uniform", "Право на форму", "так / ні", DISMISSAL),
    ("dismissal_note", "Додатковий рядок", "Чинність контракту припиняється 18.12.2026", DISMISSAL),
    ("new_rank", "Присвоюється звання", "майор", RANK),
    ("rank_seniority", "Вислуга у званні", "11 років", RANK),
    ("rank_since", "Строк рахувати з", "04.12.2026", RANK),
    ("rank_note", "Примітка до звання", "достроково на 6 місяців", RANK),
)


class ManualPersonDialog(QDialog):
    """Ручний ввід особи: ті самі поля, що й у записі `PersonRecord`."""

    def __init__(self, action: str, parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle(f"Особа вручну — наказ про {action}")
        self._edits: dict[str, QLineEdit] = {}
        layout = QVBoxLayout(self)
        grid = QGridLayout()
        grid.setHorizontalSpacing(8)
        grid.setVerticalSpacing(6)
        row = 0
        for key, title, hint, only_for in MANUAL_FIELDS:
            if only_for and only_for != action:
                continue
            edit = QLineEdit()
            edit.setPlaceholderText(hint)
            grid.addWidget(label(f"{title}:", "FieldLabel"), row // 2, (row % 2) * 2)
            grid.addWidget(edit, row // 2, (row % 2) * 2 + 1)
            self._edits[key] = edit
            row += 1
        grid.setColumnStretch(1, 1)
        grid.setColumnStretch(3, 1)
        layout.addLayout(grid)
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.resize(760, 380)

    def values(self) -> dict[str, str]:
        return {key: edit.text().strip() for key, edit in self._edits.items() if edit.text().strip()}


class OrdersWindow(QMainWindow):
    """Окреме вікно генератора наказів."""

    def __init__(self, page: QWidget, parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("Генератор наказів по особовому складу (у роботі)")
        scroll = QScrollArea()
        scroll.setObjectName("TabScroll")
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setWidget(page)
        self.setCentralWidget(scroll)
        self.resize(1240, 920)


class OrdersWindowMixin:
    """Інтерфейс і дії вікна «Накази». Домішується до `QtShellMixin`."""

    # ── Замок ────────────────────────────────────────────────────────────
    def ask_orders_password(self) -> str:
        """Питає пароль; окремим методом, щоб тести не відкривали модальне вікно."""
        value, accepted = QInputDialog.getText(
            self.main_window,
            "Генератор наказів",
            "Генератор наказів ще в роботі. Введіть пароль:",
            QLineEdit.EchoMode.Password,
        )
        return value if accepted else ""

    def open_orders_window(self):
        """Відкриває вікно наказів після пароля; у межах запуску питає один раз."""
        if not getattr(self, "_orders_unlocked", False):
            if self.ask_orders_password().strip() != ORDERS_PASSWORD:
                self.log("Генератор наказів ще в роботі: пароль не підійшов.")
                return
            self._orders_unlocked = True
        if getattr(self, "_orders_window", None) is None:
            self._orders_window = OrdersWindow(self._build_orders_page(), self.main_window)
        self._orders_window.show()
        self._orders_window.raise_()
        self._orders_window.activateWindow()
        return self._orders_window

    # ── Побудова ─────────────────────────────────────────────────────────
    def _build_orders_page(self) -> QWidget:
        page, layout = self._page()
        self._order_records = []
        self._order_draft = None
        self._order_check = None

        layout.addWidget(self._orders_step_one())
        layout.addWidget(self._orders_step_two())
        layout.addWidget(self._orders_requisites())
        layout.addWidget(self._orders_result_card(), 1)

        strip = ActionStrip()
        strip.add(self._lock(button("① Проіндексувати накази", "secondary", self.run_order_indexing)))
        strip.add(self._lock(button("② Скласти проєкт", "primary", self.run_order_compose)))
        strip.add(self._lock(button("③ Перевірити", "amber", self.run_order_check)))
        strip.add(self._lock(button("④ Зібрати документ", "run", self.run_order_assemble)))
        strip.add(button("📂 Папка результату", "secondary", self.open_new_order_output_folder))
        self._strip_status[3] = strip.finish().status
        layout.addWidget(strip)

        layout.addWidget(
            self._log_card("Журнал наказу", self.log_model, "log_text", self.clear_log), 1
        )
        return page

    def _orders_step_one(self) -> SectionCard:
        card = SectionCard("① Індекс попередніх наказів", accent="teal")
        card.add_action(self._lock(button("Папка наказів", "secondary", self.select_order_archive_folder)))
        card.add_action(self._lock(button("Папка індексу", "secondary", self.select_order_index_folder)))
        note = label(
            "Індекс потрібен, щоб брати з останнього наказу вивірену біографію особи "
            "й посаду, на яку її призначали торік — вона стає посадою, з якої призначають тепер. "
            "Індекс лишається на цьому комп'ютері.",
            "MutedLabel",
        )
        note.setWordWrap(True)
        card.body.addWidget(note)
        grid = QGridLayout()
        grid.setHorizontalSpacing(8)
        grid.addWidget(label("Накази:", "FieldLabel"), 0, 0)
        grid.addWidget(line_edit(self.new_order_archive_folder, "тека зі старими наказами"), 0, 1)
        grid.addWidget(label("Індекс:", "FieldLabel"), 1, 0)
        grid.addWidget(line_edit(self.new_order_index_folder, "тека, де лежатиме індекс"), 1, 1)
        grid.setColumnStretch(1, 1)
        card.body.addLayout(grid)
        card.body.addWidget(check_box("Перебудувати індекс заново", self.new_order_index_force))
        return card

    def _orders_step_two(self) -> SectionCard:
        card = SectionCard("② Джерела для пунктів", accent="indigo")
        card.add_action(self._lock(button("＋ План або таблиця", "primary", self.select_order_plan)))
        card.add_action(self._lock(button("＋ Документи", "secondary", self.select_order_documents)))
        card.add_action(self._lock(button("＋ Особа вручну", "secondary", self.add_manual_person)))
        card.add_divider()
        card.add_action(button("✖ Очистити", "ghost", self.clear_order_sources))
        card.body.addWidget(bind_label(label("", "SummaryLabel"), self.new_order_source_summary))
        card.body.addWidget(
            OrderDropZone(
                "⇪  Перетягніть сюди план чи список (.xlsx) або подання, рапорт, витяг",
                self.accept_order_source_path,
            )
        )
        return card

    def _orders_requisites(self) -> SectionCard:
        card = SectionCard("③ Реквізити наказу", accent="slate")

        kind_row = QGridLayout()
        kind_row.setHorizontalSpacing(8)
        kind_row.addWidget(label("Вид наказу:", "FieldLabel"), 0, 0)
        self.order_action_box = QComboBox()
        self.order_action_box.addItems([APPOINTMENT, DISMISSAL, RANK])
        self.order_action_box.setCurrentText(self.new_order_action.get() or APPOINTMENT)
        self.order_action_box.currentTextChanged.connect(self._on_order_action_changed)
        kind_row.addWidget(self.order_action_box, 0, 1)
        kind_row.addWidget(label("Підстава закону:", "FieldLabel"), 0, 2)
        kind_row.addWidget(
            line_edit(self.new_order_law_points, "для звільнення: пункту другого частини п'ятої статті 26"),
            0,
            3,
        )
        kind_row.setColumnStretch(1, 1)
        kind_row.setColumnStretch(3, 2)
        card.body.addLayout(kind_row)

        grid = QGridLayout()
        grid.setHorizontalSpacing(8)
        grid.setVerticalSpacing(6)
        fields = (
            ("Номер:", self.new_order_number, "525", 0, 0),
            ("Дата:", self.new_order_date, "17.09.2026", 0, 2),
            ("Розділ §:", self.new_order_section, "§ 1 — можна лишити порожнім", 0, 4),
            ("Пункти Положення:", self.new_order_points, "пункту 45", 1, 0),
            ("Тип о/с:", self.new_order_kind, "осіб офіцерського складу", 1, 2),
            ("Чиїх осіб:", self.new_order_unit, "шифр, скорочення або назва з таблиці", 1, 4),
            ("Призначити до:", self.new_order_target_unit, "коли всі йдуть в одну частину", 2, 0),
            ("Підстави в шапці:", self.new_order_bases, "через крапку з комою", 2, 2),
            ("Підстава в кінці:", self.new_order_footer_bases, "через крапку з комою", 2, 4),
        )
        for title, variable, hint, row, column in fields:
            grid.addWidget(label(title, "FieldLabel"), row, column)
            grid.addWidget(line_edit(variable, hint), row, column + 1)
        for column in (1, 3, 5):
            grid.setColumnStretch(column, 1)
        card.body.addLayout(grid)

        template_row = QGridLayout()
        template_row.setHorizontalSpacing(8)
        template_row.addWidget(label("Шаблон наказу:", "FieldLabel"), 0, 0)
        template_row.addWidget(
            line_edit(self.new_order_template_path, "не обов'язково — без нього аркуш чистий"), 0, 1
        )
        template_row.addWidget(
            button("📄", "secondary", self.select_new_order_template, "Обрати шаблон наказу"), 0, 2
        )
        template_row.addWidget(button("✖", "ghost", self.clear_new_order_template, "Прибрати шаблон"), 0, 3)
        template_row.setColumnStretch(1, 1)
        card.body.addLayout(template_row)
        card.body.addWidget(
            hint_bar(
                "У шаблоні місце для пунктів позначається тегом {{зміст}}; "
                "також підставляються {{номер_наказу}}, {{дата_наказу}}, {{підписант}}, {{виконавець}}."
            )
        )

        card.body.addWidget(
            run_params_card(
                "Реквізити цього прогону",
                executor_var=self.new_order_executor,
                folder_var=self.new_order_out_folder,
                on_pick_folder=self.select_new_order_output_folder,
                people=(
                    (
                        "Підписант наказу",
                        self.new_order_signer_position,
                        self.new_order_signer_rank,
                        self.new_order_signer_name,
                    ),
                ),
            )
        )
        return card

    def _orders_result_card(self) -> SectionCard:
        card = SectionCard("Проєкт наказу", accent="emerald")
        self._order_badge = Badge("", "emerald")
        self._order_badge.hide()
        card.add_badge(self._order_badge)
        card.add_action(button("👁 Показати текст", "sky", self.show_order_preview))
        self.order_items_table = make_table(
            ITEM_COLUMNS, selection="extended", widths=ITEM_WIDTHS, min_height=150
        )
        card.body.addWidget(self.order_items_table)
        self.order_problems_table = make_table(
            PROBLEM_COLUMNS, header_tone="amber", widths=PROBLEM_WIDTHS, min_height=110
        )
        card.body.addWidget(self.order_problems_table)
        return card

    # ── Вибір файлів і тек ───────────────────────────────────────────────
    def select_order_archive_folder(self):
        folder = compat.FileDialogBridge.askdirectory(title="Папка зі старими наказами")
        if folder:
            self.new_order_archive_folder.set(folder)
            if not self.new_order_index_folder.get():
                self.new_order_index_folder.set(os.path.join(folder, "order_index"))
            self.save_config()

    def select_order_index_folder(self):
        folder = compat.FileDialogBridge.askdirectory(title="Папка для індексу наказів")
        if folder:
            self.new_order_index_folder.set(folder)
            self.save_config()

    def select_order_plan(self):
        path = compat.FileDialogBridge.askopenfilename(
            title="План переміщення, план звільнення або список", filetypes=PLAN_FILETYPES
        )
        if path:
            self.new_order_plan_path.set(path)
            self._refresh_order_sources()

    def select_order_documents(self):
        paths = compat.FileDialogBridge.askopenfilenames(
            title="Подання, рапорти, витяги", filetypes=DOCUMENT_FILETYPES
        )
        if paths:
            self.new_order_document_paths = list(paths)
            self._refresh_order_sources()

    def accept_order_source_path(self, path: str):
        """Один перетягнутий файл: .xlsx — план, решта — документ."""
        if Path(path).suffix.casefold() in {".xlsx", ".xlsm"}:
            self.new_order_plan_path.set(path)
        else:
            self.new_order_document_paths = [*self.new_order_document_paths, path]
        self._refresh_order_sources()

    def clear_order_sources(self):
        self.new_order_plan_path.set("")
        self.new_order_document_paths = []
        self.new_order_manual_people = []
        self._refresh_order_sources()

    def _on_order_action_changed(self, value: str):
        self.new_order_action.set(value)
        self.save_config()
        self.log(f"Вид наказу: {value}.")

    def add_manual_person(self):
        """Ручний ввід особи — коли плану немає або особу треба дописати."""
        dialog = ManualPersonDialog(self.new_order_action.get() or APPOINTMENT, self.main_window)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        values = dialog.values()
        if not values.get("surname"):
            self.log("Особу не додано: без прізвища пункт скласти нема з чого.")
            return
        self.new_order_manual_people = [*self.new_order_manual_people, values]
        self.log(f"Додано особу вручну: {values.get('surname')}.")
        self._refresh_order_sources()

    def select_new_order_template(self):
        path = compat.FileDialogBridge.askopenfilename(
            title="Шаблон наказу", filetypes=TEMPLATE_FILETYPES
        )
        if path:
            self.new_order_template_path.set(path)
            self.save_config()

    def clear_new_order_template(self):
        self.new_order_template_path.set("")
        self.save_config()

    def select_new_order_output_folder(self):
        folder = compat.FileDialogBridge.askdirectory(title="Папка для готового наказу")
        if folder:
            self.new_order_out_folder.set(folder)
            self.save_config()

    def open_new_order_output_folder(self):
        folder = self.new_order_out_folder.get()
        if not folder or not os.path.isdir(folder):
            self.log("Папку результату ще не обрано.")
            return
        os.startfile(folder)  # noqa: S606 - штатне відкриття теки провідником

    def _refresh_order_sources(self):
        plan = self.new_order_plan_path.get()
        documents = self.new_order_document_paths
        parts = []
        if plan:
            parts.append(f"план: {os.path.basename(plan)}")
        if documents:
            parts.append(f"документів: {len(documents)}")
        if self.new_order_manual_people:
            parts.append(f"осіб вручну: {len(self.new_order_manual_people)}")
        self.new_order_source_summary.set(
            "; ".join(parts)
            if parts
            else "Джерела не обрано — план переміщення, документи або ручний ввід"
        )
        self.save_config()

    # ── Кроки ────────────────────────────────────────────────────────────
    def run_order_indexing(self):
        return self._guarded(self._run_order_indexing)

    def _run_order_indexing(self):
        from nodeautomationtoolkit.order_generator import pipeline

        orders, index = self.new_order_archive_folder.get(), self.new_order_index_folder.get()
        if not orders or not os.path.isdir(orders):
            self.log("Спершу оберіть папку зі старими наказами.")
            return
        if not index:
            index = os.path.join(orders, "order_index")
            self.new_order_index_folder.set(index)
        self.log(f"Індексую накази: {orders}")
        summary = pipeline.index_orders(
            orders, index, force=bool(self.new_order_index_force.get()), log=self.log
        )
        self.log(
            f"Готово: переглянуто {summary.scanned}, проіндексовано {summary.indexed}, "
            f"пунктів {summary.items}, без змін {summary.unchanged}, не вдалося {summary.failed}."
        )
        self.save_config()

    def run_order_compose(self):
        return self._guarded(self._run_order_compose)

    def _run_order_compose(self):
        from nodeautomationtoolkit.order_generator import pipeline

        index = self.new_order_index_folder.get() or ""
        if index and not os.path.isdir(index):
            index = ""
        records = []
        plan = self.new_order_plan_path.get()
        action = self.new_order_action.get() or APPOINTMENT
        try:
            if plan and action == APPOINTMENT:
                records += pipeline.records_from_plan(plan, index or None, log=self.log)
                if not records:
                    # Не план переміщення за додатком 16 — читаємо як звичайну
                    # таблицю, впізнаючи графи за заголовками.
                    self.log("Форму додатка 16 не впізнано — читаю таблицю за заголовками граф.")
                    records += pipeline.records_from_table(
                        plan, index or None, log=self.log, action=action
                    )
            elif plan:
                records += pipeline.records_from_table(
                    plan, index or None, log=self.log, action=action
                )
            if self.new_order_document_paths:
                records += pipeline.records_from_documents(
                    list(self.new_order_document_paths), index or None, log=self.log
                )
            if self.new_order_manual_people:
                records += pipeline.records_from_manual(
                    list(self.new_order_manual_people), action=action
                )
        except FileNotFoundError as error:
            self.log(f"Не знайдено файл: {error.filename or error}")
            return
        if not records:
            self.log(
                "Немає з чого складати наказ: оберіть план переміщення, документи "
                "або додайте особу вручну."
            )
            return

        self._order_records = records
        self._order_draft = pipeline.generate(records, self._order_params())
        self._show_order_items()
        self.log(f"Складено пунктів: {len(self._order_draft.items)}. Далі — перевірка.")
        self._run_order_check(quiet=True)

    def run_order_check(self):
        return self._guarded(self._run_order_check)

    def _run_order_check(self, quiet: bool = False):
        from nodeautomationtoolkit.order_generator import pipeline

        if self._order_draft is None:
            self.log("Спершу складіть проєкт наказу.")
            return
        self._order_check = pipeline.verify(self._order_draft)
        self._show_order_problems()
        errors, warnings = len(self._order_check.errors), len(self._order_check.warnings)
        if not errors and not warnings:
            self.log("Перевірка: зауважень немає.")
        else:
            self.log(f"Перевірка: помилок {errors}, попереджень {warnings}.")
            if not quiet:
                for line in self._order_check.lines():
                    self.log(line)

    def run_order_assemble(self):
        return self._guarded(self._run_order_assemble)

    def _run_order_assemble(self):
        from nodeautomationtoolkit.order_generator import pipeline

        if self._order_draft is None:
            self.log("Спершу складіть проєкт наказу.")
            return
        if self._order_check is None:
            self._run_order_check(quiet=True)
        folder = self.new_order_out_folder.get()
        if not folder:
            self.log("Оберіть папку результату.")
            return
        if self._order_check.errors:
            answer = compat.MessageBoxBridge.askyesno(
                "Є помилки",
                f"Перевірка знайшла помилок: {len(self._order_check.errors)}.\n"
                "Зібрати документ усе одно?",
            )
            if not answer:
                self.log("Збірку скасовано: спершу виправте помилки, позначені ✖.")
                return
        path = pipeline.assemble(
            self._order_draft,
            folder,
            self.new_order_template_path.get(),
            self._order_document_parts(),
        )
        self.log(f"✅ Наказ зібрано: {path}")
        self.save_config()

    # ── Показ ────────────────────────────────────────────────────────────
    def _order_params(self):
        from nodeautomationtoolkit.order_generator.compose import OrderParams

        def split(text: str) -> list[str]:
            return [piece.strip() for piece in str(text or "").split(";") if piece.strip()]

        return OrderParams(
            units_table=self.excel_path.get(),
            number=self.new_order_number.get(),
            date=self.new_order_date.get(),
            section=self.new_order_section.get(),
            points=self.new_order_points.get() or "пункту ___",
            kind=self.new_order_kind.get() or "осіб офіцерського складу",
            unit=self._full_unit(self.new_order_unit.get(), "Р"),
            target_unit=self._full_unit(self.new_order_target_unit.get()),
            bases=split(self.new_order_bases.get()),
            footer_bases=split(self.new_order_footer_bases.get()),
            action=self.new_order_action.get() or APPOINTMENT,
            law_points=self.new_order_law_points.get() or "пункту ___ частини ___ статті 26",
        )

    def _full_unit(self, text: str, case_label: str = "") -> str:
        """Коротку назву частини («А1234», «72 омбр») міняє на повну з таблиці."""
        from nodeautomationtoolkit.order_generator.units import full_unit_name

        table = self.excel_path.get()
        if not text.strip() or not table:
            return text
        try:
            name, how = full_unit_name(text, table, case_label)
        except Exception as error:  # таблиця недоступна чи має інший вигляд
            self.log(f"Таблицю частин прочитати не вдалося ({type(error).__name__}) — беру, що ввели.")
            return text
        if how:
            self.log(f"Частину знайдено в таблиці за {how}: {name}")
        else:
            self.log(f"У таблиці частин немає «{text}» — у наказ піде те, що ввели.")
        return name

    def _order_document_parts(self):
        from nodeautomationtoolkit.order_generator.build import OrderDocumentParts

        return OrderDocumentParts(
            signer_position=self.new_order_signer_position.get(),
            signer_rank=self.new_order_signer_rank.get(),
            signer_name=self.new_order_signer_name.get(),
            executor=self.new_order_executor.get(),
        )

    def _show_order_items(self):
        self.order_items_table.clear()
        for item in self._order_draft.items:
            record = item.record
            sources = sorted({source for source in record.sources().values() if source})
            values = [
                str(item.number),
                str(record.rank),
                record.full_name,
                str(record.current.text),
                str(record.target.text),
                ", ".join(sources),
                "; ".join(record.problems),
            ]
            self.order_items_table.addTopLevelItem(QTreeWidgetItem(values))
        self._order_badge.setText(f"Пунктів: {len(self._order_draft.items)}")
        self._order_badge.setVisible(bool(self._order_draft.items))

    def _show_order_problems(self):
        self.order_problems_table.clear()
        for problem in self._order_check.problems:
            mark = "✖" if problem.level == "помилка" else "⚠"
            self.order_problems_table.addTopLevelItem(
                QTreeWidgetItem([mark, problem.where, problem.person, problem.what, problem.how])
            )

    def show_order_preview(self):
        if self._order_draft is None:
            self.log("Проєкт наказу ще не складено.")
            return
        text = self._order_draft.text.replace(" ", " ")
        compat.MessageBoxBridge.showinfo("Текст наказу", text[:4000])
