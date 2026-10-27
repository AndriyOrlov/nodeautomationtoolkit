"""Головне вікно Qt-оболонки генератора (PySide6-Fluent-Widgets, світла й темна теми).

Оболонка — це міксин поверх `generate_extracts.App`: `App.__init__`,
конфіг, маршрутизація, робота з Word і пакети лишаються тими самими.
Міксин підміняє лише те, що в Tk-версії малює інтерфейс.

Каркас — `FluentWindow`: зліва навігація (Примірники, Витяги, Повідомлення;
внизу — Інструкція, Зразки, Накази, тема), праворуч сторінка й рядок стану.
"""

from __future__ import annotations

import os
import sys
from typing import Callable

from PySide6.QtCore import QObject, Qt, Signal
from PySide6.QtGui import QCloseEvent, QDragEnterEvent, QDragMoveEvent, QDropEvent, QGuiApplication
from PySide6.QtWidgets import (
    QApplication,
    QFrame,
    QHBoxLayout,
    QScrollArea,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from . import compat, theme
from .compare_window import CompareWindow
from .fluent import FluentIcon, FluentWindow, NavigationItemPosition, TitleLabel
from .instruction_dialog import InstructionDialog
from nodeautomationtoolkit.generator.printing import default_printer

from .orders_window import OrdersWindowMixin
from .print_dialog import PrintPauseDialog
from .samples_dialog import SamplesDialog
from .widgets import (
    COPY_RESULT_COLUMNS,
    COPY_RESULT_WIDTHS,
    ORDER_COLUMNS,
    ORDER_WIDTHS,
    ActionStrip,
    OrderDropZone,
    Badge,
    LogConsole,
    PillDelegate,
    SectionCard,
    StatusStrip,
    bind_label,
    button,
    check_box,
    hint_bar,
    label,
    line_edit,
    make_orders_renderer,
    make_table,
    plural_files,
    render_analysis_row,
    render_copy_result,
    run_params_card,
)

RESULT_TABS = (
    (
        "calculation",
        "Розсилка",
        ["Військова частина / Відправник", "Пункти витягу", "Кількість"],
        "— Очікується розрахунок розсилки —",
    ),
    ("unmatched", "Пропущені", ["Пункт", "Текст пункту", "Причина"], "— Пропущені пункти відсутні —"),
    (
        "routing",
        "Контроль маршрутизації",
        ["Пункт", "Збіги з таблиці", "Застосовані правила", "Підсумкові адресати"],
        "— Очікується аналіз маршрутизації —",
    ),
    ("layout", "Макет витягів", ["Стан макета"], "— Очікується формування витягів —"),
)

#: Сторінки навігації в порядку, на який спирається логіка (`handle_drag_and_drop`):
#: 0 — примірники, 1 — витяги, 2 — повідомлення.
PAGES = (
    ("copiesPage", FluentIcon.COPY, "Примірники 2/3", "Примірники наказів № 2 і № 3"),
    ("extractsPage", FluentIcon.DOCUMENT, "Розрахунок і витяги", "Розрахунок розсилки та витяги"),
    ("messagesPage", FluentIcon.MAIL, "Повідомлення", "Шифровані повідомлення за наказом"),
    ("printPage", FluentIcon.PRINT, "Друк", "Друк пакета"),
)


def install_qt_bridge(legacy) -> None:
    """Перемикає модуль генератора з Tk на Qt-реалізацію тієї самої поверхні."""
    legacy.tk = compat.make_tk_namespace()
    legacy.messagebox = compat.MessageBoxBridge
    legacy.filedialog = compat.FileDialogBridge
    legacy.DISABLED = compat.DISABLED
    legacy.NORMAL = compat.NORMAL


def _package_version() -> str:
    try:
        import nodeautomationtoolkit

        version = getattr(nodeautomationtoolkit, "__version__", "")
        if version:
            return str(version)
        from importlib.metadata import version as installed_version

        return installed_version("nodeautomationtoolkit")
    except Exception:
        return ""


def _short_folder(path: str) -> str:
    parts = [part for part in os.path.normpath(path).split(os.sep) if part]
    return os.sep.join(parts[-2:]) if len(parts) > 2 else os.path.normpath(path)


class DropFluentWindow(FluentWindow):
    """Вікно Fluent, що приймає перетягнуті файли й зберігає конфіг під час закриття.

    Під сторінками — рядок стану (у FluentWindow його немає, тож вставляємо
    в його ж розкладку), анімацію перемикання сторінок вимкнено: логіка
    читає поточну сторінку одразу після перемикання.
    """

    def __init__(self, shell):
        super().__init__()
        self._shell = shell
        self.setAcceptDrops(True)
        # Без Mica: вікно малює власний фон теми однаково на Windows 10 і 11
        # (з Mica в темній темі рамка й навігація лишались світлими там, де ефекту немає).
        self.setMicaEffectEnabled(False)
        self.stackedWidget.setAnimationEnabled(False)
        self.navigationInterface.setExpandWidth(230)
        # Стрілка «назад» потрібна лише для вкладених розділів — у нас їх немає.
        self.navigationInterface.setReturnButtonVisible(False)

    def attach_status_bar(self, widget: QWidget) -> None:
        self.widgetLayout.removeWidget(self.stackedWidget)
        column = QVBoxLayout()
        column.setContentsMargins(0, 0, 0, 0)
        column.setSpacing(0)
        column.addWidget(self.stackedWidget, 1)
        column.addWidget(widget)
        self.widgetLayout.addLayout(column)

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
        self._shell.on_window_close()
        super().closeEvent(event)


#: Стара назва — для коду, що імпортує її напряму.
GeneratorWindow = DropFluentWindow


class PageSwitcher(QObject):
    """Підмножина API QTabWidget поверх навігації Fluent.

    `NotebookBridge` (логіка: `notebook.select()/index()`) і `_refresh_badges`
    працюють із вкладками; тут ті самі виклики перемикають сторінки вікна.
    """

    currentChanged = Signal(int)

    def __init__(self, window: DropFluentWindow, pages: list[QWidget]):
        super().__init__(window)
        self._window = window
        self._pages = pages
        window.stackedWidget.currentChanged.connect(self.currentChanged)

    def count(self) -> int:
        return len(self._pages)

    def widget(self, index: int) -> QWidget | None:
        return self._pages[index] if 0 <= index < len(self._pages) else None

    def indexOf(self, widget) -> int:  # noqa: N802 - API Qt
        return self._pages.index(widget) if widget in self._pages else -1

    def currentWidget(self) -> QWidget:  # noqa: N802
        return self._window.stackedWidget.currentWidget()

    def currentIndex(self) -> int:  # noqa: N802
        return self.indexOf(self.currentWidget())

    def setCurrentWidget(self, widget: QWidget) -> None:  # noqa: N802
        if widget in self._pages:
            self._window.switchTo(widget)

    def setCurrentIndex(self, index: int) -> None:  # noqa: N802
        page = self.widget(index)
        if page is not None:
            self._window.switchTo(page)

    def tabText(self, index: int) -> str:  # noqa: N802
        return PAGES[index][2] if 0 <= index < len(PAGES) else ""


class QtShellMixin(OrdersWindowMixin):
    """Qt-інтерфейс для `generate_extracts.App`. Порядок сторінок незмінний
    (0 — примірники, 1 — витяги, 2 — повідомлення): на нього
    спирається `handle_drag_and_drop`."""

    legacy = None  # модуль generate_extracts; задається в create_qt_app_class
    # Нові частини з наказу дописуються в таблицю заготовками (AGENT.md 9.5.7).
    ADD_NEW_UNITS_TO_TABLE = True

    def __init__(self):
        self.main_window = DropFluentWindow(self)
        compat.set_dialog_parent(self.main_window)
        self.log_model = compat.LogModel()
        self.p2_log_model = compat.LogModel()
        self._busy_depth = 0
        self._lockable_buttons: list[compat.TkButtonMixin] = []
        self._samples_dialog: SamplesDialog | None = None
        self._instruction_dialog: InstructionDialog | None = None
        self._compare_windows: list[CompareWindow] = []
        self._ui_ready = False
        super().__init__(compat.RootBridge(self.main_window))
        self.main_window.resize(1400, 940)
        self.main_window.setMinimumSize(1100, 720)
        self._ui_ready = True
        self._refresh_badges()

    # ═════════════════════════════════════════════════════════════════════
    # Побудова інтерфейсу
    # ═════════════════════════════════════════════════════════════════════
    def create_widgets(self):
        self.log_text = compat.TextBridge(self.log_model)
        self.p2_log_text = compat.TextBridge(self.p2_log_model)
        self._tab_badges: dict[int, Badge] = {}

        builders = (
            self._build_copies_tab,
            self._build_extracts_tab,
            self._build_messages_tab,
            self._build_print_tab,
        )
        pages = []
        for index, ((name, icon, nav_text, title), build) in enumerate(zip(PAGES, builders)):
            page = self._scroll_page(self._page_with_header(index, title, build()))
            page.setObjectName(name)
            self.main_window.addSubInterface(page, icon, nav_text)
            pages.append(page)
        self.tab_copies_page, self.tab_extracts_page, self.tab_messages_page, self.tab_print_page = pages
        self.tab_copies = self.tab_copies_page
        self.tab_extracts = self.tab_extracts_page
        self.tab_messages = self.tab_messages_page

        self._tabs = PageSwitcher(self.main_window, pages)
        self.notebook = compat.NotebookBridge(self._tabs)
        self._build_navigation_footer()

        self.status_strip = StatusStrip()
        self.main_window.attach_status_bar(self.status_strip)
        self._ribbon_info = self.status_strip.info

        refresh = lambda *_args: self._refresh_badges()  # noqa: E731
        for variable in (
            self.source_summary,
            self.p2_source_summary,
            self.template_path,
            self.p2_back_page_path,
            self.message_content_template_path,
            self.duplex_2up_layout,
            self.p2_orders_folder,
        ):
            variable.subscribe(refresh)
        self._tabs.currentChanged.connect(refresh)
        self.progress_text.subscribe(lambda value: self.status_strip.set_progress(caption=value))
        self.progress_value.subscribe(lambda value: self.status_strip.set_progress(percent=value))
        theme.on_change(self._on_theme_changed)

        self._refresh_orders_view()
        self._refresh_p2_orders_view()
        self.log("Готовий до роботи. Оберіть наказ (або кілька) і натисніть потрібну дію.")

    def _build_navigation_footer(self) -> None:
        """Нижня частина навігації: довідка, зразки, генератор наказів, тема."""
        nav = self.main_window.navigationInterface
        bottom = NavigationItemPosition.BOTTOM
        self._nav_items = {
            "instruction": nav.addItem(
                "instruction", FluentIcon.HELP, "Інструкція",
                onClick=self.open_instruction_window, selectable=False, position=bottom,
            ),
            "samples": nav.addItem(
                "samples", FluentIcon.SETTING, "Зразки та шаблони",
                onClick=self.open_samples_window, selectable=False, position=bottom,
            ),
            "orders": nav.addItem(
                "orders", FluentIcon.EDIT, "Накази (у роботі)",
                onClick=self.open_orders_window, selectable=False, position=bottom,
                tooltip="Генератор наказів (у роботі)",
            ),
            "theme": nav.addItem(
                "theme", FluentIcon.CONSTRACT, "Світла / темна тема",
                onClick=self.toggle_theme, selectable=False, position=bottom,
            ),
        }

    def _lock(self, widget: compat.TkButtonMixin) -> compat.TkButtonMixin:
        self._lockable_buttons.append(widget)
        return widget

    @staticmethod
    def _scroll_page(page: QWidget) -> QScrollArea:
        scroll = QScrollArea()
        scroll.setObjectName("TabScroll")
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setWidget(page)
        return scroll

    @staticmethod
    def _page() -> tuple[QWidget, QVBoxLayout]:
        page = QWidget()
        page.setObjectName("TabPage")
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)
        return page, layout

    def _page_with_header(self, index: int, title: str, body: QWidget) -> QWidget:
        """Заголовок сторінки з лічильником черги над її вмістом."""
        page = QWidget()
        page.setObjectName("TabPage")
        layout = QVBoxLayout(page)
        layout.setContentsMargins(28, 8, 28, 16)
        layout.setSpacing(12)
        header = QHBoxLayout()
        header.setSpacing(10)
        header.addWidget(TitleLabel(title))
        badge = Badge("", "teal")
        badge.hide()
        self._tab_badges[index] = badge
        header.addWidget(badge, 0, Qt.AlignmentFlag.AlignVCenter)
        header.addStretch(1)
        layout.addLayout(header)
        layout.addWidget(body, 1)
        return page

    def _orders_table(
        self,
        on_marks: Callable[[], None],
        header_tone: str = "slate",
        min_height: int = 130,
        selectable: bool = False,
    ):
        table = make_table(
            ORDER_COLUMNS,
            selection="single" if selectable else "none",
            header_tone=header_tone,
            stretch_column=2,
            widths=ORDER_WIDTHS,
            min_height=min_height,
        )
        table.setItemDelegateForColumn(5, PillDelegate(table))
        bridge = compat.TreeBridge(table, make_orders_renderer(self.legacy.extract_metadata_from_filename))
        bridge.mark_callbacks.append(on_marks)
        bridge.changed_callbacks.append(self._refresh_badges)
        return table, bridge

    def _log_card(self, title: str, model: compat.LogModel, text_bridge_name: str, clear: Callable[[], None]):
        card = SectionCard(title, accent="slate", variant="log")
        card.add_action(
            button("📋 Скопіювати", "ghost", lambda: self.copy_log(getattr(self, text_bridge_name)))
        )
        card.add_action(
            button("Знеособлено", "ghost", lambda: self.copy_log(getattr(self, text_bridge_name), redacted=True))
        )
        card.add_action(button("🗑 Очистити", "ghost", clear))
        card.body.addWidget(LogConsole(model))
        return card

    # ── Сторінка 1: примірники ───────────────────────────────────────────
    def _build_copies_tab(self) -> QWidget:
        page, layout = self._page()

        orders = SectionCard("Накази для примірників", accent="indigo")

        # Явний перемикач джерела: окремі накази (один або кілька) чи вся тека.
        # Без нього одинарний режим був непомітним, а після перезапуску ще й
        # «зникав» — режим не зберігався в конфіг.
        self._p2_mode_files = orders.add_badge(
            self._lock(button("📄 Окремі накази", "segment", tooltip="Один або кілька обраних файлів наказів"))
        )
        self._p2_mode_folder = orders.add_badge(
            self._lock(button("📁 Уся тека", "segment", tooltip="Усі накази з обраної теки"))
        )
        for mode_button in (self._p2_mode_files, self._p2_mode_folder):
            mode_button.setCheckable(True)

        def show_mode() -> None:
            folder_mode = self.p2_source_mode.get() == "folder"
            self._p2_mode_folder.setChecked(folder_mode)
            self._p2_mode_files.setChecked(not folder_mode)

        def use_files_mode() -> None:
            if self.p2_manual_order_paths or self._p2_manual_paths():
                self.p2_source_mode.set("file")
                self._refresh_p2_orders_view()
                self.save_config()
            else:
                self.select_p2_file()
            show_mode()

        def use_folder_mode() -> None:
            folder = self.p2_orders_folder.get()
            if folder and os.path.isdir(folder):
                self.p2_source_mode.set("folder")
                self._refresh_p2_orders_view()
                self.save_config()
            else:
                self.select_p2_folder()
            show_mode()

        self._p2_mode_files.clicked.connect(lambda _checked=False: use_files_mode())
        self._p2_mode_folder.clicked.connect(lambda _checked=False: use_folder_mode())
        self.p2_source_mode.subscribe(lambda _value: show_mode())
        show_mode()

        self._p2_queue_badge = orders.add_badge(Badge("", "emerald"))
        orders.add_action(self._lock(button("＋ Обрати накази", "primary", self.select_p2_file)))
        orders.add_action(self._lock(button("📁 Обрати папку", "secondary", self.select_p2_folder)))
        orders.add_action(self._lock(button("🗑 Очистити", "ghost", self.clear_p2_orders)))
        table, self.p2_orders_tree = self._orders_table(lambda: self._refresh_p2_source_summary())
        orders.body.addWidget(table)
        orders.body.addWidget(hint_bar("Формується № 2 · № 3 — той самий, з «Прим. № 3», лише під час друку"))
        orders.body.addWidget(bind_label(label("", "SummaryLabel"), self.p2_source_summary))
        layout.addWidget(orders)

        def copies_options(row: QHBoxLayout) -> None:
            row.addWidget(check_box("Режим превʼю (повільно, з видимим Word)", self.p2_preview))
            row.addWidget(label("пауза, сек:", "FieldLabel"))
            delay = line_edit(self.p2_preview_delay)
            delay.setFixedWidth(64)
            row.addWidget(delay)
            row.addSpacing(12)
            row.addWidget(button("⇊ Засвідчувач як у витягах", "secondary", self.copy_certifier_from_extracts))
            row.addStretch(1)

        layout.addWidget(
            run_params_card(
                "Реквізити цього прогону",
                executor_var=self.p2_executor,
                folder_var=self.p2_out_folder,
                on_pick_folder=self.select_p2_out_folder,
                executor_hint="порожньо = з витягів",
                people=(
                    (
                        "Засвідчувач примірників («Згідно з оригіналом»)",
                        self.p2_certifier_position,
                        self.p2_certifier_rank,
                        self.p2_certifier_name,
                    ),
                ),
                options=copies_options,
            )
        )

        strip = ActionStrip()
        self.btn_run_p2 = self._lock(
            button("⚡ Сформувати примірники (заміна останнього аркуша)", "run", self.run_generate_copies)
        )
        strip.add(self.btn_run_p2)
        strip.add(button("ⓘ Правила примірника 2", "secondary", self.show_p2_info))
        strip.add(button("🔍 Порівняти з еталоном", "secondary", self.open_compare_copies))
        self._strip_status = {0: strip.finish().status}
        layout.addWidget(strip)

        results = SectionCard("Сформовані примірники наказів", accent="emerald", variant="results")
        self._p2_results_count = results.add_action(label("", "MutedLabel"))
        table = make_table(
            COPY_RESULT_COLUMNS,
            selection="extended",
            header_tone="teal",
            stretch_column=1,
            widths=COPY_RESULT_WIDTHS,
            min_height=180,
        )
        delegate = PillDelegate(table)
        table.setItemDelegateForColumn(4, delegate)
        table.setItemDelegateForColumn(5, delegate)
        table.itemDoubleClicked.connect(lambda *_args: self.open_selected_copy_in_word())
        self.p2_tree = compat.TreeBridge(table, render_copy_result)
        self.p2_tree.changed_callbacks.append(self._refresh_badges)
        self.p2_tree.insert(
            "", compat.END, values=("1", "— Очікується запуск формування примірників —", "—", "—", "—", "—", "—")
        )
        results.body.addWidget(table, 1)
        footer = QHBoxLayout()
        footer.addWidget(
            self._lock(button("→ Передати вибрані у витяги", "secondary", self.transfer_selected_copy_to_extracts))
        )
        footer.addStretch(1)
        footer.addWidget(button("☑ Вибрати всі", "ghost", self.select_all_copies))
        footer.addWidget(button("◻ Зняти вибір", "ghost", self.deselect_all_copies))
        footer.addWidget(button("📂 Папка результату", "secondary", self.open_p2_output_folder))
        footer.addWidget(button("W  Відкрити у Word", "secondary", self.open_selected_copy_in_word))
        results.body.addLayout(footer)
        layout.addWidget(results, 1)

        layout.addWidget(self._log_card("Журнал примірників", self.p2_log_model, "p2_log_text", self.p2_log_model.clear))
        return page

    # ── Сторінка 2: розрахунок і витяги ──────────────────────────────────
    def _build_extracts_tab(self) -> QWidget:
        page, layout = self._page()

        orders = SectionCard("Накази для обробки", accent="indigo")
        self._orders_queue_badge = orders.add_badge(Badge("", "emerald"))
        orders.add_action(self._lock(button("＋ Обрати накази", "primary", self.select_orders)))
        orders.add_action(self._lock(button("📁 Обрати папку", "secondary", self.select_orders_folder)))
        orders.add_action(self._lock(button("🗑 Очистити", "ghost", self.clear_orders)))
        table, self.orders_tree = self._orders_table(
            lambda: self._refresh_source_summary(), selectable=True
        )
        table.itemSelectionChanged.connect(
            lambda: self._activate_order_from_tree(self.orders_tree, alongside=False)
        )
        orders.body.addWidget(table)
        orders.body.addWidget(hint_bar())
        summary = bind_label(label("", "SummaryLabel"), self.source_summary)
        orders.body.addWidget(summary)
        layout.addWidget(orders)

        copies = SectionCard("Примірники № 2 з останнього пакета", accent="emerald")
        self._copies_queue_badge = copies.add_badge(Badge("", "emerald"))
        copies.add_badge(Badge("мають перевагу над списком наказів", "slate"))
        table, self.copy_two_tree = self._orders_table(
            lambda: self._refresh_source_summary(),
            header_tone="emerald",
            min_height=92,
            selectable=True,
        )
        table.itemSelectionChanged.connect(
            lambda: self._activate_order_from_tree(self.copy_two_tree, alongside=True)
        )
        self.copy_two_tree.insert(
            "", compat.END, values=("—", "Спершу сформуйте примірники № 2 на сторінці «Примірники 2/3».")
        )
        copies.body.addWidget(table)
        layout.addWidget(copies)

        def extract_options(row: QHBoxLayout) -> None:
            row.addWidget(check_box("Друк «2 сторінки на 1 аркуш»", self.duplex_2up_layout))
            row.addSpacing(12)
            row.addWidget(check_box("Групувати підпорядковані частини по Корпусах", self.group_corps_var))
            row.addStretch(1)

        layout.addWidget(
            run_params_card(
                "Реквізити цього прогону",
                executor_var=self.executor,
                folder_var=self.out_folder,
                on_pick_folder=self.select_folder,
                people=(
                    (
                        "Підписант оригіналу наказу (зчитується автоматично після пунктів)",
                        self.order_signer_position,
                        self.order_signer_rank,
                        self.order_signer_name,
                    ),
                    (
                        "Засвідчувач витягів («Згідно з оригіналом»)",
                        self.certifier_position,
                        self.certifier_rank,
                        self.certifier_name,
                    ),
                ),
                options=extract_options,
            )
        )

        strip = ActionStrip()
        self.btn_calc = self._lock(button("📊 1. Розрахувати розсилку", "secondary", self.run_rozrahunok_action))
        self.btn_extracts = self._lock(button("⚡ 2. Створити всі витяги", "run", self.run_extracts_action))
        self.btn_management_extracts = self._lock(
            button("🏛 3. Витяги до управління", "secondary", self.run_management_extracts_action)
        )
        self.btn_management_extracts.setToolTip(
            "Створити окремий файл витягів до управління без адресатів і компонування 2 на 1"
        )
        self.btn_full_cycle = self._lock(button("⇄ 4. Повний цикл", "secondary", self.run_full_cycle))
        self.btn_full_cycle.setToolTip(
            "Примірники № 2 → розрахунок розсилки → адресні витяги + витяги до управління"
        )
        for widget in (self.btn_calc, self.btn_extracts, self.btn_management_extracts, self.btn_full_cycle):
            strip.add(widget)
        self.btn_order_review = self._lock(button("🎓 Перевірити наказ", "secondary", self.run_order_review_action))
        self.btn_order_review.setToolTip(
            "Перевірити обрані накази без еталона: нумерація, адресати, РНОКПП, підписант, "
            "історія осіб і написання посад за попередніми наказами. Звіт і позначена копія — "
            "у теці «Перевірка» поруч із наказом; оригінал не змінюється."
        )
        strip.add(self.btn_order_review)
        strip.add(button("📂 Папка результату", "ghost", self.open_extracts_output_folder))
        strip.add(button("ⓘ Теги шаблону", "ghost", self.show_template_tags))
        strip.add(button("🔍 Порівняти з еталоном", "ghost", self.open_compare_extracts))
        self._strip_status[1] = strip.finish().status
        layout.addWidget(strip)

        results = SectionCard("Результати аналізу", accent="teal", variant="results")
        self._results_tabs = QTabWidget()
        self._results_tabs.setObjectName("ResultTabs")
        self._results_tabs.setDocumentMode(True)
        self.results_notebook = compat.NotebookBridge(self._results_tabs)
        self.result_views = {}
        self.result_tabs = {}
        for key, title, columns, placeholder in RESULT_TABS:
            view = make_table(columns, selection="single", header_tone="teal", min_height=170)
            bridge = compat.TreeBridge(view, render_analysis_row)
            bridge["columns"] = columns
            bridge.insert("", compat.END, values=(placeholder, *["—"] * (len(columns) - 1)))
            self._results_tabs.addTab(view, title)
            self.result_views[key] = bridge
            self.result_tabs[key] = view
        results.body.addWidget(self._results_tabs, 1)
        layout.addWidget(results, 1)

        layout.addWidget(self._log_card("Журнал виконання", self.log_model, "log_text", self.clear_log))
        return page

    # ── Сторінка 3: повідомлення ─────────────────────────────────────────
    def _build_messages_tab(self) -> QWidget:
        page, layout = self._page()

        source = SectionCard("Наказ для повідомлень", accent="indigo")
        source.add_action(self._lock(button("＋ Обрати один наказ", "primary", self.select_message_order)))
        source.body.addWidget(
            label("Повідомлення формуються по одному наказу за запуск.", "MutedLabel")
        )
        source.body.addWidget(bind_label(label("", "SummaryLabel"), self.message_source_summary))
        source.body.addWidget(
            OrderDropZone(
                "Перетягніть наказ сюди — одразу створяться повідомлення",
                self.generate_messages_for_dropped_order,
            )
        )
        tags_note = label(
            "Стандартні теги: {{номер_наказу}}/{{номер}}, {{дата_наказу}}/{{дата}}, {{кому_список}}, "
            "{{куди}}, {{тцк чі вч}}, {{виконавець}} та спільні теги підписанта/затверджувача. "
            "Кожен {{кому_список}} у таблиці — один унікальний адресат. У другому зразку: "
            "{{зміст_шифр}}. Невпізнані відкриті назви частин виділяються жовтим.",
            "MutedLabel",
        )
        tags_note.setWordWrap(True)
        source.body.addWidget(tags_note)
        layout.addWidget(source)

        def message_options(row: QHBoxLayout) -> None:
            skip_box = check_box(
                "Не додавати внутрішнє переміщення в управлінні",
                self.message_skip_internal_management,
            )
            skip_box.setToolTip(
                "Пункт, де «управління» є і в посаді «звідки», і в посаді «КУДИ», "
                "не переноситься у шифрований зміст. Витягів до управління це не стосується."
            )
            row.addWidget(skip_box)
            row.addStretch(1)

        layout.addWidget(
            run_params_card(
                "Реквізити цього прогону",
                executor_var=self.message_executor,
                folder_var=self.message_out_folder,
                on_pick_folder=self.select_message_output_folder,
                options=message_options,
            )
        )

        strip = ActionStrip()
        self.btn_generate_messages = self._lock(
            button("✉ Створити повідомлення", "run", self.run_generate_messages)
        )
        strip.add(self.btn_generate_messages)
        strip.add(button("📂 Папка результату", "ghost", self.open_message_output_folder))
        strip.add(button("🔍 Порівняти з еталоном", "ghost", self.open_compare_messages))
        strip.add(button("ⓘ Теги повідомлень", "ghost", self.show_message_tags))
        self._strip_status[2] = strip.finish().status
        layout.addWidget(strip)

        # Повідомлення пишуть у спільний журнал (як у Tk), тож тут — друге
        # відображення тієї самої моделі.
        layout.addWidget(self._log_card("Журнал повідомлень", self.log_model, "log_text", self.clear_log), 1)
        return page

    # ── Сторінка 4: друк пакета ──────────────────────────────────────────
    def _build_print_tab(self) -> QWidget:
        page, layout = self._page()

        source = SectionCard("Пакет для друку", accent="indigo")
        source.add_action(self._lock(button("📁 Обрати теку", "secondary", self.select_print_folder)))
        source.add_action(button("⇄ Оновити чергу", "ghost", self.refresh_print_queue))
        row = QHBoxLayout()
        row.setSpacing(8)
        row.addWidget(label("Тека пакета:", "FieldLabel"))
        row.addWidget(line_edit(self.print_folder, "порожньо — тека примірників № 2"), 1)
        source.body.addLayout(row)
        table = make_table(
            ["НАКАЗ", "ЩО ДРУКУЄТЬСЯ", "ФАЙЛ"],
            selection="none",
            stretch_column=2,
            widths={0: 220, 1: 320},
            min_height=220,
        )
        self.print_queue_tree = compat.TreeBridge(table, render_analysis_row)
        source.body.addWidget(table)
        layout.addWidget(source)

        settings = SectionCard("Як друкувати", accent="slate")
        settings.body.addWidget(check_box("Друкувати витяги до управління", self.print_management))
        hint_row = QHBoxLayout()
        hint_row.setSpacing(8)
        hint_row.addWidget(label("Як класти в ручну подачу:", "FieldLabel"))
        hint_row.addWidget(line_edit(self.print_feed_hint, "наприклад: лицем догори, верхом до принтера"), 1)
        settings.body.addLayout(hint_row)
        printer = default_printer()
        settings.body.addWidget(
            label(f"Принтер: {printer or 'не визначено'} — принтер Windows за замовчуванням.", "MutedLabel")
        )
        rules = label(
            "Для кожного наказу по черзі: витяги (2 сторінки на аркуш) → витяги до управління (якщо "
            "увімкнено; 2 примірники, кожен витяг на окремому аркуші) → примірник № 2: лицьові сторінки, потім верхній аркуш у ручну подачу й задня "
            "сторінка на звороті → те саме для примірника № 3 (у № 2 «Прим. № 2» міняється на № 3 лише "
            "для друку). Після кожного кроку — пауза: Пробіл або Enter — далі, Esc — зупинити.",
            "MutedLabel",
        )
        rules.setWordWrap(True)
        settings.body.addWidget(rules)
        layout.addWidget(settings)

        strip = ActionStrip()
        strip.add(self._lock(button("🖨 Друкувати пакет", "run", self.run_print_action)))
        strip.add(self._lock(button("📄 Перевірити у PDF", "secondary", self.run_print_check_pdf)))
        strip.add(self._lock(button("🖨 Пробний аркуш (ручна подача)", "secondary", self.run_print_test_sheet)))
        self._strip_status[3] = strip.finish().status
        layout.addWidget(strip)

        layout.addWidget(self._log_card("Журнал друку", self.log_model, "log_text", self.clear_log), 1)
        self.print_management.subscribe(lambda _value: self.refresh_print_queue())
        self.refresh_print_queue()
        return page

    def refresh_print_queue(self) -> None:
        bridge = getattr(self, "print_queue_tree", None)
        if bridge is None:
            return
        bridge.delete(*bridge.get_children())
        rows = self.print_queue_rows()
        for row in rows:
            bridge.insert("", compat.END, values=row)
        if not rows:
            bridge.insert("", compat.END, values=("—", "— Оберіть теку пакета: там, де примірники № 2 і витяги —", "—"))

    def select_print_folder(self) -> None:
        start = self.print_folder.get() or self.p2_out_folder.get()
        path = self.legacy.filedialog.askdirectory(initialdir=start if start and os.path.isdir(start) else "")
        if path:
            self.print_folder.set(os.path.abspath(path))
            self.save_config()
            self.refresh_print_queue()

    def _print_pause(self, done: str, after: str, upcoming: str) -> bool:
        dialog = PrintPauseDialog(done, after, upcoming, self.main_window)
        theme.use_dark_title_bar(dialog)
        return dialog.exec() == PrintPauseDialog.DialogCode.Accepted

    # ═════════════════════════════════════════════════════════════════════
    # Тема
    # ═════════════════════════════════════════════════════════════════════
    def toggle_theme(self) -> None:
        mode = theme.LIGHT if theme.is_dark() else theme.DARK
        if hasattr(self, "ui_theme"):
            self.ui_theme.set(mode)
        theme.apply_theme(QApplication.instance(), mode)
        self.save_config()

    def _on_theme_changed(self) -> None:
        compat.TreeBridge.rerender_all()
        self._refresh_badges()

    # ═════════════════════════════════════════════════════════════════════
    # Бейджі, підсумки, рядок стану
    # ═════════════════════════════════════════════════════════════════════
    @staticmethod
    def _counts(bridge: compat.TreeBridge) -> tuple[int, int]:
        rows = [values for values in bridge.rows() if not compat.is_placeholder_row(values)]
        marked = [values for values in rows if values and values[0] == compat.MARK_ON]
        return len(marked), len(rows)

    def _refresh_badges(self, *_args) -> None:
        if not getattr(self, "_ui_ready", False):
            return
        p2_marked, p2_total = self._counts(self.p2_orders_tree)
        orders_marked, orders_total = self._counts(self.orders_tree)
        copies_marked, copies_total = self._counts(self.copy_two_tree)
        extract_paths, from_copies = self._selected_order_paths()
        p2_paths = self._selected_p2_order_paths()

        for badge, marked, total in (
            (self._p2_queue_badge, p2_marked, p2_total),
            (self._orders_queue_badge, orders_marked, orders_total),
            (self._copies_queue_badge, copies_marked, copies_total),
        ):
            badge.setText(f"У черзі: {marked} з {total}")
            badge.setVisible(total > 0)

        for index, count in ((0, len(p2_paths)), (1, len(extract_paths)), (2, len(extract_paths))):
            self._tab_badges[index].setText(f"{count} у черзі")
            self._tab_badges[index].setVisible(count > 0)

        results = sum(1 for values in self.p2_tree.rows() if not compat.is_placeholder_row(values))
        self._p2_results_count.setText(f"Результатів: {results}")

        busy = self._busy_depth > 0
        for index, count in ((0, len(p2_paths)), (1, len(extract_paths)), (2, len(extract_paths))):
            self._strip_status[index].setText(self._strip_text(count, busy))

        current = self._tabs.currentIndex()
        c = theme.C
        if current == 0:
            marked, total = p2_marked, p2_total
            template = self.p2_back_page_path.get()
            folder = os.path.dirname(p2_paths[0]) if p2_paths else self.p2_orders_folder.get()
            mode = "Примірник № 2"
        else:
            marked, total = (copies_marked, copies_total) if from_copies else (orders_marked, orders_total)
            template = self.template_path.get() if current == 1 else self.message_content_template_path.get()
            folder = os.path.dirname(extract_paths[0]) if extract_paths else ""
            mode = "2 сторінки на 1 аркуш" if self.duplex_2up_layout.get() else "звичайний"
        self.status_strip.selected.setText(f"Вибрано: {marked}/{total}")
        template_name = os.path.basename(template) if template else "не обрано"
        self.status_strip.template.setText(
            f'Зразок: <b style="color:{c["text_soft"]}">{template_name}</b>'
        )
        self._ribbon_info.setText(
            f'Тека: <span style="color:{c["text_soft"]}">{_short_folder(folder) if folder else "—"}</span>'
            f'  <span style="color:{c["text_dim"]}">•</span>  '
            f'Друк: <b style="color:{c["teal_text"]}">{mode}</b>'
        )

    @staticmethod
    def _strip_text(count: int, busy: bool) -> str:
        c = theme.C
        if busy:
            state = f'<span style="color:{c["amber_text"]}">виконується…</span>'
        elif count:
            state = f'<span style="color:{c["teal_text"]}">готово до пуску</span>'
        else:
            state = f'<span style="color:{c["text_faint"]}">оберіть накази</span>'
        return (
            f'<span style="color:{c["text_muted"]}">Пакет:</span> '
            f'<b style="color:{c["text"]}">{plural_files(count)}</b> '
            f'<span style="color:{c["text_dim"]}">·</span> {state}'
        )

    # ═════════════════════════════════════════════════════════════════════
    # Прогони: блокування інтерфейсу на час пакета
    # ═════════════════════════════════════════════════════════════════════
    def _guarded(self, runner: Callable, *args):
        """Блокує кнопки на весь пакет. Лічильник, а не прапорець: повний цикл
        викликає вкладені прогони, і внутрішній не має розблоковувати зовнішній."""
        self._busy_depth += 1
        if self._busy_depth == 1:
            self._set_busy(True)
        try:
            return runner(*args)
        finally:
            self._busy_depth -= 1
            if self._busy_depth == 0:
                self._set_busy(False)

    def _set_busy(self, busy: bool) -> None:
        for widget in self._lockable_buttons:
            widget.set_busy_locked(busy)
        self.status_strip.set_busy(busy)
        if busy:
            QApplication.setOverrideCursor(Qt.CursorShape.BusyCursor)
        else:
            QApplication.restoreOverrideCursor()
        self._refresh_badges()
        compat.pump_events()

    def run_rozrahunok_action(self):
        return self._guarded(super().run_rozrahunok_action)

    def run_extracts_action(self, *args, **kwargs):
        parent_action = super().run_extracts_action
        return self._guarded(lambda: parent_action(*args, **kwargs))

    def run_management_extracts_action(self):
        return self._guarded(super().run_management_extracts_action)

    def run_full_cycle(self):
        return self._guarded(super().run_full_cycle)

    def run_generate_copies(self):
        return self._guarded(super().run_generate_copies)

    def run_generate_messages(self):
        return self._guarded(super().run_generate_messages)

    def run_order_review_action(self):
        return self._guarded(super().run_order_review_action)

    def run_print_action(self, pdf_check: bool = False):
        parent_action = super().run_print_action
        return self._guarded(lambda: parent_action(pdf_check))

    def run_print_test_sheet(self):
        return self._guarded(super().run_print_test_sheet)

    def generate_messages_for_dropped_order(self, path: str) -> None:
        """Наказ, перетягнутий у рамку сторінки повідомлень: обрати й одразу створити."""
        if self._busy_depth:
            self.log("Перетягування проігноровано: зачекайте завершення поточного пакета.")
            return
        self._set_orders([path])
        self.log(f"📥 [Drag-and-Drop] Наказ для повідомлень: {os.path.basename(path)} — створюю повідомлення.")
        self._refresh_badges()
        self.run_generate_messages()

    def accept_dropped_paths(self, paths: list[str]) -> None:
        if self._busy_depth:
            self.log("Перетягування проігноровано: зачекайте завершення поточного пакета.")
            return
        self.handle_drag_and_drop(paths)
        self._refresh_badges()

    # ═════════════════════════════════════════════════════════════════════
    # Журнал, превʼю, вікна
    # ═════════════════════════════════════════════════════════════════════
    def log(self, message: str):
        self.log_model.append(message)
        compat.pump_events()

    def log_p2(self, message: str):
        self.p2_log_model.append(message)
        compat.pump_events()

    def clear_log(self):
        self.log_model.clear()

    def copy_log(self, text_widget=None, redacted: bool = False):
        model = self.p2_log_model if text_widget is getattr(self, "p2_log_text", None) else self.log_model
        content = model.plain_text().strip()
        if not content:
            return
        if redacted:
            content = self.legacy.redact_sensitive_text(content)
        QGuiApplication.clipboard().setText(content)
        self.log(
            "🔒 Журнал скопійовано з автоматичним знеособленням. "
            "Перед пересиланням перевірте текст: нетипові назви й персональні дані можуть залишитися."
            if redacted
            else "📋 Текст журналу скопійовано в буфер обміну."
        )

    def _preview_pause(self, seconds: float) -> None:
        compat.sleep_responsive(seconds)

    def open_samples_window(self):
        # Раніше кнопка «Зразки» блокувалась на час пакета; пункт навігації — ні,
        # тож перевіряємо тут: зразки не міняють посеред генерації.
        if self._busy_depth:
            self.log("Зразки можна змінити після завершення поточного пакета.")
            return
        if self._samples_dialog is None:
            dialog = SamplesDialog(self, self.main_window)
            dialog.finished.connect(lambda *_args: setattr(self, "_samples_dialog", None))
            self._samples_dialog = dialog
        self._samples_dialog.show()
        theme.use_dark_title_bar(self._samples_dialog)
        self._samples_dialog.raise_()
        self._samples_dialog.activateWindow()

    def open_instruction_window(self):
        if self._instruction_dialog is None:
            dialog = InstructionDialog(self.main_window)
            dialog.finished.connect(lambda *_args: setattr(self, "_instruction_dialog", None))
            self._instruction_dialog = dialog
        self._instruction_dialog.show()
        theme.use_dark_title_bar(self._instruction_dialog)
        self._instruction_dialog.raise_()
        self._instruction_dialog.activateWindow()

    def _open_output_folder(self, folder: str, what: str) -> None:
        if folder and os.path.isdir(folder):
            os.startfile(folder)  # noqa: S606 - лише Windows, як і весь генератор
        else:
            self.legacy.messagebox.showwarning(
                "Папка результату",
                f"Папка {what} ще не створена. Вона з'явиться після першого запуску.",
            )

    def open_extracts_output_folder(self):
        self._open_output_folder(self.out_folder.get(), "витягів")

    def open_message_output_folder(self):
        folder = self.message_out_folder.get()
        if not folder and self.doc_path.get():
            folder = os.path.join(os.path.dirname(os.path.abspath(self.doc_path.get())), "Messages_Output")
        self._open_output_folder(folder, "повідомлень")

    def _open_compare(self, mode: str, generated_path: str = "") -> CompareWindow:
        window = CompareWindow(generated_path=generated_path, mode=mode, parent=self.main_window)
        self._compare_windows.append(window)
        window.destroyed.connect(
            lambda *_args, gone=window: self._compare_windows.remove(gone) if gone in self._compare_windows else None
        )
        window.show()
        theme.use_dark_title_bar(window)
        return window

    def open_compare_extracts(self):
        self._open_compare("витяги")

    def open_compare_copies(self):
        generated = self.p2_single_file.get() if self.p2_source_mode.get() == "file" else ""
        self._open_compare("примірник_2", generated)

    def open_compare_messages(self):
        self._open_compare("повідомлення_зміст")

    def on_window_close(self) -> None:
        self.save_config()
        for window in tuple(self._compare_windows):
            window.close()
        if self._samples_dialog is not None:
            self._samples_dialog.close()
        if self._instruction_dialog is not None:
            self._instruction_dialog.close()


def create_qt_app_class(legacy):
    """Клас застосунку: Qt-міксин поверх `App` саме цього модуля генератора."""
    return type("QtGeneratorApp", (QtShellMixin, legacy.App), {"legacy": legacy})


def launch(legacy, argv: list[str] | None = None) -> int:
    qt_app = QApplication.instance() or QApplication(list(argv if argv is not None else sys.argv))
    qt_app.setApplicationName("Node Automation Toolkit")
    theme.apply_theme(qt_app)
    install_qt_bridge(legacy)
    shell = create_qt_app_class(legacy)()
    # Тема з конфігу (`ui_theme`) відома лише після `load_config` у `App.__init__`.
    theme.apply_theme(qt_app, getattr(shell, "ui_theme", None) and shell.ui_theme.get())
    shell.main_window.show()
    return qt_app.exec()
