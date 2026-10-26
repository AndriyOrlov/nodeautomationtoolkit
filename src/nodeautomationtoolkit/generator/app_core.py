"""Основа вікна: змінні, конфіг, перетягування файлів, прогрес.

Винесено з `generate_extracts.py` без змін логіки (26.10.2026).
"""
from __future__ import annotations

from . import _paths
import os
import json
import time
import tkinter as tk
from tkinter import filedialog
from nodeautomationtoolkit import __version__ as APP_VERSION
from nodeautomationtoolkit.builtin_nodes.copy_generator import (
    PREVIEW_DEFAULT_DELAY,
)



class CoreMixin:
    """Основа вікна: змінні, конфіг, перетягування файлів, прогрес."""

    def __init__(self, root):
        self.root = root
        self.root.title(
            f"Node Automation Toolkit {APP_VERSION} — Генератор витягів та примірників"
        )
        self.root.geometry("1180x880")
        self.root.minsize(920, 700)

        # Конфігураційні змінні — Вкладка 1: Розрахунок та Витяги
        self.excel_path = tk.StringVar()
        self.doc_path = tk.StringVar()
        self.template_path = tk.StringVar()
        self.out_folder = tk.StringVar()
        self.executor = tk.StringVar()
        self.group_corps_var = tk.BooleanVar(value=True)
        self.duplex_2up_layout = tk.BooleanVar(value=True)
        # Повідомлення: не переносити пункти внутрішнього переміщення в управлінні.
        self.message_skip_internal_management = tk.BooleanVar(value=True)

        # Підписант оригіналу наказу (Командувач/Командир — зчитується автоматично з наказу)
        self.order_signer_position = tk.StringVar()
        self.order_signer_rank = tk.StringVar()
        self.order_signer_name = tk.StringVar()

        # Особа, яка засвідчує витяг («Згідно з оригіналом» / Засвідчувач)
        self.certifier_position = tk.StringVar(value="Т.в.о. начальника штабу – першого заступника командувача військ")
        self.certifier_rank = tk.StringVar(value="полковник")
        self.certifier_name = tk.StringVar()
        self.p2_certifier_position = tk.StringVar()
        self.p2_certifier_rank = tk.StringVar()
        self.p2_certifier_name = tk.StringVar()
        self.manual_order_paths: list[str] = []
        self.p2_manual_order_paths: list[str] = []
        self.source_summary = tk.StringVar(value="Джерело не обрано — оберіть наказ або кілька наказів")
        self.message_source_summary = tk.StringVar(value="Наказ для повідомлень не обрано")
        self.p2_source_summary = tk.StringVar(value="Накази не обрано")
        # Вікно зразків одне на програму: друге таке ж лише плодило б
        # розбіжні значення в тих самих полях.
        self._samples_window = None
        self._batch_running = False

        # Конфігураційні змінні — Вкладка 2: Примірники 2/3
        self.p2_source_mode = tk.StringVar(value="folder")
        self.p2_orders_folder = tk.StringVar()
        self.p2_single_file = tk.StringVar()
        self.p2_back_page_path = tk.StringVar()
        self.p2_out_folder = tk.StringVar()
        # Чи обрано папку результату вручну. Якщо ні — вона щоразу
        # переобчислюється від поточного джерела, інакше примірники
        # продовжували б писатись у папку попереднього наказу.
        self.p2_out_folder_manual = tk.BooleanVar(value=False)
        self.p2_copy_title = tk.StringVar(value="Примірник № 2")
        # Виконавець примірників. Порожнє поле = береться виконавець витягів.
        self.p2_executor = tk.StringVar()
        # Режим превʼю: Word видимий, після кожного кроку пауза, у журналі —
        # що саме робиться. Пауза рядком, бо це поле вводить користувач.
        self.p2_preview = tk.BooleanVar(value=False)
        self.p2_preview_delay = tk.StringVar(value=str(PREVIEW_DEFAULT_DELAY))
        self.last_copy_two_paths: list[str] = []

        # Конфігурація вкладки «Накази» — генератор наказів по особовому складу.
        # «new_order_» означає наказ, який ми складаємо; «order_signer_» вище —
        # підписант чужого наказу, з якого робимо витяги.
        self.new_order_archive_folder = tk.StringVar()
        self.new_order_index_folder = tk.StringVar()
        self.new_order_index_force = tk.BooleanVar(value=False)
        self.new_order_plan_path = tk.StringVar()
        self.new_order_document_paths: list[str] = []
        self.new_order_source_summary = tk.StringVar(
            value="Джерела не обрано — план переміщення або документи"
        )
        self.new_order_template_path = tk.StringVar()
        self.new_order_out_folder = tk.StringVar()
        self.new_order_executor = tk.StringVar()
        self.new_order_number = tk.StringVar()
        self.new_order_date = tk.StringVar()
        self.new_order_section = tk.StringVar()
        self.new_order_points = tk.StringVar(value="пункту ___")
        # Вид наказу: «призначення» (за замовчуванням) або «звільнення».
        self.new_order_action = tk.StringVar(value="призначення")
        self.new_order_law_points = tk.StringVar(value="пункту ___ частини ___ статті 26")
        # Особи, введені вручну. У конфіг НЕ пишуться: там ПІБ і РНОКПП.
        self.new_order_manual_people: list[dict[str, str]] = []
        self.new_order_kind = tk.StringVar(value="осіб офіцерського складу")
        self.new_order_unit = tk.StringVar()
        self.new_order_target_unit = tk.StringVar()
        self.new_order_bases = tk.StringVar()
        self.new_order_footer_bases = tk.StringVar()
        self.new_order_signer_position = tk.StringVar()
        self.new_order_signer_rank = tk.StringVar()
        self.new_order_signer_name = tk.StringVar()

        # Конфігурація вкладки «Повідомлення».
        self.message_cover_template_path = tk.StringVar()
        self.message_content_template_path = tk.StringVar()
        # Необов'язковий третій зразок — «Розрахунок розсилки» з {{кому_список}}.
        self.message_calc_template_path = tk.StringVar()
        # Тема вікна («dark» / «light»); перемикач — у навігації оболонки.
        self.ui_theme = tk.StringVar(value="dark")
        # Друк пакета: тека, витяги до управління за бажанням, як класти в ручну подачу.
        self.print_folder = tk.StringVar()
        self.print_management = tk.BooleanVar(value=False)
        self.print_feed_hint = tk.StringVar()
        self.message_out_folder = tk.StringVar()
        self.message_executor = tk.StringVar()

        # Смуга прогресу генерації (спільна для всіх вкладок)
        self.progress_value = tk.DoubleVar(value=0.0)
        self.progress_text = tk.StringVar(value="")
        self._progress_total = 0
        self._progress_done = 0
        self._progress_title = ""
        self._progress_started_at = None

        self.config_file = os.path.join(_paths.APP_DIR, "config.json")

        self.load_config()
        self.create_widgets()

    def create_widgets(self):
        """Інтерфейс будує Qt-оболонка (`QtShellMixin.create_widgets`)."""
        raise NotImplementedError("Запускайте генератор через generate_extracts_qt.py")

    def handle_drag_and_drop(self, raw_files):
        """Обробка перетягування файлів мишею (Drag-and-Drop) у вікно програми."""
        if not raw_files:
            return

        files = []
        for item in raw_files:
            if isinstance(item, bytes):
                try:
                    path_str = item.decode("utf-8")
                except UnicodeDecodeError:
                    path_str = item.decode("mbcs", errors="ignore")
            else:
                path_str = str(item)
            if os.path.exists(path_str):
                files.append(os.path.abspath(path_str))

        if not files:
            return

        # Визначаємо активну вкладку
        current_tab = self.notebook.index(self.notebook.select())
        handled_count = 0

        TAB_COPIES = 0
        TAB_EXTRACTS = 1
        TAB_MESSAGES = 2

        # Накази збираються в СПИСОК: перетягнути одразу кілька — звичайна
        # справа, а старий обробник лишав тільки останній файл.
        dropped_orders: list[str] = []

        for fpath in files:
            fname = os.path.basename(fpath).lower()
            if os.path.isdir(fpath):
                if current_tab == TAB_COPIES:
                    self.p2_orders_folder.set(fpath)
                    self.p2_source_mode.set("folder")
                    self._on_p2_source_mode_changed()
                    self.log_p2(f"📥 [Drag-and-Drop] Папку наказів встановлено: {fpath}")
                elif current_tab == TAB_MESSAGES:
                    orders_in_folder = self._orders_in_folder(fpath)
                    if orders_in_folder:
                        self._set_orders([orders_in_folder[0]])
                        self.log(
                            f"📥 [Drag-and-Drop] Наказ для повідомлень: {os.path.basename(orders_in_folder[0])}. "
                            "Повідомлення одинарні; з теки взято лише перший наказ."
                        )
                    else:
                        self.log("УВАГА: у перетягнутій теці немає наказів DOCX.")
                else:
                    orders_in_folder = self._orders_in_folder(fpath)
                    if orders_in_folder:
                        self._set_orders(orders_in_folder)
                        self.log(f"📥 [Drag-and-Drop] З теки взято наказів: {len(orders_in_folder)}")
                    else:
                        self.out_folder.set(fpath)
                        self.log(f"📥 [Drag-and-Drop] Папку результату встановлено: {fpath}")
                handled_count += 1
            elif fname.endswith((".xlsx", ".xls")):
                self.excel_path.set(fpath)
                self.log(f"📥 [Drag-and-Drop] Словник Excel встановлено: {os.path.basename(fpath)}")
                handled_count += 1
            elif fname.endswith(".docx") and not fname.startswith("~$"):
                if current_tab == TAB_COPIES and ("задн" in fname or "back" in fname or "шаблон" in fname):
                    self.p2_back_page_path.set(fpath)
                    self.log_p2(f"📥 [Drag-and-Drop] Заготовку примірника встановлено: {os.path.basename(fpath)}")
                elif current_tab == TAB_EXTRACTS and ("шаблон" in fname or "template" in fname or "зразок" in fname):
                    self.template_path.set(fpath)
                    self.log(f"📥 [Drag-and-Drop] Зразок витягу встановлено: {os.path.basename(fpath)}")
                elif current_tab == TAB_MESSAGES and ("зміст" in fname or "content" in fname):
                    self.message_content_template_path.set(fpath)
                    self.log(f"📥 [Drag-and-Drop] Зразок змісту повідомлення встановлено: {os.path.basename(fpath)}")
                elif current_tab == TAB_MESSAGES and ("титул" in fname or "cover" in fname or "шаблон" in fname):
                    self.message_cover_template_path.set(fpath)
                    self.log(f"📥 [Drag-and-Drop] Зразок супроводу встановлено: {os.path.basename(fpath)}")
                else:
                    dropped_orders.append(fpath)
                handled_count += 1

        if dropped_orders:
            if current_tab == TAB_COPIES:
                self._set_p2_orders(dropped_orders)
                self.log_p2(f"📥 [Drag-and-Drop] Наказів для примірників: {len(dropped_orders)}")
            elif current_tab == TAB_MESSAGES:
                self._set_orders([dropped_orders[0]])
                self.log(
                    f"📥 [Drag-and-Drop] Наказ для повідомлень: {os.path.basename(dropped_orders[0])}"
                    + (f". Решту файлів ({len(dropped_orders) - 1}) не взято: повідомлення одинарні."
                       if len(dropped_orders) > 1 else "")
                )
            else:
                self._set_orders(dropped_orders)
                self.log(f"📥 [Drag-and-Drop] Наказів до обробки: {len(dropped_orders)}")

        if handled_count > 0:
            self.save_config()

    def load_config(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.excel_path.set(data.get("excel_path", ""))
                    self.doc_path.set(data.get("doc_path", ""))
                    self.template_path.set(data.get("template_path", ""))
                    self.out_folder.set(data.get("out_folder", ""))
                    self.executor.set(data.get("executor", ""))
                    self.certifier_position.set(
                        data.get("certifier_position", "Т.в.о. начальника штабу – першого заступника командувача військ")
                    )
                    self.certifier_rank.set(data.get("certifier_rank", "полковник"))
                    self.certifier_name.set(data.get("certifier_name", ""))
                    # Разова міграція старих спільних полів; далі незалежні.
                    for field in ("position", "rank", "name"):
                        getattr(self, f"p2_certifier_{field}").set(
                            data.get(f"p2_certifier_{field}", getattr(self, f"certifier_{field}").get())
                        )
                    if "group_corps" in data:
                        self.group_corps_var.set(data["group_corps"])
                    if "duplex_2up_layout" in data:
                        self.duplex_2up_layout.set(data["duplex_2up_layout"])
                    if "message_skip_internal_management" in data:
                        self.message_skip_internal_management.set(
                            data["message_skip_internal_management"]
                        )

                    self.p2_orders_folder.set(data.get("p2_orders_folder", ""))
                    self.p2_single_file.set(data.get("p2_single_file", ""))
                    self.manual_order_paths = [
                        path for path in data.get("manual_order_paths", []) if os.path.isfile(path)
                    ]
                    self.p2_manual_order_paths = [
                        path for path in data.get("p2_manual_order_paths", []) if os.path.isfile(path)
                    ]
                    # Режим джерела примірників раніше не зберігався: після
                    # перезапуску програма поверталась до «теки» й мовчки
                    # ігнорувала обрані файли. Для старого конфігу режим
                    # виводимо з того, що в ньому є.
                    saved_mode = data.get("p2_source_mode")
                    if saved_mode in ("folder", "file"):
                        self.p2_source_mode.set(saved_mode)
                    elif self.p2_manual_order_paths or os.path.isfile(self.p2_single_file.get()):
                        self.p2_source_mode.set("file")
                    self.p2_back_page_path.set(data.get("p2_back_page_path", ""))
                    self.p2_out_folder.set(data.get("p2_out_folder", ""))
                    self.p2_out_folder_manual.set(bool(data.get("p2_out_folder_manual", False)))
                    # Правила для примірника № 3 ще не погоджені.
                    self.p2_copy_title.set("Примірник № 2")
                    self.p2_executor.set(data.get("p2_executor", ""))
                    self.p2_preview.set(bool(data.get("p2_preview", False)))
                    self.p2_preview_delay.set(
                        str(data.get("p2_preview_delay", PREVIEW_DEFAULT_DELAY))
                    )

                    for field in (
                        "archive_folder", "index_folder", "plan_path", "template_path",
                        "out_folder", "executor", "number", "date", "section", "unit",
                        "target_unit", "bases", "footer_bases",
                        "signer_position", "signer_rank", "signer_name", "law_points",
                    ):
                        value = data.get(f"new_order_{field}")
                        if value:
                            getattr(self, f"new_order_{field}").set(value)
                    for field, default in (
                        ("points", "пункту ___"),
                        ("kind", "осіб офіцерського складу"),
                        ("action", "призначення"),
                    ):
                        getattr(self, f"new_order_{field}").set(data.get(f"new_order_{field}") or default)
                    self.new_order_document_paths = [
                        path for path in data.get("new_order_document_paths", []) if os.path.isfile(path)
                    ]

                    self.message_cover_template_path.set(data.get("message_cover_template_path", ""))
                    self.message_content_template_path.set(data.get("message_content_template_path", ""))
                    self.message_calc_template_path.set(data.get("message_calc_template_path", ""))
                    self.ui_theme.set(data.get("ui_theme") or "dark")
                    self.print_folder.set(data.get("print_folder", ""))
                    self.print_management.set(bool(data.get("print_management", False)))
                    self.print_feed_hint.set(data.get("print_feed_hint", ""))
                    self.message_out_folder.set(data.get("message_out_folder", ""))
                    self.message_executor.set(data.get("message_executor", ""))

            except Exception:
                pass

    def save_config(self):
        data = {
            "excel_path": self.excel_path.get(),
            "doc_path": self.doc_path.get(),
            "template_path": self.template_path.get(),
            "out_folder": self.out_folder.get(),
            "executor": self.executor.get(),
            "certifier_position": self.certifier_position.get(),
            "certifier_rank": self.certifier_rank.get(),
            "certifier_name": self.certifier_name.get(),
            "p2_certifier_position": self.p2_certifier_position.get(),
            "p2_certifier_rank": self.p2_certifier_rank.get(),
            "p2_certifier_name": self.p2_certifier_name.get(),
            "group_corps": self.group_corps_var.get(),
            "duplex_2up_layout": self.duplex_2up_layout.get(),
            "message_skip_internal_management": self.message_skip_internal_management.get(),
            "p2_orders_folder": self.p2_orders_folder.get(),
            "p2_single_file": self.p2_single_file.get(),
            "p2_source_mode": self.p2_source_mode.get(),
            "manual_order_paths": self.manual_order_paths,
            "p2_manual_order_paths": self.p2_manual_order_paths,
            "p2_back_page_path": self.p2_back_page_path.get(),
            "p2_out_folder": self.p2_out_folder.get(),
            "p2_out_folder_manual": self.p2_out_folder_manual.get(),
            "p2_copy_title": self.p2_copy_title.get(),
            "p2_executor": self.p2_executor.get(),
            "p2_preview": self.p2_preview.get(),
            "p2_preview_delay": self.p2_preview_delay.get(),
            **{
                f"new_order_{field}": getattr(self, f"new_order_{field}").get()
                for field in (
                    "archive_folder", "index_folder", "plan_path", "template_path",
                    "out_folder", "executor", "number", "date", "section", "points",
                    "kind", "unit", "target_unit", "bases", "footer_bases",
                    "signer_position", "signer_rank", "signer_name",
                    "action", "law_points",
                )
            },
            "new_order_document_paths": list(self.new_order_document_paths),
            "message_cover_template_path": self.message_cover_template_path.get(),
            "message_content_template_path": self.message_content_template_path.get(),
            "message_calc_template_path": self.message_calc_template_path.get(),
            "ui_theme": self.ui_theme.get(),
            "print_folder": self.print_folder.get(),
            "print_management": self.print_management.get(),
            "print_feed_hint": self.print_feed_hint.get(),
            "message_out_folder": self.message_out_folder.get(),
            "message_executor": self.message_executor.get(),
        }
        try:
            with open(self.config_file, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
        except Exception:
            pass

    # -------------------------------------------------------------------------
    # СМУГА ПРОГРЕСУ ГЕНЕРАЦІЇ
    # -------------------------------------------------------------------------
    def progress_begin(self, total: int, title: str) -> None:
        """Починає показ прогресу: `total` кроків під назвою `title`."""
        self._progress_total = max(1, int(total))
        self._progress_done = 0
        self._progress_title = title
        self._progress_started_at = time.monotonic()
        self._progress_render(f"{title}: 0 / {self._progress_total}", 0.0)

    def progress_step(self, done: int, detail: str = "") -> None:
        """Показує, що виконано `done` кроків із загальної кількості."""
        total = getattr(self, "_progress_total", 0)
        if not total:
            return
        self._progress_done = max(0, min(int(done), total))
        title = getattr(self, "_progress_title", "Обробка")
        caption = f"{title}: {self._progress_done} / {total}"
        if detail:
            caption += f" — {detail}"
        self._progress_render(caption, 100.0 * self._progress_done / total)

    def progress_end(self, detail: str = "") -> None:
        """Завершує показ прогресу.

        З підсумком (`detail`) смуга лишається заповненою — це успішне
        завершення. Без підсумку виклик означає, що робота урвалась (виняток,
        закритий діалог), тож смуга застигає там, де була, з поміткою.
        """
        total = getattr(self, "_progress_total", 0)
        if not total:
            return  # прогрес уже завершено успішним викликом
        started_at = getattr(self, "_progress_started_at", None)
        done = getattr(self, "_progress_done", 0)
        caption = detail or "Перервано"
        if started_at is not None:
            caption += f" ({time.monotonic() - started_at:.0f} с)"
        percent = 100.0 if detail else (100.0 * done / total)
        self._progress_total = 0
        self._progress_render(caption, percent)

    def _progress_render(self, caption: str, percent: float) -> None:
        """Малює стан смуги. Мовчить, якщо інтерфейс не побудовано (тести, e2e)."""
        try:
            self.progress_text.set(caption)
            self.progress_value.set(percent)
            self.root.update_idletasks()
        except Exception:
            pass

    # ── Вибір файлів ─────────────────────────────────────────────────────────
    def select_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.excel_path.set(path)
            self.save_config()
