"""Розрахунок розсилки та витяги.

Винесено з `generate_extracts.py` без змін логіки (26.10.2026).
"""
from __future__ import annotations

import os
import sys
import re
import shutil
import time
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.constants import DISABLED, NORMAL, W
import win32com.client
from nodeautomationtoolkit.builtin_nodes.recipient_mapping import (
    read_recipient_mapping,
    map_military_units,
    _format_item_numbers_range,
)
from nodeautomationtoolkit.builtin_nodes.blank_images import remove_blank_images
from nodeautomationtoolkit.builtin_nodes.template_tags import (
    SIGNER_TAGS, tag_aliases, certifier_tags, signer_tags,
)

from .errors import UserError, explain_error
from .naming import _slash_to_lines, build_extracts_filename, extract_metadata_from_filename, format_ukr_date, sanitize_filename
from .order_text import order_extracts_for_2up
from .reports import MANAGEMENT_RESULT_LABEL, _save_table_to_excel, control_report_path, management_calculation_rows, select_extracts_for_scope
from .word_doc import DEATH_ITEM_FONT_SIZES, EXTRA_BLANK_SPACINGS, EXTRA_TEXT_SPACINGS, WD_UNDEFINED, copy_template_for_editing, force_quit_word, is_death_item_text, is_page_break_only_text, is_path_writable, iter_paragraphs, last_paragraph, paragraph_line_count, remove_page_breaks_in_range, restore_word_settings, tune_word_for_batch


class ExtractsMixin:
    """Розрахунок розсилки та витяги."""

    def show_template_tags(self):
        messagebox.showinfo(
            "Теги шаблону витягу",
            "Додайте ці теги до DOCX-шаблону витягу:\n\n"
            "• {{кому}} — значення «Кому» з Excel; якщо порожнє — шифр.\n"
            "• {{куди}} — значення «Куди» з Excel.\n"
            "• {{номер_наказу}} або {{номер}} — номер лише з назви файла наказу.\n"
            "• {{дата_наказу}} або {{дата}} — дата лише з назви файла наказу.\n"
            "• {{пункти}} — перелік пунктів цього витягу.\n"
            "• {{зміст}} — повний текст пунктів з оригіналу наказу.\n\n"
            "Підписант оригіналу наказу (зчитується автоматично після пунктів):\n"
            "• {{підписант_посада}} (або {{посада_підписанта}})\n"
            "• {{підписант_звання}} (або {{звання_підписанта}})\n"
            "• {{підписант_піб}} (або {{підписант_імя}}, {{прізвище_підписанта}})\n\n"
            "• {{підписант}} — увесь блок: посада, а нижче звання та ім’я.\n\n"
            "Засвідчення («Згідно з оригіналом»):\n"
            "• {{згідно_з_оригіналом}} або {{засвідчення}} → «Згідно з оригіналом»\n"
            "• {{засвідчувач_посада}} (або {{затверджувач_посада}}, {{згідно_з_оригіналом_посада}})\n"
            "• {{засвідчувач_звання}} (або {{затверджувач_звання}}, {{згідно_з_оригіналом_звання}})\n"
            "• {{засвідчувач_піб}} (або {{затверджувач_піб}}, {{згідно_з_оригіналом_піб}}, {{засвідчувач}})\n\n"
            "Назви в дужках рівноцінні: підставляється те саме значення.\n"
            "Засвідчувач витягів і засвідчувач примірників — окремі поля у вікні «Зразки».\n\n"
            "Витяги до управління:\n"
            "• той самий зразок, але {{кому}} та {{куди}} лишаються порожніми;\n"
            "• складаються в окремий файл і не входять у розрахунок розсилки.\n\n"
            "Службові реквізити:\n"
            "• {{виконавець}} — виконавець з форми програми.",
        )

    def _populate_result_tab(self, key: str, columns: list[str], rows: list[tuple]):
        tree = self.result_views[key]
        for item in tree.get_children():
            tree.delete(item)
        tree["columns"] = columns
        for position, title in enumerate(columns):
            values = [str(row[position]) if position < len(row) else "" for row in rows]
            width = max(110, min(420, max([len(title), *(len(value) for value in values)] or [len(title)]) * 8))
            tree.heading(title, text=title)
            tree.column(title, width=width, minwidth=90, stretch=True, anchor=W)
        if key == "calculation" and hasattr(tree, "tag_configure"):
            tree.tag_configure(
                "management",
                background="#d1fae5",
                foreground="#065f46",
                font=("Segoe UI", 9, "bold"),
            )
        for row in rows:
            values = tuple("-" if value in (None, "") else str(value) for value in row)
            tags = (
                ("management",)
                if key == "calculation" and values and values[0] == MANAGEMENT_RESULT_LABEL
                else ()
            )
            tree.insert("", tk.END, values=values, tags=tags)

    def show_analysis_results(self, map_result: dict):
        units_table = map_result.get("units_table")
        calculation_rows = []
        if units_table and hasattr(units_table, "rows"):
            calculation_rows = [(row[0], row[2], row[3]) for row in units_table.rows]
        calculation_rows.extend(management_calculation_rows(map_result))
        self._populate_result_tab(
            "calculation",
            ["Військова частина / Відправник", "Пункти витягу", "Кількість"],
            calculation_rows,
        )

        unmatched_rows = [
            (item.get("label", "-"), item.get("text", ""), item.get("reason", ""))
            for item in map_result.get("unmatched_items", [])
        ]
        self._populate_result_tab("unmatched", ["Пункт", "Текст пункту", "Причина"], unmatched_rows)

        routing_rows = [
            (
                item.get("label", "-"),
                item.get("matched_entries", "-"),
                item.get("applied_rules", "-"),
                item.get("final_recipients", "-"),
            )
            for item in map_result.get("routing_audit", [])
        ]
        self._populate_result_tab(
            "routing",
            ["Пункт", "Збіги з таблиці", "Застосовані правила", "Підсумкові адресати"],
            routing_rows,
        )
        self._populate_result_tab("layout", ["Стан макета"], [])

    def show_layout_warnings(self, warnings: list[str]):
        rows = [(warning,) for warning in warnings]
        if not rows:
            rows = [("Макет сформовано без попереджень.",)]
        self._populate_result_tab("layout", ["Стан макета"], rows)
        self.results_notebook.select(self.result_tabs["layout" if warnings else "calculation"])

    def select_template(self):
        path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx *.doc")])
        if path:
            self.template_path.set(path)
            self.save_config()

    def edit_template(self):
        path = self.template_path.get()
        if not path or not os.path.exists(path):
            messagebox.showwarning("Помилка", "Спочатку виберіть існуючий файл шаблону!")
            return
        try:
            self.log("Відкриваємо шаблон витягу у Word для редагування...")
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = True
            try:
                doc = word.Documents.Open(os.path.abspath(path))
            except Exception:
                # Інакше лишалося порожнє вікно Word без документа.
                force_quit_word(word)
                raise
            doc.Activate()
        except Exception as e:
            self.log(f"Помилка відкриття шаблону: {str(e)}")

    def select_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.out_folder.set(path)
            self.save_config()

    def _set_extract_action_buttons_state(self, state):
        """Разом блокує дії вкладки, щоб два пакети не стартували паралельно."""
        for name in ("btn_calc", "btn_extracts", "btn_management_extracts"):
            button = getattr(self, name, None)
            if button is not None:
                button.config(state=state)

    def run_rozrahunok_action(self):
        self.save_config()
        order_paths, from_copies = self._selected_order_paths()
        if not self.excel_path.get() or not order_paths:
            messagebox.showwarning(
                "Помилка",
                "Виберіть словник Excel і хоча б один наказ або примірник № 2.",
            )
            return

        missing = self._missing_selected_files(("Словник Excel", self.excel_path.get()))
        if missing:
            self._report_missing_files(missing)
            return

        self._set_extract_action_buttons_state(DISABLED)
        try:
            # Кожен наказ обробляється окремо: збій одного не має зривати
            # решту пакета.
            failures = self._run_batch(
                order_paths, self.run_rozrahunok, "Розрахунок для", alongside=from_copies
            )
            self.save_config()
            self._report_batch_result(
                len(order_paths), failures, "Розрахунок розсилки", "розрахунок розсилки"
            )
        finally:
            self._set_extract_action_buttons_state(NORMAL)

    def _run_extracts_scope_action(
        self,
        scope: str,
        stage_label: str,
        result_title: str,
        result_action: str,
    ):
        self.save_config()
        order_paths, from_copies = self._selected_order_paths()
        if not self.excel_path.get() or not self.template_path.get() or not order_paths:
            messagebox.showwarning(
                "Помилка",
                "Виберіть словник Excel, шаблон витягу і хоча б один наказ або примірник № 2.",
            )
            return

        missing = self._missing_selected_files(
            ("Словник Excel", self.excel_path.get()),
            ("Зразок витягу", self.template_path.get()),
        )
        if missing:
            self._report_missing_files(missing)
            return

        self._set_extract_action_buttons_state(DISABLED)
        try:
            failures = self._run_batch(
                order_paths,
                lambda: self.run_extracts(scope=scope),
                stage_label,
                alongside=from_copies,
            )
            self.save_config()
            self._report_batch_result(
                len(order_paths), failures, result_title, result_action
            )
        finally:
            self._set_extract_action_buttons_state(NORMAL)

    def run_extracts_action(self, scope: str = "all"):
        """Створює адресний файл і окремий файл витягів до управління."""
        return self._run_extracts_scope_action(
            "all", "Усі витяги для", "Генерація всіх витягів", "генерацію всіх витягів"
        )

    def run_management_extracts_action(self):
        """Створює лише окремі витяги до управління без адресування й 2-на-1."""
        return self._run_extracts_scope_action(
            "management",
            "Витяги до управління для",
            "Витяги до управління",
            "генерацію витягів до управління",
        )

    def run_full_cycle(self):
        """Повний прохід по обраних наказах: примірники № 2 → розрахунок → витяги.

        Джерелом є список наказів головного вікна. Примірники формуються
        першими, і далі розрахунок та витяги йдуть уже по них — саме так, як
        це робиться вручну послідовними діями вкладки.
        """
        if self._batch_running:
            self.log("Пакет уже виконується — дочекайтеся завершення.")
            return

        order_paths, from_copies = self._selected_order_paths()
        if from_copies:
            # Примірники вже є: другий раз їх робити нема з чого, тож повний
            # цикл починався б із самого себе. Беремо вихідні накази.
            order_paths = [path for path in self.manual_order_paths if os.path.isfile(path)]
        if not order_paths:
            messagebox.showwarning("Помилка", "Спочатку оберіть наказ або кілька наказів.")
            return
        if not self.p2_back_page_path.get() or not self.excel_path.get() or not self.template_path.get():
            messagebox.showwarning(
                "Не вистачає зразків",
                "Для повного циклу потрібні: заготовка примірника, словник Excel і зразок витягу.\n\n"
                "Заповніть їх у вікні «Зразки та реквізити».",
            )
            return

        missing = self._missing_selected_files(
            ("Словник Excel", self.excel_path.get()),
            ("Зразок витягу", self.template_path.get()),
            ("Заготовка примірника (остання сторінка)", self.p2_back_page_path.get()),
        )
        if missing:
            self._report_missing_files(missing)
            return

        self._batch_running = True
        self.btn_full_cycle.config(state=DISABLED)
        try:
            self.log(
                f"\n⚙️ ПОВНИЙ ЦИКЛ для {len(order_paths)} наказ(ів): "
                "примірники → розрахунок → адресні витяги + витяги до управління"
            )
            # 1. Примірники № 2 з тих самих наказів.
            self._set_p2_orders(order_paths)
            self.run_generate_copies()

            # 2 і 3. Далі працюємо з примірниками, якщо вони створились:
            # _set_copy_two_sources уже наповнив список і позначив їх.
            if not self.last_copy_two_paths:
                self.log("Примірники не створено — розрахунок і витяги виконуємо за самими наказами.")
            self.run_rozrahunok_action()
            self.run_extracts_action(scope="all")
            self.log("⚙️ Повний цикл завершено.")
        finally:
            self._batch_running = False
            self.btn_full_cycle.config(state=NORMAL)

    def _run_batch(
        self, order_paths: list[str], handler, stage_label: str, alongside: bool = False
    ) -> list[tuple[str, str]]:
        """Обробляє накази по черзі, не перериваючи пакет через збій одного.

        Повертає перелік `(назва файлу, текст помилки)` для тих, що не вдалися.
        """
        import traceback

        failures: list[tuple[str, str]] = []
        for index, order_path in enumerate(order_paths, start=1):
            name = os.path.basename(order_path)
            self.log(f"\n[{index}/{len(order_paths)}] {stage_label}: {name}")
            try:
                self._set_processing_order(order_path, alongside=alongside)
                handler()
            except Exception as error:
                traceback.print_exc()
                explanation = explain_error(error)
                self.log(f"  ПОМИЛКА під час обробки «{name}»:\n    "
                         + explanation.replace("\n", "\n    "))
                failures.append((name, explanation))
        return failures

    def _report_batch_result(
        self, total: int, failures: list[tuple[str, str]], title: str, action_name: str
    ) -> None:
        """Показує підсумок пакета: скільки вдалося, а що саме — ні."""
        succeeded = total - len(failures)
        if not failures:
            messagebox.showinfo("Успіх", f"{title}: успішно оброблено {succeeded} з {total} файл(ів).")
            return

        details = "\n\n".join(
            f"Наказ «{name}» — не оброблено.\n{error}" for name, error in failures[:10]
        )
        if len(failures) > 10:
            details += f"\n\n… і ще {len(failures) - 10} файл(ів) з такими самими збоями."
        self.log(f"\nНе вдалося виконати {action_name} для {len(failures)} файл(ів).")
        messagebox.showwarning(
            title,
            f"Готово: {succeeded} з {total} файл(ів). Не вийшло: {len(failures)}.\n\n"
            "Нижче названо наказ, на якому програма зупинилась, і причину. "
            "Назва наказу — це те, що оброблялося, а не обов'язково те, чого бракує.\n\n"
            f"{details}",
        )

    def _log_routing_module(self):
        """Пише в журнал, який саме модуль маршрутизації завантажено."""
        mapping_module = sys.modules.get(map_military_units.__module__)
        module_file = getattr(mapping_module, "__file__", "")
        module_version = getattr(mapping_module, "ROUTING_VERSION", "без версії")
        self.log(
            "Модуль маршрутизації: "
            + (os.path.abspath(str(module_file)) if module_file else "невідомо")
            + f" · версія {module_version}"
        )

    def _log_mapping_source(self, excel_res: dict):
        """Пише в журнал реквізити прочитаного Excel-еталона.

        Правило: вибраний Excel — єдиний еталон маршрутизації, і кожен запуск
        читає його з диска без кешу. Журнал має це підтверджувати, щоб було
        видно, що розрахунок зроблено саме за поточним файлом.
        """
        modified_ns = excel_res.get("source_modified_ns")
        if modified_ns:
            source_modified = datetime.fromtimestamp(
                modified_ns / 1_000_000_000
            ).strftime("%d.%m.%Y %H:%M:%S")
        else:
            source_modified = "невідомо"
        self.log(
            f"Еталон Excel: {excel_res.get('source_path', self.excel_path.get())} · "
            f"змінено {source_modified} · {excel_res.get('source_size', 0)} байт. "
            "Файл перечитано з диска без кешу."
        )

    def run_rozrahunok(self):
        self.log("\n=== РОЗРАХУНОК РОЗСИЛКИ ===")
        self._log_routing_module()
        self.log(f"Читаємо словник: {self.excel_path.get()}")
        excel_res = read_recipient_mapping(path=self.excel_path.get())
        self._log_mapping_source(excel_res)
        mapping_base = excel_res.get("mapping", {})
        unique_entries = {id(value): value for value in mapping_base.values() if isinstance(value, dict)}
        abbreviation_count = sum(bool(str(value.get("abbreviation", "")).strip()) for value in unique_entries.values())

        self.log(f"Читаємо наказ: {self.doc_path.get()}")
        source_text = self._read_word_text(self.doc_path.get())
        text, _ = self._refresh_order_signer(source_text)

        self.log(
            f"Зчитано {len(unique_entries)} записів з Excel. "
            f"Скорочень з колонки C: {abbreviation_count}."
        )
        os.makedirs(self.out_folder.get(), exist_ok=True)
        import copy

        mapping_extracts = copy.deepcopy(mapping_base)

        if not self.group_corps_var.get():
            self.log("Режим: БЕЗ угруповання по Корпусах (всі частини окремо)")
            for k, v in mapping_extracts.items():
                if isinstance(v, dict) and "corps" in v:
                    v["corps"] = ""
        else:
            self.log("Режим: З угрупованням по Корпусах (Варіант 1)")

        self.log("Аналізуємо текст наказу (структура, ТЦК, ВЧ)...")
        map_res_extracts = map_military_units(
            text=text, mapping=mapping_extracts, bio_lines=self._word_bio_lines(self.doc_path.get())
        )
        self.show_analysis_results(map_res_extracts)

        for invalid_link in map_res_extracts.get("invalid_corps_links", []):
            self.log(
                "УВАГА: корпус не знайдено окремим рядком у таблиці; "
                f"частину залишено самостійним адресатом: {invalid_link.get('unit', '')} → {invalid_link.get('corps', '')}"
            )
        for tck_reference in map_res_extracts.get("unresolved_tck_references", []):
            self.log(f"УВАГА: ТЦК не визначено або його ОТЦК відсутній у таблиці; пункт не додано до розсилки: {tck_reference}")

        order_base = sanitize_filename(os.path.splitext(os.path.basename(self.doc_path.get()))[0])
        unmatched_items = map_res_extracts.get("unmatched_items", [])
        unmatched_report = control_report_path(self.out_folder.get(), "Контроль_пропущених_пунктів", order_base)
        _save_table_to_excel(
            unmatched_report,
            ["Пункт", "Текст пункту", "Причина"],
            [
                (item.get("label", ""), item.get("text", ""), item.get("reason", ""))
                for item in unmatched_items
            ],
        )
        if unmatched_items:
            self.log(f"УВАГА: {len(unmatched_items)} пункт(ів) без адресата. Контрольний файл: {unmatched_report}")
        else:
            self.log(f"Контрольний файл створено: {unmatched_report} (пропущених пунктів немає)")

        routing_report = control_report_path(self.out_folder.get(), "Контроль_маршрутизації", order_base)
        routing_data = [
            (
                item.get("label", ""),
                item.get("matched_entries", ""),
                item.get("applied_rules", ""),
                item.get("item_recipients", ""),
                item.get("context_recipients", ""),
                item.get("final_recipients", ""),
            )
            for item in map_res_extracts.get("routing_audit", [])
        ]
        _save_table_to_excel(
            routing_report,
            ["Пункт", "Збіги з таблиці", "Застосовані правила", "Адресати з пункту", "Адресати з контексту", "Підсумкові адресати"],
            routing_data,
        )
        self.log(f"Контроль маршрутизації: {routing_report}")

        table_extracts = map_res_extracts.get("units_table")
        if table_extracts and hasattr(table_extracts, "rows"):
            filtered_rows_ext = [(row[0], row[2], row[3]) for row in table_extracts.rows]
            out_file_ext = os.path.join(self.out_folder.get(), f"Розрахунок_розсилки_{order_base}.xlsx")
            _save_table_to_excel(
                out_file_ext,
                ["Військова частина / Відправник", "Номери пунктів витягу", "Кількість пунктів"],
                filtered_rows_ext,
            )
            self.log(f"Збережено розрахунок: {out_file_ext}")
        else:
            self.log("Помилка: NAT не повернув таблицю.")

        self.log("Готово! Завершено розрахунок розсилки.")

    def run_extracts(self, scope: str = "all"):
        scope_titles = {
            "general": "ГЕНЕРАЦІЯ АДРЕСНИХ ВИТЯГІВ",
            "management": "ГЕНЕРАЦІЯ ВИТЯГІВ ДО УПРАВЛІННЯ",
            "all": "ГЕНЕРАЦІЯ ВСІХ ВИТЯГІВ",
        }
        if scope not in scope_titles:
            raise ValueError(f"Невідомий режим генерації витягів: {scope}")
        self.log(f"\n=== {scope_titles[scope]} ===")
        self._log_routing_module()
        excel_res = read_recipient_mapping(path=self.excel_path.get())
        self._log_mapping_source(excel_res)
        mapping = excel_res.get("mapping", {})
        unique_entries = {id(value): value for value in mapping.values() if isinstance(value, dict)}
        abbreviation_count = sum(bool(str(value.get("abbreviation", "")).strip()) for value in unique_entries.values())
        self.log(
            f"Зчитано {len(unique_entries)} записів з Excel. "
            f"Скорочень з колонки C: {abbreviation_count}."
        )

        source_text = self._read_word_text(self.doc_path.get())
        text, _detected_order_signer = self._refresh_order_signer(source_text)
        # Реквізити підписанта беремо з ПОЛІВ інтерфейсу, а не з розпізнаного
        # напряму: `_refresh_order_signer` щойно заповнив їх із самого наказу,
        # але користувач має право виправити розпізнане — так само, як для
        # засвідчувача. (Тимчасове правило «підписанта у витягах не переносимо»
        # скасоване користувачем 04.09.2026.)
        order_signer = {
            "position": self.order_signer_position.get().strip(),
            "rank": self.order_signer_rank.get().strip(),
            "name": self.order_signer_name.get().strip(),
        }
        if any(order_signer.values()):
            signer_summary = " ".join(
                part for part in (order_signer["rank"], order_signer["name"]) if part
            )
            self.log(
                "Підписант оригіналу наказу переноситься у витяг"
                + (f": {signer_summary}." if signer_summary else ".")
            )
        else:
            self.log(
                "УВАГА: реквізитів підписанта наказу немає — теги {{підписант_…}} "
                "буде прибрано з витягу разом із їхніми рядками."
            )
        filename = os.path.basename(self.doc_path.get())
        order_num, order_date = extract_metadata_from_filename(filename)
        if not order_num:
            self.log("УВАГА: Номер наказу не знайдено в назві файлу. Тег {{номер_наказу}} залишиться для ручного заповнення.")
        if not order_date:
            self.log("УВАГА: Дату наказу не знайдено в назві файлу. Тег {{дата_наказу}} залишиться для ручного заповнення.")

        order_date_formatted = format_ukr_date(order_date)

        if not self.group_corps_var.get():
            self.log("УВАГА: Вибрано режим без угруповання по корпусах для витягів!")
            for k, v in mapping.items():
                if isinstance(v, dict) and "corps" in v:
                    v["corps"] = ""

        self.log("Аналізуємо структуру наказу (блоки, адресати)...")
        map_res = map_military_units(
            text=text, mapping=mapping, bio_lines=self._word_bio_lines(self.doc_path.get())
        )
        self.show_analysis_results(map_res)
        preamble_recipient = str(map_res.get("preamble_recipient") or "").strip()
        if preamble_recipient:
            self.log(f"Адресат у преамбулі знайдено за колонкою A: {preamble_recipient}")
        else:
            self.log("Адресата у преамбулі за колонкою A не знайдено.")

        for invalid_link in map_res.get("invalid_corps_links", []):
            self.log(
                "УВАГА: корпус не знайдено окремим рядком у таблиці; "
                f"частину залишено самостійним адресатом: {invalid_link.get('unit', '')} → {invalid_link.get('corps', '')}"
            )
        for tck_reference in map_res.get("unresolved_tck_references", []):
            self.log(f"УВАГА: ТЦК не визначено або його ОТЦК відсутній у таблиці; витяг не буде створено: {tck_reference}")

        order_base = sanitize_filename(os.path.splitext(os.path.basename(self.doc_path.get()))[0])
        routing_report = control_report_path(self.out_folder.get(), "Контроль_маршрутизації", order_base)
        routing_data = [
            (
                item.get("label", ""),
                item.get("matched_entries", ""),
                item.get("applied_rules", ""),
                item.get("item_recipients", ""),
                item.get("context_recipients", ""),
                item.get("final_recipients", ""),
            )
            for item in map_res.get("routing_audit", [])
        ]
        _save_table_to_excel(
            routing_report,
            ["Пункт", "Збіги з таблиці", "Застосовані правила", "Адресати з пункту", "Адресати з контексту", "Підсумкові адресати"],
            routing_data,
        )

        unmatched_items = map_res.get("unmatched_items", [])
        missing_report = control_report_path(self.out_folder.get(), "Контроль_пропущених_пунктів", order_base)
        _save_table_to_excel(
            missing_report,
            ["Пункт", "Текст пункту", "Причина"],
            [
                (item.get("label", ""), item.get("text", ""), item.get("reason", ""))
                for item in unmatched_items
            ],
        )
        if unmatched_items:
            self.log(
                f"УВАГА: {len(unmatched_items)} пункт(ів) без адресата. "
                f"Контрольний файл: {missing_report}"
            )

        # Рахуємо записи аудиту, а не унікальні мітки: однакова мітка («1.»)
        # трапляється в різних § і тоді множина злила б їх в один пункт.
        audited_items = map_res.get("routing_audit", [])
        routed_count = sum(
            1
            for item in audited_items
            if str(item.get("final_recipients", "")).strip() not in ("", "—")
        )
        skipped_items = map_res.get("skipped_items", [])
        self.log(
            f"Розібрано пунктів: {len(audited_items)}; "
            f"з адресатами: {routed_count}; без адресата: {len(unmatched_items)}; "
            f"виключено із загального переліку (зміна до управління): {len(skipped_items)}."
        )

        units_data, management_data = select_extracts_for_scope(map_res, scope)
        # Зміни по управлінню нікому не розсилаються, тому вони НЕ входять до
        # розрахунку розсилки. Але витяг за таким пунктом усе одно потрібен —
        # окремим файлом, без адресата й без компонування під друк «2 на 1».
        if not units_data and not management_data:
            if scope == "management":
                empty_message = "У вибраному наказі немає пунктів для витягів до управління."
            elif scope == "general":
                empty_message = "У вибраному наказі немає адресних витягів для військових частин."
            else:
                empty_message = "Жодної військової частини або пункту до управління не знайдено."
            self.log(empty_message)
            messagebox.showwarning("Результат", empty_message)
            return
        if management_data:
            self.log(
                f"Витяги до управління (окремий файл, поза розрахунком розсилки): "
                f"{len(management_data)} — " + ", ".join(management_data)
            )

        # Склад кожного витягу з рядками наказу, з яких береться пункт. Саме тут
        # видно причину задвоєного пункту: два записи з ОДНАКОВИМИ рядками або
        # шапка, чиї рядки перекриваються з рядками самого пункту.
        composition_report = control_report_path(self.out_folder.get(), "Контроль_складу_витягів", order_base)
        composition_rows = []
        duplicate_spans = 0
        for unit_key, unit_data in {**units_data, **management_data}.items():
            seen_spans: dict[tuple, str] = {}
            for unit_item in sorted(
                unit_data.get("items", []),
                key=lambda value: value.get("source_start_line", 10**9),
            ):
                span = (unit_item.get("source_start_line"), unit_item.get("source_end_line"))
                headings = ", ".join(
                    f"{heading[0]}–{heading[1]}"
                    for heading in (unit_item.get("heading_ranges") or [])
                    if isinstance(heading, (list, tuple)) and len(heading) == 2
                )
                note = ""
                if span in seen_spans:
                    note = f"ДУБЛЬ рядків із пунктом «{seen_spans[span]}»"
                    duplicate_spans += 1
                else:
                    seen_spans[span] = unit_item.get("label", "")
                composition_rows.append(
                    (
                        unit_key,
                        unit_item.get("label", ""),
                        f"{span[0]}–{span[1]}",
                        headings or "—",
                        (unit_item.get("text", "") or "").strip()[:120],
                        note,
                    )
                )
        _save_table_to_excel(
            composition_report,
            ["Витяг", "Пункт", "Рядки наказу", "Рядки шапок", "Початок тексту", "Примітка"],
            composition_rows,
        )
        if duplicate_spans:
            self.log(
                f"УВАГА: маршрутизація віддала {duplicate_spans} пункт(ів) із тими самими "
                f"рядками наказу — у витягу вони не задвоюються, але причину видно у файлі "
                f"{os.path.basename(composition_report)}."
            )

        # Пункт, який зник БЕЗ СЛІДУ: його номер є в тексті наказу, але він не
        # потрапив ні у витяги, ні в «Пропущені», ні у виключені. Так буває,
        # коли пункт поглинула шапка розділу або сусідній пункт — тоді він
        # узагалі не існує як окремий блок, і жоден інший контроль його не
        # показує. Рахуємо лише «дірки» всередині наявної нумерації, щоб не
        # чіплятися до нумерованих рядків усередині самих пунктів.
        item_number_re = re.compile(r"^\s*(\d{1,3})\s*[.)]")

        def numbers_from_labels(labels) -> set[int]:
            found = set()
            for label in labels:
                match = re.search(r"(\d{1,3})", str(label))
                if match:
                    found.add(int(match.group(1)))
            return found

        numbers_in_text = {
            int(match.group(1))
            for match in (item_number_re.match(line) for line in text.splitlines())
            if match
        }
        numbers_reached = numbers_from_labels(
            [item.get("label", "") for data in units_data.values() for item in data.get("items", [])]
            + [item.get("label", "") for item in unmatched_items]
            + [item.get("label", "") for item in skipped_items]
        )
        lost_numbers = sorted(
            number
            for number in numbers_in_text - numbers_reached
            if numbers_reached and number < max(numbers_reached)
        )
        if lost_numbers:
            self.log(
                "УВАГА: у тексті наказу є пункт(и) "
                + ", ".join(str(number) for number in lost_numbers)
                + ", але їх немає ні у витягах, ні в контрольних переліках. "
                "Найімовірніше, такий пункт поглинула шапка розділу або сусідній "
                f"пункт — подивіться діапазони рядків у файлі {os.path.basename(composition_report)}."
            )

        template_path = os.path.abspath(self.template_path.get())
        out_file = os.path.join(
            self.out_folder.get(), build_extracts_filename(order_num, order_date)
        )
        os.makedirs(self.out_folder.get(), exist_ok=True)
        # Наприкінці генерації результат відкривається у Word, тому при
        # повторному запуску він може бути ще зайнятий. Перевіряємо одразу,
        # щоб не витрачати час на обробку й показати зрозумілу причину.
        management_out_file = (
            os.path.join(
                self.out_folder.get(),
                build_extracts_filename(order_num, order_date, "Витяги до управління за наказом"),
            )
            if management_data
            else None
        )
        busy_candidates = []
        if units_data:
            busy_candidates.append(out_file)
        if management_out_file:
            busy_candidates.append(management_out_file)
        for busy_candidate in busy_candidates:
            if busy_candidate and not is_path_writable(busy_candidate):
                raise UserError(
                    f"файл витягів «{os.path.basename(busy_candidate)}» уже відкритий — "
                    "найімовірніше у Word, з минулого разу.",
                    "Закрийте це вікно Word і натисніть кнопку ще раз. "
                    "Поки файл відкритий, програма не може записати в нього нові витяги.",
                )
        temp_dir = os.path.join(self.out_folder.get(), "_nat_temp")
        os.makedirs(temp_dir, exist_ok=True)

        temp_files = []
        management_temp_files = []
        layout_warnings = []

        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0  # wdAlertsNone: не показувати блокуючі діалоги (напр. "Зберегти зміни?")

        word_settings = tune_word_for_batch(word)
        batch_started_at = time.monotonic()

        try:
            source_path = os.path.abspath(self.doc_path.get())
            source_doc = word.Documents.Open(source_path, ReadOnly=True)

            self.log("Індексуємо абзаци оригіналу наказу...")
            source_paragraphs = iter_paragraphs(source_doc)
            para_count = len(source_paragraphs)
            source_line_to_para = []
            for pi, source_paragraph in enumerate(source_paragraphs, start=1):
                raw = source_paragraph.Range.Text
                source_line_to_para.extend([pi] * paragraph_line_count(raw))

            usable_line_count = len(text.splitlines())
            source_line_to_para = source_line_to_para[:usable_line_count]
            self.log(f"Проіндексовано {para_count} абзаців.")

            def has_manual_page_break(para_or_range):
                r = getattr(para_or_range, "Range", para_or_range)
                return "\x0c" in getattr(r, "Text", "")

            def is_page_break_only(para_or_range):
                r = getattr(para_or_range, "Range", para_or_range)
                return is_page_break_only_text(getattr(r, "Text", "") or "")

            def remove_page_break_chars(start, end):
                return remove_page_breaks_in_range(doc, start, end)

            def is_blank_paragraph(para_or_range):
                r = getattr(para_or_range, "Range", para_or_range)
                return not getattr(r, "Text", "").strip("\r\x07\v\f \t")

            def clean_redundant_blanks():
                """Прибирає зайві «висячі» порожні абзаци в самому кінці документа,
                щоб дотримати правило 5.3 AGENT.md: 0 порожніх абзаців наприкінці."""
                try:
                    while doc.Paragraphs.Count > 1:
                        tail_paragraph = last_paragraph(doc)
                        if not is_blank_paragraph(tail_paragraph):
                            break
                        # Останній знак абзацу документа (і останній у комірці)
                        # Word не видаляє й помилки не кидає. Без перевірки, що
                        # документ справді скоротився, цикл був би вічним.
                        length_before = doc.Content.End
                        tail_paragraph.Range.Delete()
                        if doc.Content.End >= length_before:
                            break
                except Exception:
                    pass

            def normalize_signature_gap(signer_start, content_end_pos):
                """Нормалізує відступ перед підписантом: рівно 2 порожні абзаци.

                Після видалення/додавання абзаців позиції зсуваються, тому
                підписант шукається заново по непорожньому тексту після content_end_pos.
                """
                document_paragraphs = iter_paragraphs(doc)
                signer_index = None
                for paragraph_index, paragraph in enumerate(document_paragraphs, start=1):
                    if paragraph.Range.Start == signer_start:
                        signer_index = paragraph_index
                        break
                if signer_index is None:
                    return signer_start

                blank_indexes = []
                for paragraph_index in range(signer_index - 1, 0, -1):
                    if not is_blank_paragraph(document_paragraphs[paragraph_index - 1]):
                        break
                    blank_indexes.append(paragraph_index)

                # Перед підписантом залишаємо рівно два порожні абзаци
                while len(blank_indexes) < 2:
                    new_p = doc.Paragraphs.Add(doc.Paragraphs(signer_index).Range)
                    new_p.Range.Text = "\r"
                    blank_indexes.append(signer_index)
                    signer_index += 1
                if len(blank_indexes) > 2:
                    for paragraph_index in sorted(blank_indexes[2:], reverse=True):
                        doc.Paragraphs(paragraph_index).Range.Delete()

                # Після видалення/вставки — позиції зсунулись, шукаємо підписанта
                # заново, але лише в хвості документа після вставленого змісту.
                for paragraph in iter_paragraphs(doc.Range(content_end_pos, doc.Content.End)):
                    paragraph_range = paragraph.Range
                    if paragraph_range.Start >= content_end_pos and not is_blank_paragraph(paragraph_range):
                        return paragraph_range.Start
                return signer_start


            def line_span_to_paragraphs(start_line, end_line):
                if not isinstance(start_line, int) or not isinstance(end_line, int):
                    return []
                if start_line < 0 or end_line < start_line:
                    return []
                # Обрізаємо до останнього валідного індексу замість повернення []
                end_line = min(end_line, len(source_line_to_para) - 1)
                if start_line >= len(source_line_to_para):
                    return []
                return list(dict.fromkeys(source_line_to_para[start_line:end_line + 1]))

            def bind_range(range_object):
                """Зчіплює ВСІ абзаци діапазону, включно з порожніми всередині.

                Word застосовує ParagraphFormat діапазону до кожного його абзацу,
                тож це два звернення замість двох на КОЖЕН абзац.
                """
                pf = range_object.ParagraphFormat
                pf.KeepTogether = True
                pf.KeepWithNext = True

            def release_last_paragraph(range_object):
                """Знімає KeepWithNext з ОСТАННЬОГО абзацу діапазону — саме там
                дозволено розрив сторінки (між пунктами, після ланцюга)."""
                try:
                    last_paragraph(range_object).Range.ParagraphFormat.KeepWithNext = False
                except Exception:
                    pass

            def bind_span(start_position, end_position):
                """Зчіплює діапазон, заданий позиціями."""
                if start_position is None or end_position is None:
                    return
                if end_position <= start_position:
                    return
                bind_range(doc.Range(start_position, end_position))

            def apply_keep_rules(item_ranges, heading_ranges, heading_item_pairs, signer_start, executor_start=None):
                # Рідкісний Word-баг: FormattedText переносить KeepWithNext із
                # наказу, а вставлений після пункту порожній абзац іноді
                # успадковує цей прапорець. Тоді два сусідні пункти випадково
                # стають одним великим нерозривним блоком і псують підбір
                # інтервалу. Спочатку скидаємо лише зчеплення, а нижче заново
                # накладаємо всі наші правила шапок, пунктів і підписанта.
                all_content_ranges = [*heading_ranges, *item_ranges]
                if all_content_ranges:
                    keep_start = min(value.Start for value in all_content_ranges)
                    keep_end = signer_start
                    if keep_end is None:
                        keep_end = max(value.End for value in all_content_ranges)
                    if keep_end > keep_start:
                        doc.Range(keep_start, keep_end).ParagraphFormat.KeepWithNext = False

                # 1. Шапки (§, розділ, підрозділ) зчіплюються з ПЕРШИМ пунктом
                # розділу — разом з порожніми абзацами між ними.
                #
                # Раніше KeepWithNext діставався лише абзацам самої шапки. Але
                # після кожної шапки генератор вставляє порожній абзац, і саме
                # він лишався без зчеплення: Word чесно тримав шапку з наступним
                # абзацом (порожнім) і мав повне право рвати сторінку одразу за
                # ним. Через це шапка лишалась унизу сторінки, а пункт їхав на
                # наступну. Зчіплюємо суцільним діапазоном — так само, як це
                # зроблено для фінального ланцюга в пункті 3 нижче.
                #
                # Зчіплюємо кожну шапку з тим, що стоїть ПІД нею — до початку
                # наступного вставленого блока (це або наступний рівень шапки,
                # або сам пункт). Так само тримається й шапка, для якої пари
                # «шапка → пункт» не склалось: інакше вона висіла б сама.
                blocks = [(heading_range.Start, "heading", heading_range) for heading_range in heading_ranges]
                blocks += [(item_range.Start, "item", item_range) for item_range in item_ranges]
                blocks.sort(key=lambda block: block[0])
                for index, (start_position, kind, range_object) in enumerate(blocks):
                    if kind != "heading":
                        continue
                    if index + 1 < len(blocks):
                        bind_span(start_position, blocks[index + 1][0])
                    else:
                        bind_range(range_object)

                # 2. Пункти витягу: кожен пункт цілісний
                for item_index, item_range in enumerate(item_ranges):
                    bind_range(item_range)
                    if item_index != len(item_ranges) - 1:
                        # Після пункту сторінку рвати можна — крім останнього
                        # пункту, який зчеплений з підписантом (ланцюг нижче).
                        release_last_paragraph(item_range)

                # 3. Суцільний нерозривний ланцюг: [Шапка перед останнім пунктом (якщо є)] -> [Останній пункт] -> [Відступ] -> [Підписант] -> [Згідно з оригіналом]
                # Гарантує, що шапка, останній пункт і підписант становлять неподільний блок і вміщуються разом!
                if item_ranges:
                    last_item_range = item_ranges[-1]
                    chain_start = last_item_range.Start
                    for heading_range, first_item_range, _ in heading_item_pairs:
                        if first_item_range.Start == last_item_range.Start:
                            chain_start = heading_range.Start
                            break

                    signature_end = executor_start if executor_start is not None else doc.Content.End
                    chain_range = doc.Range(chain_start, signature_end)
                    bind_range(chain_range)
                    release_last_paragraph(chain_range)

            def layout_issues(item_ranges, item_labels, heading_item_pairs, signer_start=None, executor_start=None):
                issues = []
                for item_range, label in zip(item_ranges, item_labels):
                    start_page, end_page = range_pages(item_range)
                    if start_page is not None and start_page != end_page:
                        issues.append(f"{label} не вміщується на одну сторінку")

                for heading_range, first_item_range, label in heading_item_pairs:
                    heading_start, heading_end = range_pages(heading_range)
                    item_start, _ = range_pages(first_item_range)
                    if (
                        heading_start is not None
                        and (heading_start != heading_end or heading_start != item_start)
                    ):
                        issues.append(f"шапка перед {label} відірвана від першого пункту")

                # Контроль нерозривності останнього пункту (та його шапки) з підписантом та блоком «Згідно з оригіналом»:
                if item_ranges and signer_start is not None:
                    last_item_range = item_ranges[-1]
                    last_item_start_p, _ = range_pages(last_item_range)
                    signer_p = page_of(signer_start)
                    sig_end_pos = max(signer_start, (executor_start - 1) if executor_start is not None else (doc.Content.End - 1))
                    zg_end_p = page_of(sig_end_pos)

                    # Перевірка для шапки перед останнім пунктом:
                    for heading_range, first_item_range, label in heading_item_pairs:
                        if first_item_range.Start == last_item_range.Start:
                            h_start, h_end = range_pages(heading_range)
                            if h_start is not None and signer_p is not None and h_start != signer_p:
                                issues.append("шапка, останній пункт та підписант розділені на різні сторінки")

                    if last_item_start_p is not None and signer_p is not None and last_item_start_p != signer_p:
                        issues.append("підписант та блок «Згідно з оригіналом» відірвані від останнього пункту")
                    elif signer_p is not None and zg_end_p is not None and signer_p != zg_end_p:
                        issues.append("блок «Згідно з оригіналом» відірваний від підписанта наказу")

                return list(dict.fromkeys(issues))

            def find_executor_paragraph_index(bookmark_name=None):
                """Знаходить номер абзацу виконавця за закладкою, текстом або як останній непорожній абзац."""
                # 1. За закладкою — точна позиція підстановки {{виконавець}}
                if bookmark_name and doc.Bookmarks.Exists(bookmark_name):
                    try:
                        b_pos = doc.Bookmarks(bookmark_name).Range.Start
                        for pi, paragraph in enumerate(iter_paragraphs(doc), start=1):
                            pr = paragraph.Range
                            if pr.Start <= b_pos < pr.End:
                                return pi
                    except Exception:
                        pass

                document_paragraphs = iter_paragraphs(doc)

                # 2. За текстом виконавця (фолбек, якщо закладку втрачено)
                exec_val = self.executor.get().strip()
                if exec_val:
                    lines = [ln.strip() for ln in exec_val.replace("/", "\n").splitlines() if ln.strip()]
                    if lines:
                        first_line = lines[0].lower()
                        for pi in range(len(document_paragraphs), 0, -1):
                            pt = document_paragraphs[pi - 1].Range.Text.strip().lower()
                            if first_line in pt or (len(first_line) > 4 and first_line[:4] in pt):
                                return pi

                # 3. Фолбек: останній непорожній абзац у документі
                for pi in range(len(document_paragraphs), 0, -1):
                    if not is_blank_paragraph(document_paragraphs[pi - 1]):
                        return pi
                return None

            def executor_start_from_bookmark(bookmark_name):
                """Повертає позицію (Start) закладки виконавця, або None, якщо її немає."""
                if bookmark_name and doc.Bookmarks.Exists(bookmark_name):
                    try:
                        return doc.Bookmarks(bookmark_name).Range.Start
                    except Exception:
                        return None
                return None

            def executor_paragraph_start(bookmark_name):
                """Початок АБЗАЦУ виконавця без проходу по всіх абзацах документа.

                Закладка віддає позицію одразу, а `Range(pos, pos).Paragraphs(1)`
                — її абзац. Потрібно для меж, у яких дозволено міняти інтервал:
                блок виконавця не чіпаємо ніколи.
                """
                position = executor_start_from_bookmark(bookmark_name)
                if position is None:
                    return None
                try:
                    return doc.Range(position, position).Paragraphs(1).Range.Start
                except Exception:
                    return None

            def executor_paragraph(bookmark_name=None):
                """Абзац виконавця як ОБ'ЄКТ — із закладки, без обходу документа.

                Увага: сама закладка для цього не годиться після вставок —
                вставлений перед нею текст Word включає В НЕЇ. Тому позицію з
                закладки беремо один раз, а далі тримаємось за абзац.
                """
                position = executor_start_from_bookmark(bookmark_name)
                if position is not None:
                    try:
                        return doc.Range(position, position).Paragraphs(1)
                    except Exception:
                        pass
                index = find_executor_paragraph_index(bookmark_name)
                if index is None:
                    return None
                try:
                    return doc.Paragraphs(index)
                except Exception:
                    return None

            def previous_paragraph(paragraph):
                """Попередній абзац або None (на початку документа Word кидає помилку)."""
                try:
                    return paragraph.Previous()
                except Exception:
                    return None

            def paragraph_in_table(paragraph):
                """Чи стоїть абзац усередині таблиці (wdWithInTable)."""
                try:
                    return bool(paragraph.Range.Information(12))
                except Exception:
                    # Не змогли визначити — вважаємо, що в таблиці: чіпати
                    # невідомий абзац небезпечніше, ніж лишити його на місці.
                    return True

            def delete_paragraph_if_possible(paragraph):
                """Видаляє абзац і повертає True, лише якщо він СПРАВДІ зник.

                Word мовчки ігнорує видалення знака абзацу, який прибрати не
                можна: останнього в документі та останнього в комірці таблиці.
                Помилки при цьому не буде, тож цикл, який спирається на
                `Delete()`, без перевірки прогресу крутиться вічно.
                """
                try:
                    before = doc.Content.End
                    paragraph.Range.Delete()
                    return doc.Content.End < before
                except Exception:
                    return False

            def position_executor_at_page_bottom(bookmark_name=None, signer_start=None, needs_manual_review=False):
                """Вирівнює виконавця до низу сторінки, окрім випадків, де це неможливо
                без наставляння штучних порожніх абзаців — тоді лишаємо позицію зі зразка
                (правило 5.5 AGENT.md, виключення №2).

                needs_manual_review навмисно НЕ блокує спробу підштовхування: загальна
                проблема макета деінде у витягу (напр. інший пункт не вміщується) не
                означає, що саме область виконавця не можна коректно розмістити.
                Єдина причина відкату — фізична неможливість (перехід на іншу сторінку),
                що обробляється нижче в самому циклі."""
                # Працюємо з АБЗАЦОМ-ОБ'ЄКТОМ, а не з його номером: `doc.Paragraphs(i)`
                # щоразу проходить документ, а самих звернень тут десятки. Абзац
                # Word тримає прив'язку до свого тексту навіть після вставок перед
                # ним (перевірено), тож об'єкт лишається чинним увесь підбір.
                executor = executor_paragraph(bookmark_name)
                if executor is None:
                    clean_redundant_blanks()
                    return

                # Запам'ятовуємо кількість порожніх абзаців зі зразка перед виконавцем,
                # щоб мати змогу відновити цю позицію, якщо автоматичне вирівнювання
                # донизу не вдасться виконати чисто (без переходу на іншу сторінку).
                # Порожні комірки таблиці порожніми абзацами зразка НЕ рахуємо:
                # у шаблоні з табличним підписантом (див. drop_blanks_before_executor)
                # виконавцю передує саме комірка, а не відступ.
                original_blank_count = 0
                probe = previous_paragraph(executor)
                while (
                    probe is not None
                    and not paragraph_in_table(probe)
                    and is_blank_paragraph(probe)
                ):
                    original_blank_count += 1
                    probe = previous_paragraph(probe)

                def drop_blanks_before_executor():
                    """Знімає порожні абзаци безпосередньо перед виконавцем.

                    Зупиняємось на межі таблиці: якщо блок підписанта/засвідчувача
                    у зразку зверстано таблицею, попередній абзац — це остання
                    (часто порожня) КОМІРКА. Її знак абзацу Word не видаляє й не
                    повідомляє про це помилкою, тож цикл крутився вічно — програма
                    зависала на першому ж витягу. Комірка тут і не є відступом:
                    зсовувати виконавця треба порожніми абзацами ПІСЛЯ таблиці.
                    """
                    while True:
                        previous = previous_paragraph(executor)
                        if (
                            previous is None
                            or paragraph_in_table(previous)
                            or not is_blank_paragraph(previous)
                        ):
                            return previous
                        if not delete_paragraph_if_possible(previous):
                            return previous

                previous_before_executor = drop_blanks_before_executor()
                if previous_before_executor is not None:
                    try:
                        previous_before_executor.Range.ParagraphFormat.KeepWithNext = False
                    except Exception:
                        pass

                clean_redundant_blanks()
                doc.Repaginate()

                def restore_sample_position():
                    """Відновлює рівно original_blank_count порожніх абзаців перед
                    виконавцем (позиція зі зразка) без штучного доштовхування донизу."""
                    drop_blanks_before_executor()
                    if original_blank_count:
                        position = executor.Range.Start
                        doc.Range(position, position).InsertBefore("\r" * original_blank_count)
                    doc.Repaginate()

                # Опускаємо виконавця до самого низу поточної сторінки.
                #
                # ВАЖЛИВО: вертикальну позицію (Information(6),
                # wdVerticalPositionRelativeToPage) тут НЕ використовуємо — у
                # прихованому екземплярі Word (Visible=False) вона стабільно кидає
                # E_FAIL (-2147467259), через що опускання мовчки не спрацьовувало.
                # Орієнтуємось виключно на номер сторінки (Information(3)), який
                # працює надійно: додаємо порожні абзаци, доки виконавець не
                # перескочить на наступну сторінку, і прибираємо останню порцію —
                # це і є найнижча позиція, яка ще вміщується на сторінці.
                # Кількість ентерів підбираємо ДВІЙКОВИМ пошуком: ознака
                # «виконавець ще на своїй сторінці» монотонна (що більше
                # ентерів, то нижче), тож замість до 38 покрокових вставок,
                # кожна зі своїм Repaginate, вистачає ~7 вимірювань.
                current_enters = 0

                def set_enters(count):
                    """Доводить кількість порожніх абзаців перед виконавцем до `count`."""
                    nonlocal current_enters
                    position = executor.Range.Start
                    if count > current_enters:
                        doc.Range(position, position).InsertBefore("\r" * (count - current_enters))
                    elif count < current_enters:
                        # Кожен доданий порожній абзац — рівно один символ,
                        # тож зайві знімаємо одним діапазоном, а не по одному.
                        delta = current_enters - count
                        doc.Range(max(0, position - delta), position).Delete()
                    current_enters = count
                    doc.Repaginate()

                def executor_still_on_page(count, start_page):
                    set_enters(count)
                    return page_of(executor.Range.Start) == start_page

                try:
                    start_page = page_of(executor.Range.Start)
                    if start_page is None:
                        self.log(f"  ⚠️ {cipher}: не вдалося визначити сторінку виконавця — лишено як у зразку.")
                    else:
                        # fits — найбільша кількість ентерів, яка ще тримає
                        # виконавця на його сторінці; overflow — перша, яка вже ні.
                        fits, overflow = 0, 80
                        if executor_still_on_page(overflow, start_page):
                            fits = overflow
                        else:
                            while overflow - fits > 1:
                                middle = (fits + overflow) // 2
                                if executor_still_on_page(middle, start_page):
                                    fits = middle
                                else:
                                    overflow = middle
                            set_enters(fits)
                        self.log(f"  ↧ {cipher}: виконавця опущено вниз сторінки (+{fits} ентер(ів)).")
                except Exception as exec_error:
                    self.log(f"  ⚠️ {cipher}: помилка опускання виконавця: {exec_error}")
                    try:
                        # Повертаємо відступ зі зразка, щоб не лишити виконавця
                        # без жодного порожнього абзацу після невдалої спроби.
                        restore_sample_position()
                    except Exception:
                        pass

                clean_redundant_blanks()

            def page_of(position):
                try:
                    return doc.Range(position, position).Information(3)  # wdActiveEndPageNumber
                except Exception:
                    return None

            def range_pages(item_range):
                start_page = page_of(item_range.Start)
                end_position = max(item_range.Start, item_range.End - 1)
                return start_page, page_of(end_position)


            # Витяги управління формуються тим самим проходом і за тим самим
            # зразком: різниця лише в порожніх {{кому}}/{{куди}} та в тому, що
            # вони збираються в окремий файл без компонування під друк.
            generation_plan = [(key, value, False) for key, value in units_data.items()]
            generation_plan += [(key, value, True) for key, value in management_data.items()]
            self.progress_begin(len(generation_plan) + 1, "Витяги")
            for idx, (cipher, data, is_management) in enumerate(generation_plan):
                self.log(f"[{idx+1}/{len(generation_plan)}] Генеруємо витяг для: {cipher}")
                self.progress_step(idx, cipher)
                extract_started_at = time.monotonic()

                extract_needs_manual_review = False
                # Початок вставленого змісту: усе, що ВИЩЕ цієї позиції (герб,
                # {{кому}}/{{куди}}, назва, дата/номер) — недоторкана шапка зразка.
                content_start_pos = None
                signer_start = None
                # Робимо копію шаблону, щоб Word COM ніколи не відкривав і не
                # змінював оригінальний файл шаблону. Розширення копії має
                # відповідати реальному вмісту шаблону (він може бути у форматі
                # Word 97-2003), інакше Word відмовиться її відкрити.
                temp_path = copy_template_for_editing(
                    template_path,
                    os.path.join(temp_dir, f"extract_{idx:04d}.docx"),
                    label="зразок витягу",
                )
                doc = word.Documents.Open(os.path.abspath(temp_path), ReadOnly=False)
                # Текст шаблону читаємо один раз: більшості з тих двох десятків
                # тегів у конкретному зразку немає (різні регістри «кому/КОМУ»,
                # окремі теги засвідчувача), а порожній пошук Word коштує
                # дорожче, ніж це читання.
                try:
                    template_text = str(doc.Content.Text or "").casefold()
                except Exception:
                    template_text = None

                def replace_tag(tag, replacement_text, document=doc, collect_paragraphs=False,
                               highlight_red=False, bold_pattern=None):
                    if template_text is not None and tag.casefold() not in template_text:
                        return []
                    replaced_paragraphs = []
                    find_obj = document.Content.Find
                    find_obj.Text = tag
                    while find_obj.Execute():
                        found_start = find_obj.Parent.Start
                        find_obj.Parent.Text = str(replacement_text)
                        found_end = found_start + len(str(replacement_text))
                        repl_range = document.Range(found_start, found_end)
                        if highlight_red:
                            try:
                                repl_range.Font.Color = 255  # wdColorRed (RGB 255, 0, 0)
                                repl_range.Font.Bold = 1
                            except Exception:
                                pass
                        if bold_pattern:
                            # Жирним виділяється лише ЧАСТИНА підставленого
                            # тексту: цифри дня в даті та номер наказу без «№».
                            match = re.search(bold_pattern, str(replacement_text))
                            if match:
                                try:
                                    document.Range(
                                        found_start + match.start(), found_start + match.end()
                                    ).Font.Bold = 1
                                except Exception:
                                    pass
                        if collect_paragraphs:
                            replaced_paragraphs.append(
                                document.Range(found_start, found_start).Paragraphs(1).Range.Duplicate
                            )
                        find_obj = document.Content.Find
                        find_obj.Text = tag
                    return replaced_paragraphs

                def replace_tag_group(tag, replacement_text, **kwargs):
                    """Підставляє значення в усі синоніми тега зі зразка.

                    У заготовках трапляються і «засвідчувач», і «затверджувач»,
                    і «згідно_з_оригіналом» — це те саме поле, тому шаблон не
                    повинен залежати від того, яку назву обрав автор.
                    """
                    for alias in tag_aliases(tag):
                        replace_tag(alias, replacement_text, **kwargs)

                signer_template_tags = SIGNER_TAGS

                def remove_original_signer_template_block():
                    """Видаляє з шаблону весь блок тегів підписанта оригіналу.

                    Викликається лише тоді, коли реквізитів підписанта немає
                    зовсім. Якщо тег стоїть у таблиці, видаляється відповідний
                    рядок, щоб не лишалися порожні комірки.
                    """
                    # У зразку без цих тегів обходити абзаци нема сенсу.
                    if template_text is not None and not any(
                        tag.casefold() in template_text for tag in signer_template_tags
                    ):
                        return
                    # Щоразу шукаємо ПЕРШИЙ абзац із тегом заново. Зібраний
                    # наперед список тут не годиться: після видалення рядка
                    # таблиці абзаци-об'єкти з наступних рядків «сповзають» на
                    # сусідній рядок, і разом із підписантом зникав рядок
                    # «Згідно з оригіналом». Тегів одиниці, а шаблон на цьому
                    # кроці ще без змісту — перепрохід коштує копійки.
                    for _ in range(len(signer_template_tags) * 4):  # запобіжник від вічного циклу
                        target_range = None
                        for paragraph in iter_paragraphs(doc):
                            paragraph_range = paragraph.Range
                            text_of_paragraph = str(paragraph_range.Text or "").casefold()
                            if any(tag.casefold() in text_of_paragraph for tag in signer_template_tags):
                                target_range = paragraph_range
                                break
                        if target_range is None:
                            return
                        try:
                            if target_range.Information(12):  # wdWithInTable
                                target_range.Cells(1).Row.Delete()
                            else:
                                length_before = doc.Content.End
                                target_range.Delete()
                                if doc.Content.End >= length_before:
                                    # Знак абзацу Word не віддав (останній у
                                    # документі) — прибираємо принаймні тег.
                                    target_range.Text = ""
                        except Exception:
                            try:
                                # Абзац видалити не вдалось — прибираємо хоча б
                                # текст, інакше наступний прохід знайде той самий.
                                target_range.Text = ""
                            except Exception:
                                return

                if is_management:
                    # Такий витяг нікуди не надсилається: адресні теги просто
                    # стираємо, а не підсвічуємо червоним для ручної правки.
                    rec_to_val = ""
                    dest_where_val = ""
                    is_dest_manual = False
                else:
                    rec_to_val = data.get("recipient_to") or cipher
                    dest_where_val = (data.get("destination_where") or "").strip()
                    is_dest_manual = not bool(dest_where_val) or dest_where_val.upper() in ("КУДИ", "[КУДИ]")
                    if is_dest_manual:
                        dest_where_val = "КУДИ"

                for tag_var in ("{{кому}}", "{{Кому}}", "{{КОМУ}}"):
                    replace_tag(tag_var, rec_to_val)
                for tag_var in ("{{куди}}", "{{Куди}}", "{{КУДИ}}"):
                    replace_tag(tag_var, dest_where_val, highlight_red=is_dest_manual)
                if order_date_formatted:
                    # Жирним — лише цифри дня: “29” серпня 2026 року.
                    replace_tag_group("{{дата_наказу}}", order_date_formatted,
                                      bold_pattern=r"\d+")
                if order_num:
                    # Жирним — лише номер, знак «№» лишається звичайним.
                    replace_tag_group("{{номер_наказу}}", f"№{order_num}",
                                      bold_pattern=r"(?<=№).+")

                # Підписант оригіналу наказу
                if any(order_signer.values()):
                    for signer_tag, signer_value in signer_tags(
                        _slash_to_lines(order_signer["position"]),
                        order_signer["rank"],
                        order_signer["name"],
                    ).items():
                        if signer_value:
                            replace_tag(signer_tag, signer_value)
                    # Реквізит, якого в наказі немає (буває звання), просто
                    # стираємо. Видаляти тут рядок таблиці НЕ можна: у ньому
                    # стоять сусідні, вже заповнені теги підписанта.
                    for empty_tag in signer_template_tags:
                        replace_tag(empty_tag, "")
                else:
                    # Реквізитів немає зовсім — прибираємо блок разом з
                    # абзацами та рядками таблиці, щоб не лишити порожнечі.
                    remove_original_signer_template_block()

                # Особа, яка засвідчує витяг («Згідно з оригіналом» / Засвідчувач)
                replace_tag_group("{{засвідчення}}", "Згідно з оригіналом")
                for cert_tag, cert_value in certifier_tags(
                    _slash_to_lines(self.certifier_position.get().strip()),
                    self.certifier_rank.get().strip(),
                    self.certifier_name.get().strip(),
                ).items():
                    if cert_value:
                        replace_tag(cert_tag, cert_value)

                executor_bookmark = None
                if self.executor.get().strip():
                    executor_paragraphs = replace_tag(
                        "{{виконавець}}", _slash_to_lines(self.executor.get().strip()), collect_paragraphs=True
                    )
                    if executor_paragraphs:
                        executor_bookmark = f"nat_executor_{idx}"
                        doc.Bookmarks.Add(executor_bookmark, executor_paragraphs[0])
                        if len(executor_paragraphs) > 1:
                            layout_warnings.append(
                                f"{cipher}: у зразку кілька тегів {{виконавець}}; донизу вирівняно лише перший."
                            )

                raw_labels = [item.get("label", "") for item in data.get("items", [])]
                points_text = _format_item_numbers_range(raw_labels)
                replace_tag("{{пункти}}", points_text)

                # Вставка змісту
                find_zmist = doc.Content.Find
                find_zmist.Text = "{{зміст}}"
                if find_zmist.Execute():
                    zmist_rng = find_zmist.Parent
                    items = sorted(
                        data.get("items", []),
                        key=lambda item: item.get("source_start_line", 10**9),
                    )
                    zmist_para = zmist_rng.Paragraphs(1).Range
                    insert_point = zmist_para.Start
                    content_start_pos = insert_point
                    zmist_para.Delete()
                    inserted_item_ranges = []
                    inserted_item_labels = []
                    inserted_heading_ranges = []
                    heading_item_pairs = []
                    copied_heading_keys = ()
                    extract_needs_manual_review = False

                    def insert_empty_paragraph():
                        """Вставляє один порожній абзац у поточну позицію змісту."""
                        nonlocal insert_point
                        doc.Range(insert_point, insert_point).InsertBefore("\r")
                        insert_point += 1

                    def source_runs(paragraph_indexes):
                        """Ділить абзаци наказу на суцільні прогони для копіювання.

                        Розрив прогону робить лише РУЧНИЙ розрив сторінки в наказі:
                        такий абзац у витяг не переноситься, пагінація будується
                        заново. Наявність розриву перевіряємо одним читанням тексту
                        всього діапазону, а поабзацно — тільки якщо він там справді є.
                        """
                        if not paragraph_indexes:
                            return []
                        span_text = ""
                        try:
                            span_text = str(
                                source_doc.Range(
                                    source_paragraphs[paragraph_indexes[0] - 1].Range.Start,
                                    source_paragraphs[paragraph_indexes[-1] - 1].Range.End,
                                ).Text
                                or ""
                            )
                        except Exception:
                            span_text = "\x0c"  # не вдалося прочитати — перевіряємо поабзацно
                        if "\x0c" not in span_text:
                            return [list(paragraph_indexes)]

                        runs, current = [], []
                        for paragraph_index in paragraph_indexes:
                            # Пропускаємо лише абзац, де НІЧОГО, крім розриву, немає.
                            # Розрив у кінці абзацу з текстом (межа сторінки в наказі)
                            # раніше забирав увесь абзац — зникала біографія. Такий
                            # абзац копіюється, а сам розрив прибирає
                            # `remove_page_break_chars`.
                            if is_page_break_only(source_paragraphs[paragraph_index - 1]):
                                self.log("Пропущено вихідний розрив сторінки; пагінація витягу буде побудована заново.")
                                if current:
                                    runs.append(current)
                                    current = []
                                continue
                            current.append(paragraph_index)
                        if current:
                            runs.append(current)
                        return runs

                    def carry_geometry(source_range, run, start, end):
                        """Переносить геометрію абзаців наказу на вставлений діапазон.

                        Геометрію переносимо ЯВНО, а не покладаємось на FormattedText:
                        Word не записує властивість, яка дорівнює типовій (Alignment=Left,
                        FirstLineIndent=0), і такий абзац у витягу успадковував стиль
                        `Normal` ШАБЛОНА. Якщо в шаблоні стоїть «за шириною» та відступ
                        1.25 см — біографічний блок наказу (left 8 см, перший рядок 0,
                        уліво) з'їжджав і розтягувався по ширині, і так само «плив» §.
                        `ParagraphFormat` діапазону віддає ДІЮЧІ значення наказу.

                        Для прогону з однаковою геометрією вистачає одного
                        `ParagraphFormat` на весь діапазон. Якщо всередині геометрія
                        різна, Word віддає wdUndefined — тоді переносимо поабзацно.
                        """
                        geometry_props = ("Alignment", "LeftIndent", "RightIndent", "FirstLineIndent")

                        def read_geometry(range_object):
                            source_format = range_object.ParagraphFormat
                            values = {}
                            for prop in geometry_props:
                                try:
                                    values[prop] = getattr(source_format, prop)
                                except Exception:
                                    pass
                            return values

                        def write_geometry(range_object, values):
                            dest_format = range_object.ParagraphFormat
                            try:
                                dest_format.PageBreakBefore = False
                            except Exception:
                                pass
                            for prop, value in values.items():
                                try:
                                    setattr(dest_format, prop, value)
                                except Exception:
                                    pass

                        geometry = read_geometry(source_range)
                        if all(value != WD_UNDEFINED for value in geometry.values()):
                            write_geometry(doc.Range(start, end), geometry)
                            return

                        destination_paragraphs = iter_paragraphs(doc.Range(start, end))
                        if len(destination_paragraphs) != len(run):
                            # Кількість абзаців не збіглася (таблиця, зноска):
                            # безпечніше взяти геометрію першого абзацу прогону.
                            write_geometry(
                                doc.Range(start, end),
                                read_geometry(source_paragraphs[run[0] - 1].Range),
                            )
                            return
                        for paragraph_index, destination_paragraph in zip(run, destination_paragraphs):
                            write_geometry(
                                destination_paragraph.Range,
                                read_geometry(source_paragraphs[paragraph_index - 1].Range),
                            )

                    def insert_source_span(start_line, end_line, fallback_text, kind):
                        nonlocal insert_point
                        paragraph_indexes = line_span_to_paragraphs(start_line, end_line)
                        # Абзаци наказу беремо з готового списку: наказ відкрито
                        # лише для читання й не змінюється, тож індекси стабільні,
                        # а кожне `source_doc.Paragraphs(i)` — це прохід по всьому
                        # документу заново.
                        # Обрізаємо порожні абзаци на початку та в кінці діапазону, але зберігаємо внутрішні ентери
                        while paragraph_indexes and is_blank_paragraph(source_paragraphs[paragraph_indexes[0] - 1]):
                            paragraph_indexes.pop(0)
                        while paragraph_indexes and is_blank_paragraph(source_paragraphs[paragraph_indexes[-1] - 1]):
                            paragraph_indexes.pop()

                        first_start = None
                        last_end = None
                        # Копіюємо не по абзацу, а СУЦІЛЬНИМИ прогонами. Одне
                        # звернення до Word коштує кілька мілісекунд, а на абзац
                        # їх виходило близько п'ятнадцяти; до того ж поабзацне
                        # копіювання не відтворює таблиці. Прогін рветься лише
                        # там, де в наказі стоїть ручний розрив сторінки.
                        for run in source_runs(paragraph_indexes):
                            source_range = source_doc.Range(
                                source_paragraphs[run[0] - 1].Range.Start,
                                source_paragraphs[run[-1] - 1].Range.End,
                            )
                            start = insert_point
                            destination_range = doc.Range(start, start)
                            destination_range.FormattedText = source_range.FormattedText
                            insert_point = destination_range.End
                            carry_geometry(source_range, run, start, insert_point)
                            if has_manual_page_break(source_range):
                                insert_point = remove_page_break_chars(start, insert_point)
                            # Порожні білі зображення наказу у витяг не переносимо.
                            pasted_range = doc.Range(start, insert_point)
                            removed_images = remove_blank_images(doc, pasted_range)
                            if removed_images:
                                insert_point = pasted_range.End
                                self.log(f"  Прибрано порожніх білих зображень і фігур: {removed_images}.")
                            first_start = start if first_start is None else first_start
                            last_end = insert_point

                        if first_start is None and fallback_text:
                            self.log(f"УВАГА: не знайдено позицію оригіналу для {kind}; вставлено резервний текст без форматування.")
                            start = insert_point
                            destination_range = doc.Range(start, start)
                            compact_text = fallback_text.strip()
                            destination_range.Text = compact_text + "\r"
                            insert_point = start + len(compact_text) + 1
                            first_start, last_end = start, insert_point
                        return doc.Range(first_start, last_end) if first_start is not None else None

                    # Один і той самий шматок наказу не переноситься у витяг
                    # двічі — ні як пункт, ні як шапка. Маршрутизація може
                    # віддати той самий діапазон рядків двічі (наприклад, коли
                    # пункт потрапив у два різні § або коли шапка обчислилась
                    # по рядках самого пункту) — тоді у витягу з'являвся той
                    # самий пункт двома копіями поспіль.
                    inserted_source_spans = set()

                    # Найбільший рядок наказу, який уже перенесено у витяг.
                    # Пункти йдуть у порядку документа, тож усе, що не нижче
                    # цієї межі, у витягу ВЖЕ Є.
                    max_inserted_line = -1

                    def trim_to_new_lines(start_line, end_line, what):
                        """Обрізає діапазон до рядків наказу, яких у витягу ще немає.

                        Інваріант: кожен рядок наказу потрапляє у витяг
                        щонайбільше один раз. Перевірки на ТОЧНИЙ збіг
                        діапазонів для цього не досить — класичний випадок
                        пункт «ВИКЛАСТИ В ТАКІЙ РЕДАКЦІЇ»: усередині нього
                        цитується цілий § з преамбулою, маршрутизація бачить
                        там і «шапку», і окремий пункт, і їхні діапазони
                        ПЕРЕКРИВАЮТЬСЯ зі вступним абзацом. Через це пункт
                        друкувався у витягу двічі.

                        Повертає (початок, кінець, «усе вже перенесено»).
                        """
                        if start_line is None or end_line is None:
                            return start_line, end_line, False
                        if start_line > max_inserted_line:
                            return start_line, end_line, False
                        new_start = max_inserted_line + 1
                        fully_covered = new_start > end_line
                        message = (
                            f"{what}: рядки {start_line}–"
                            f"{min(end_line, max_inserted_line)} наказу вже перенесено вище — "
                            + ("пропущено повністю" if fully_covered
                               else f"переносимо лише {new_start}–{end_line}")
                            + " (щоб не задвоїти текст)."
                        )
                        self.log(f"  ⚠️ {cipher}: {message}")
                        layout_warnings.append(f"{cipher}: {message}")
                        if fully_covered:
                            return None, None, True
                        return new_start, end_line, False

                    def span_already_inserted(start_line, end_line):
                        # Пункт без відомих рядків наказу (вставляється резервним
                        # текстом) дублем НЕ вважається: у таких пунктів «діапазон»
                        # однаковий — (None, None), — і другий та наступні такі
                        # пункти мовчки випадали б із витягу.
                        if start_line is None or end_line is None:
                            return False
                        return (start_line, end_line) in inserted_source_spans

                    def heading_repeats_item(heading_span, item_span):
                        """Чи є «шапка» насправді тим самим текстом, що й пункт.

                        Дотик рівно одним рядком (шапка закінчується там, де
                        починається пункт) — нормальний випадок: цей рядок майже
                        завжди порожній і при вставці однаково обрізається.
                        Дублювання — це коли шапка ЗАХОДИТЬ у пункт.
                        """
                        if None in heading_span or None in item_span:
                            return False
                        return heading_span[1] > item_span[0] or heading_span[0] >= item_span[0]

                    for item in items:
                        item_span = (item.get("source_start_line"), item.get("source_end_line"))
                        item_label = item.get("label", "пункт")

                        # Перевірку на дубль робимо ДО вставки шапок: інакше
                        # шапка вже стоїть у документі, а пункт під нею
                        # пропускається — і § лишається висіти сам, без нічого,
                        # з чим його можна зчепити.
                        if span_already_inserted(item_span[0], item_span[1]):
                            message = (
                                f"{item_label}: цей самий фрагмент наказу (рядки "
                                f"{item_span[0]}–{item_span[1]}) уже перенесено — пропущено як дубль."
                            )
                            self.log(f"  ⚠️ {cipher}: {message}")
                            layout_warnings.append(f"{cipher}: {message}")
                            continue

                        heading_keys = tuple(
                            (heading_range[0], heading_range[1])
                            for heading_range in item.get("heading_ranges", [])
                            if isinstance(heading_range, (list, tuple))
                            and len(heading_range) == 2
                            and all(value is not None for value in heading_range)
                        )
                        if not heading_keys:
                            fallback_key = (item.get("heading_start_line"), item.get("heading_end_line"))
                            if all(value is not None for value in fallback_key):
                                heading_keys = (fallback_key,)

                        # Шапка, що перекривається з власним пунктом, — це не
                        # шапка: інакше текст пункту вставився б двічі поспіль.
                        kept_heading_keys = tuple(
                            heading_key for heading_key in heading_keys
                            if not heading_repeats_item(heading_key, item_span)
                        )
                        if kept_heading_keys != heading_keys:
                            message = (
                                f"{item_label}: шапка заходить у рядки самого пункту "
                                f"(шапка {heading_keys}, пункт {item_span[0]}–{item_span[1]}) — "
                                "шапку не переносимо, щоб не задвоїти пункт."
                            )
                            self.log(f"  ⚠️ {cipher}: {message}")
                            layout_warnings.append(f"{cipher}: {message}")
                        heading_keys = kept_heading_keys

                        new_heading_ranges = []
                        if heading_keys != copied_heading_keys:
                            # Порівнюємо ієрархію рівень-за-рівнем (§ → шапка →
                            # підшапка → ...) і повторюємо лише ті рівні, що
                            # реально змінилися. Спільні рівні (напр. незмінний
                            # § чи шапка одного розділу) НЕ дублюються для
                            # кожного пункту — лише нова/інша частина ієрархії.
                            common_len = 0
                            while (
                                common_len < len(heading_keys)
                                and common_len < len(copied_heading_keys)
                                and heading_keys[common_len] == copied_heading_keys[common_len]
                            ):
                                common_len += 1
                            for heading_key in heading_keys[common_len:]:
                                if span_already_inserted(heading_key[0], heading_key[1]):
                                    continue
                                heading_start, heading_end, heading_covered = trim_to_new_lines(
                                    heading_key[0], heading_key[1], f"шапка перед {item_label}"
                                )
                                if heading_covered:
                                    continue
                                heading_range = insert_source_span(
                                    heading_start, heading_end, item.get("parent_heading", ""), "заголовка"
                                )
                                if heading_range:
                                    inserted_source_spans.add((heading_key[0], heading_key[1]))
                                    if heading_end is not None:
                                        max_inserted_line = max(max_inserted_line, heading_end)
                                    inserted_heading_ranges.append(heading_range)
                                    new_heading_ranges.append(heading_range)
                                    # §, основна шапка та підшапка — окремі
                                    # елементи; після кожного лишається відступ.
                                    insert_empty_paragraph()
                        copied_heading_keys = heading_keys

                        item_start, item_end, item_covered = trim_to_new_lines(
                            item.get("source_start_line"), item.get("source_end_line"), item_label
                        )
                        if item_covered:
                            continue
                        item_range = insert_source_span(
                            item_start, item_end,
                            item.get("original_text") or item.get("text", ""), "пункту"
                        )
                        if item_range:
                            inserted_source_spans.add(item_span)
                            if item_end is not None:
                                max_inserted_line = max(max_inserted_line, item_end)
                            inserted_item_ranges.append(item_range)
                            inserted_item_labels.append(item_label)
                            heading_item_pairs.extend(
                                (heading_range, item_range, item_label)
                                for heading_range in new_heading_ranges
                            )
                            # Між кожним пунктом лишаємо один порожній абзац.
                            # Наприкінці змісту normalize_signature_gap доведе
                            # відступ перед підписантом до двох таких абзаців.
                            #
                            # Це ЄДИНИЙ порожній абзац, який додає генератор.
                            # Усе, що всередині пункту (зокрема порожній абзац
                            # перед рядком «р. н.»), переноситься 1-в-1 з наказу
                            # — у реальних наказах він там завжди є. Своїх
                            # порожніх абзаців усередину пункту не досипаємо:
                            # офіційний зразок (додаток 44) сам непослідовний,
                            # тож оригінал наказу — єдиний надійний еталон.
                            insert_empty_paragraph()

                    signer_start = None
                    content_end = insert_point
                    for paragraph_index in range(1, doc.Paragraphs.Count + 1):
                        paragraph = doc.Paragraphs(paragraph_index).Range
                        if paragraph.Start >= content_end and not is_blank_paragraph(paragraph):
                            signer_start = paragraph.Start
                            break
                    if signer_start is not None:
                        signer_start = normalize_signature_gap(signer_start, content_end)

                    executor_start = executor_start_from_bookmark(executor_bookmark)

                    if inserted_item_ranges:
                        apply_keep_rules(
                            inserted_item_ranges, inserted_heading_ranges, heading_item_pairs, signer_start, executor_start
                        )
                        doc.Repaginate()
                        # Остаточну перевірку макета (і рішення про needs_manual_review)
                        # відкладаємо до ПІСЛЯ підбору міжрядкового інтервалу нижче —
                        # адже стиснення 16→14 пт часто саме й усуває цю розбіжність.

                # Закладку виконавця НЕ видаляємо тут: вона потрібна нижче для
                # точного пошуку абзацу виконавця. Видаляємо після позиціонування.

                clean_redundant_blanks()
                doc.Repaginate()

                pages_count = doc.ComputeStatistics(2)  # wdStatisticPages

                def spacing_bounds():
                    """Межі, у яких ДОЗВОЛЕНО міняти інтервали: лише вставлений зміст
                    та блок підписанта (правило «змінюємо лише текст і підписанта»).

                    Шапка зразка (герб, {{кому}}/{{куди}}, назва, дата/номер) та блок
                    виконавця не чіпаються НІКОЛИ: точний інтервал (wdLineSpaceExactly)
                    обрізає зображення по висоті рядка — саме через це герб ховався
                    під текстом."""
                    if content_start_pos is None:
                        return None, None
                    # Межу беремо із закладки виконавця — це одне звернення до
                    # Word. Пошук абзацу перебором лишаємо фолбеком: ці межі
                    # рахуються на кожне виставлення інтервалу.
                    upper = executor_paragraph_start(executor_bookmark)
                    if upper is None:
                        upper = doc.Content.End
                        exec_idx = find_executor_paragraph_index(executor_bookmark)
                        if exec_idx:
                            try:
                                upper = doc.Paragraphs(exec_idx).Range.Start
                            except Exception:
                                upper = doc.Content.End
                    return content_start_pos, upper

                table_snapshot = None
                table_spans: list[tuple[int, int]] = []

                def apply_exact_line_spacing(points):
                    """Точний міжрядковий інтервал (wdLineSpaceExactly) на весь
                    вставлений зміст і підписанта — ОДНИМ діапазоном.

                    Раніше тут був обхід усіх абзаців документа з ~10 зверненнями
                    до Word на кожен. Оскільки інтервал перебирається, це давало
                    тисячі викликів COM на кожен витяг. Word сам застосовує
                    ParagraphFormat до всіх абзаців діапазону.

                    Порожні абзаци отримують той самий інтервал, що й текст:
                    правило «максимально заповнена сторінка» стосується всієї
                    сторінки, а різна висота порожніх рядків давала різні
                    відступи між пунктами в сусідніх витягах одного наказу.
                    """
                    nonlocal table_snapshot
                    lower, upper = spacing_bounds()
                    if lower is None or upper is None or upper <= lower:
                        return
                    target = doc.Range(lower, upper)
                    # Абзац із вбудованим зображенням точного інтервалу не
                    # отримує — інакше Word обріже картинку по висоті рядка.
                    # Запам'ятовуємо їхні інтервали й повертаємо після заміни.
                    protected = []
                    try:
                        for shape in target.InlineShapes:
                            shape_format = shape.Range.Paragraphs(1).Range.ParagraphFormat
                            protected.append(
                                (shape_format, shape_format.LineSpacingRule, shape_format.LineSpacing)
                            )
                    except Exception:
                        protected = []
                    # Таблиця зразка (підписант / «Згідно з оригіналом» /
                    # засвідчувач) — оформлення зразка, а не текст наказу: її
                    # інтервали та відступи лишаються як у зразку (користувач,
                    # 26.10.2026). Знімок береться ДО першої зміни й
                    # повертається після кожної.
                    if table_snapshot is None:
                        table_snapshot = []
                        try:
                            for table in target.Tables:
                                for paragraph in table.Range.Paragraphs:
                                    paragraph_format = paragraph.Range.ParagraphFormat
                                    table_snapshot.append((
                                        paragraph_format,
                                        paragraph_format.LineSpacingRule,
                                        paragraph_format.LineSpacing,
                                        paragraph_format.SpaceBefore,
                                        paragraph_format.SpaceAfter,
                                    ))
                                table_spans.append((table.Range.Start, table.Range.End))
                        except Exception:
                            table_snapshot = []
                    target_format = target.ParagraphFormat
                    target_format.LineSpacingRule = 4  # wdLineSpaceExactly
                    target_format.LineSpacing = points
                    target_format.SpaceBefore = 0
                    target_format.SpaceAfter = 0
                    for paragraph_format, rule, spacing, before, after in table_snapshot:
                        try:
                            paragraph_format.LineSpacingRule = rule
                            paragraph_format.LineSpacing = spacing
                            paragraph_format.SpaceBefore = before
                            paragraph_format.SpaceAfter = after
                        except Exception:
                            pass
                    for shape_format, shape_rule, shape_spacing in protected:
                        try:
                            shape_format.LineSpacingRule = shape_rule
                            shape_format.LineSpacing = shape_spacing
                        except Exception:
                            pass

                # ПІДБІР МІЖРЯДКОВОГО ІНТЕРВАЛУ — один пошук замість трьох.
                #
                # Було три послідовні перебори: 16→14 з кроком 0.5 «щоб влізло на
                # сторінку», потім ще один такий самий для багатосторінкових, потім
                # 16→14 з кроком 0.25 «щоб сторінка була заповнена». Вони міряли
                # ті самі значення й разом давали до 14 репагінацій на витяг.
                #
                # Обидві ознаки монотонні: що менший інтервал, то менше сторінок і
                # то більше пунктів вміщується на першу. Тому мета одна — НАЙБІЛЬШИЙ
                # інтервал, що дає мінімальну кількість сторінок і максимум пунктів
                # на першій сторінці. Для витягу на одну сторінку це і є «максимально
                # заповнена сторінка», для багатосторінкового — колишній «бонус».
                if content_start_pos is not None:
                    spacing_max, spacing_min, spacing_step = 16.0, 14.0, 0.25
                    spacing_grid = [
                        round(spacing_max - index * spacing_step, 2)
                        for index in range(int(round((spacing_max - spacing_min) / spacing_step)) + 1)
                    ]
                    applied_spacing = None
                    spacing_measurements: dict[float, tuple[int, int]] = {}

                    def items_fitting_first_page():
                        """Скільки пунктів ПОВНІСТЮ вміщується на першу сторінку.

                        Пункти йдуть у порядку документа, тож межу шукаємо двійково:
                        інакше на кожне вимірювання інтервалу припадало по два
                        звернення `Information` на КОЖЕН пункт.
                        """
                        if not inserted_item_ranges:
                            return 0
                        low, high = 0, len(inserted_item_ranges)
                        while low < high:
                            middle = (low + high) // 2
                            start_page, end_page = range_pages(inserted_item_ranges[middle])
                            if start_page == 1 and end_page == 1:
                                low = middle + 1
                            else:
                                high = middle
                        return low

                    def measure_spacing(points):
                        """(сторінок, пунктів на першій) для інтервалу; стан документа
                        після виклику завжди відповідає виміряному значенню."""
                        nonlocal applied_spacing
                        if applied_spacing != points:
                            apply_exact_line_spacing(points)
                            doc.Repaginate()
                            applied_spacing = points
                        if points not in spacing_measurements:
                            pages = doc.ComputeStatistics(2)  # wdStatisticPages
                            # На одній сторінці лічити нема чого: там усі пункти.
                            # Кожне звернення до номера сторінки коштує дорого.
                            spacing_measurements[points] = (
                                pages,
                                len(inserted_item_ranges) if pages <= 1 else items_fitting_first_page(),
                            )
                        return spacing_measurements[points]

                    best_possible = measure_spacing(spacing_min)
                    chosen_spacing = spacing_min
                    if measure_spacing(spacing_max) == best_possible:
                        # Стискати нема сенсу: найбільший інтервал дає той самий
                        # результат, що й найменший.
                        chosen_spacing = spacing_max
                    else:
                        low, high = 0, len(spacing_grid) - 1  # high (14.0) свідомо годиться
                        while low < high:
                            middle = (low + high) // 2
                            if measure_spacing(spacing_grid[middle]) == best_possible:
                                high = middle
                            else:
                                low = middle + 1
                        chosen_spacing = spacing_grid[low]
                        if measure_spacing(chosen_spacing) != best_possible:
                            # Підстраховка на випадок, коли верстка Word виявилась
                            # немонотонною: повертаємось до гарантованого значення.
                            chosen_spacing = spacing_min
                            measure_spacing(chosen_spacing)

                    pages_count, items_on_first = spacing_measurements[chosen_spacing]
                    if chosen_spacing != spacing_max:
                        self.log(
                            f"  ℹ️ {cipher}: точний інтервал {chosen_spacing} пт — "
                            f"сторінок: {pages_count}, пунктів на першій: {items_on_first}."
                        )

                    if pages_count > 1:
                        lower, upper = spacing_bounds()
                        blank_formats = []
                        if lower is not None and upper is not None and upper > lower:
                            for paragraph in doc.Range(lower, upper).Paragraphs:
                                paragraph_range = paragraph.Range
                                if any(start <= paragraph_range.Start < end for start, end in table_spans):
                                    continue  # таблиця зразка — не чіпаємо
                                if is_blank_paragraph(paragraph):
                                    blank_formats.append(paragraph_range.ParagraphFormat)

                        def set_blank_spacing(points):
                            for blank_format in blank_formats:
                                blank_format.LineSpacingRule = 4  # wdLineSpaceExactly
                                blank_format.LineSpacing = points

                        def pages_now():
                            doc.Repaginate()
                            return doc.ComputeStatistics(2)  # wdStatisticPages

                        squeezed = None
                        # Рівні тексту від підібраного вниз. На кожному спершу
                        # одна перевірка з найменшими порожніми рядками: не
                        # вміщується навіть так — цей рівень відкидаємо одразу.
                        # Вміщується — двійковим пошуком беремо НАЙБІЛЬШІ порожні
                        # рядки, з якими ще вміщується: стискаємо рівно настільки,
                        # наскільки треба, а не до межі.
                        for text_points in (chosen_spacing, *(
                            points for points in EXTRA_TEXT_SPACINGS if points < chosen_spacing
                        )):
                            if text_points != chosen_spacing:
                                apply_exact_line_spacing(text_points)
                            blank_grid = [text_points] + [
                                points for points in EXTRA_BLANK_SPACINGS if points < text_points
                            ] if blank_formats else [text_points]
                            set_blank_spacing(blank_grid[-1])
                            if pages_now() >= pages_count:
                                continue
                            low, high = 0, len(blank_grid) - 1  # high свідомо вміщується
                            while low < high:
                                middle = (low + high) // 2
                                set_blank_spacing(blank_grid[middle])
                                if pages_now() < pages_count:
                                    high = middle
                                else:
                                    low = middle + 1
                            set_blank_spacing(blank_grid[low])
                            if pages_now() >= pages_count:
                                # Word виявився немонотонним — гарантоване значення.
                                set_blank_spacing(blank_grid[-1])
                                pages_now()
                                low = len(blank_grid) - 1
                            squeezed = (text_points, blank_grid[low])
                            break
                        if squeezed is None:
                            # Не допомогло — повертаємо підібраний інтервал.
                            apply_exact_line_spacing(chosen_spacing)
                            pages_now()
                        else:
                            chosen_spacing = squeezed[0]
                            new_pages = doc.ComputeStatistics(2)
                            self.log(
                                f"  🗜 {cipher}: стиснуто до {new_pages} стор. замість {pages_count} — "
                                f"текст {squeezed[0]} пт, порожні рядки {squeezed[1]} пт."
                            )
                            pages_count = new_pages

                    # Виняток із правила «шрифт витягу — 14» (користувач, 29.09.2026):
                    # ЛИШЕ пункт про смерть / загибель, який і після підбору
                    # інтервалу не вміщується на одну сторінку, дозволено зменшити
                    # до 12 пт — такі пункти великі. Решта витягу лишається 14.
                    if pages_count > 1:
                        shrunk = False
                        for item_range, label in zip(inserted_item_ranges, inserted_item_labels):
                            start_page, end_page = range_pages(item_range)
                            if start_page is None or start_page == end_page:
                                continue
                            if not is_death_item_text(item_range.Text or ""):
                                continue
                            for size in DEATH_ITEM_FONT_SIZES:
                                item_range.Font.Size = size
                                item_format = item_range.ParagraphFormat
                                item_format.LineSpacingRule = 4  # wdLineSpaceExactly
                                item_format.LineSpacing = round(chosen_spacing * size / 14.0, 2)
                                doc.Repaginate()
                                start_page, end_page = range_pages(item_range)
                                if start_page == end_page:
                                    break
                            shrunk = True
                            self.log(
                                f"  🔡 {cipher}: {label} (смерть / загибель) — шрифт зменшено до {size} пт"
                                + ("." if start_page == end_page else ", але пункт усе одно довший за сторінку.")
                            )
                        if shrunk:
                            pages_count = doc.ComputeStatistics(2)

                # 2. Остаточна перевірка макета — ПІСЛЯ підбору міжрядкового
                # інтервалу (16→14 пт), який часто сам усуває розбіжність
                # (відірвану шапку, перебір на 2 сторінки тощо). Лише якщо
                # проблема лишається й після цього — витяг справді потребує
                # ручної перевірки, і виконавця НЕ підштовхуємо штучно.
                #
                # Витяг на ОДНІЙ сторінці не перевіряємо взагалі: жоден пункт не
                # може її не вміститись, ніщо не може відірватись від шапки, і
                # підписант завжди там само. А кожна така перевірка — це два
                # запити номера сторінки на кожен пункт і кожну шапку.
                if inserted_item_ranges and pages_count > 1:
                    remaining_issues = layout_issues(
                        inserted_item_ranges, inserted_item_labels, heading_item_pairs, signer_start, executor_start
                    )
                    for issue in remaining_issues:
                        layout_warnings.append(f"{cipher}: {issue}.")
                    if remaining_issues:
                        extract_needs_manual_review = True

                # 3. Виконавець ЗАВЖДИ вирівнюється строго до низу поточної сторінки,
                # окрім випадків, де це вимагає ручних змін (див. position_executor_at_page_bottom)
                position_executor_at_page_bottom(executor_bookmark, signer_start, extract_needs_manual_review)
                doc.Repaginate()
                pages_count = doc.ComputeStatistics(2)

                # Службову закладку виконавця прибираємо лише після позиціонування.
                if executor_bookmark and doc.Bookmarks.Exists(executor_bookmark):
                    try:
                        doc.Bookmarks(executor_bookmark).Delete()
                    except Exception:
                        pass

                doc.Save()
                doc.Close(False)
                (management_temp_files if is_management else temp_files).append(
                    (temp_path, pages_count, cipher)
                )
                self.log(
                    f"  ✓ {cipher}: {pages_count} стор., "
                    f"{time.monotonic() - extract_started_at:.1f} с."
                )

            source_doc.Close(False)

            # Витяги до управління збираються окремо: без вирівнювання під
            # друк «2 на 1» — їх не друкують пакетом і нікому не надсилають.
            if management_temp_files:
                self.log(
                    f"\nЗбираємо {len(management_temp_files)} витяг(ів) до управління "
                    "в окремий документ..."
                )
                management_doc = word.Documents.Open(os.path.abspath(management_temp_files[0][0]))
                for next_path, _next_pages, _next_name in management_temp_files[1:]:
                    tail_range = management_doc.Content
                    tail_range.Collapse(0)
                    tail_range.InsertBreak(2)  # wdSectionBreakNextPage
                    tail_range = management_doc.Content
                    tail_range.Collapse(0)
                    tail_range.InsertFile(os.path.abspath(next_path))
                tail_paragraph = last_paragraph(management_doc)
                if not tail_paragraph.Range.Text.strip():
                    tail_paragraph.Range.Delete()
                management_doc.SaveAs2(os.path.abspath(management_out_file), 16)
                management_doc.Close(False)
                self.log(f"Збережено витяги до управління: {management_out_file}")

            if temp_files:
                # Збираємо всі витяги в один фінальний документ
                self.progress_step(len(temp_files), "складання документа")
                self.log(f"\nЗбираємо {len(temp_files)} витягів в один документ...")
                enable_2up = self.duplex_2up_layout.get()
                if enable_2up and len(temp_files) > 1:
                    reordered = order_extracts_for_2up(temp_files)
                    if reordered != temp_files:
                        self.log(
                            "Порядок витягів змінено для економії паперу: спершу "
                            "багатосторінкові, потім одинарні — без зайвих порожніх сторінок."
                        )
                    temp_files = reordered

                first_path, first_pages, first_cipher = temp_files[0]
                if len(temp_files) == 1:
                    if enable_2up and first_pages > 1 and (first_pages % 2 != 0):
                        target_doc = word.Documents.Open(os.path.abspath(first_path))
                        rng = target_doc.Content
                        rng.Collapse(0)
                        rng.InsertBreak(2)  # wdSectionBreakNextPage
                        target_doc.SaveAs2(os.path.abspath(out_file), 16)
                        target_doc.Close(False)
                        self.log(f"Додано порожню сторінку для вирівнювання витягу {first_cipher} ({first_pages} стор.) під друк 2 на 1.")
                    else:
                        # Через SaveAs2, а не копіюванням: тимчасовий файл може мати
                        # формат шаблону (Word 97-2003), а результат завжди .docx.
                        target_doc = word.Documents.Open(os.path.abspath(first_path))
                        target_doc.SaveAs2(os.path.abspath(out_file), 16)
                        target_doc.Close(False)
                else:
                    target_doc = word.Documents.Open(os.path.abspath(first_path))

                    def sheet_pages() -> int:
                        """ФАКТИЧНА кількість сторінок зібраного документа.

                        Вимір потрібен лише для друку «2 на 1»: там від парності
                        залежить, на яку половину аркуша сяде наступний витяг. Без
                        цього режиму репагінувати документ, що росте, після кожної
                        вставки — це квадратична робота ні для чого.
                        """
                        if not enable_2up:
                            return 0
                        target_doc.Repaginate()
                        return target_doc.ComputeStatistics(2)  # wdStatisticPages

                    def add_blank_page() -> None:
                        rng_blank = target_doc.Content
                        rng_blank.Collapse(0)
                        rng_blank.InsertBreak(2)  # wdSectionBreakNextPage

                    def strip_trailing_blank_paragraphs() -> None:
                        """Прибирає «висячі» порожні абзаци в кінці документа.

                        Без цього вимірювання рахувало ФАНТОМНУ останню сторінку,
                        якої в друці немає: парність виходила невірна, і наступний
                        витяг сідав на праву половину того самого аркуша, де
                        закінчувався попередній. Викликати ЛИШЕ після вставки
                        витягу — навмисні порожні сторінки чіпати не можна.
                        """
                        try:
                            while target_doc.Paragraphs.Count > 1:
                                paragraph = last_paragraph(target_doc).Range
                                if (paragraph.Text or "").strip(chr(13) + chr(7) + chr(11) + chr(12) + chr(32) + chr(9)):
                                    break
                                # Той самий захист, що й у clean_redundant_blanks:
                                # знак абзацу, який Word видалити не може, мовчки
                                # лишається на місці й зациклює обрізання хвоста.
                                length_before = target_doc.Content.End
                                paragraph.Delete()
                                if target_doc.Content.End >= length_before:
                                    break
                        except Exception:
                            pass

                    # Сторінки ЗАВЖДИ міряємо, а не рахуємо додаванням: розрив
                    # розділу не гарантовано додає рівно одну сторінку, тож
                    # лічильник «повзе». Через це багатосторінковий витяг міг
                    # опинитися на ПРАВІЙ половині аркуша, а кінець попереднього
                    # витягу — ділити аркуш із початком наступного.
                    strip_trailing_blank_paragraphs()
                    current_doc_pages = sheet_pages()

                    if enable_2up and current_doc_pages > 1 and (current_doc_pages % 2 != 0):
                        add_blank_page()
                        current_doc_pages = sheet_pages()
                        self.log(
                            f"Додано порожню сторінку після витягу {first_cipher} "
                            f"({current_doc_pages - 1} стор.) для вирівнювання аркуша."
                        )

                    for i in range(1, len(temp_files)):
                        t_path, t_pages, t_cipher = temp_files[i]

                        # Багатосторінковий витяг починається з ЛІВОЇ половини
                        # аркуша, тобто з НЕПАРНОЇ логічної сторінки.
                        if enable_2up and t_pages > 1 and (current_doc_pages % 2 != 0):
                            add_blank_page()
                            current_doc_pages = sheet_pages()
                            self.log(
                                f"Додано порожню сторінку перед багатосторінковим витягом "
                                f"{t_cipher} ({t_pages} стор.), щоб він почався з нового аркуша."
                            )

                        pages_before = current_doc_pages

                        rng = target_doc.Content
                        rng.Collapse(0)
                        rng.InsertBreak(2)  # wdSectionBreakNextPage

                        rng = target_doc.Content
                        rng.Collapse(0)
                        rng.InsertFile(os.path.abspath(t_path))
                        strip_trailing_blank_paragraphs()
                        current_doc_pages = sheet_pages()

                        # Скільки сторінок витяг займає САМЕ В ЗІБРАНОМУ документі.
                        # Окремим файлом він міг мати іншу кількість, і саме довіра
                        # до тієї, старої, ламала вирівнювання. Без режиму «2 на 1»
                        # зібраний документ не міряємо, тож беремо кількість файлу.
                        actual_pages = max(1, current_doc_pages - pages_before) if enable_2up else t_pages
                        if enable_2up and actual_pages != t_pages:
                            self.log(
                                f"УВАГА: витяг {t_cipher} у зібраному документі займає "
                                f"{actual_pages} стор. замість {t_pages}; вирівнювання "
                                "рахується за фактичною кількістю."
                            )

                        # Після багатосторінкового витягу наступний має починатися з
                        # НОВОГО аркуша: якщо документ закінчився на лівій половині
                        # (непарна сторінка) — доповнюємо порожньою.
                        if enable_2up and actual_pages > 1 and (current_doc_pages % 2 != 0):
                            add_blank_page()
                            current_doc_pages = sheet_pages()
                            self.log(
                                f"Додано порожню сторінку після багатосторінкового витягу "
                                f"{t_cipher} ({actual_pages} стор.) для вирівнювання наступного аркуша."
                            )

                    last_para = last_paragraph(target_doc)
                    if last_para.Range.Text.strip() == "":
                        last_para.Range.Delete()

                    target_doc.SaveAs2(os.path.abspath(out_file), 16)
                    target_doc.Close(False)

                # Точний підрахунок сторінок та фізичних аркушів паперу для друку
                check_doc = word.Documents.Open(os.path.abspath(out_file), ReadOnly=True)
                check_doc.Repaginate()
                final_total_pages = check_doc.ComputeStatistics(2)
                check_doc.Close(False)

                total_extracts = len(temp_files)
                single_page_cnt = sum(1 for tf in temp_files if tf[1] == 1)
                multi_page_cnt = sum(1 for tf in temp_files if tf[1] > 1)

                if enable_2up:
                    sheets_needed = (final_total_pages + 1) // 2
                    stats_msg = (
                        f"📋 Сформовано витягів: {total_extracts} шт. ({single_page_cnt} односторінк. та {multi_page_cnt} багатосторінк.)\n"
                        f"📄 Загальна кількість сторінок у документі: {final_total_pages} стор.\n"
                        f"🖨️ Кількість фізичних аркушів паперу для друку («2 сторінки на 1 аркуш»): {sheets_needed} арк. А4"
                    )
                else:
                    sheets_1side = final_total_pages
                    sheets_2side = (final_total_pages + 1) // 2
                    stats_msg = (
                        f"📋 Сформовано витягів: {total_extracts} шт. ({single_page_cnt} односторінк. та {multi_page_cnt} багатосторінк.)\n"
                        f"📄 Загальна кількість сторінок у документі: {final_total_pages} стор.\n"
                        f"🖨️ Кількість фізичних аркушів паперу для друку:\n"
                        f"   • Односторонній друк (1 на 1): {sheets_1side} арк. А4\n"
                        f"   • Двосторонній друк (Duplex): {sheets_2side} арк. А4"
                    )

                # Кількість пропущених пунктів показується ЗАВЖДИ, зокрема «0».
                # Інакше пункт, який не отримав адресата, лишався помітним лише
                # в журналі, і його легко було не побачити серед статистики друку.
                missed_count = len(unmatched_items)
                excluded_count = len(map_res.get("skipped_items", []))
                missed_lines = [
                    f"{'⚠️' if missed_count else '✅'} Пунктів без адресата (пропущено): {missed_count}"
                ]
                if missed_count:
                    missed_lines.append(f"   Перелік: {os.path.basename(missing_report)}")
                if excluded_count:
                    missed_lines.append(
                        f"ℹ️ Виключено із загального переліку (управління): {excluded_count}"
                    )
                if management_temp_files:
                    missed_lines.append(
                        f"📁 Витяги до управління — окремий файл на {len(management_temp_files)} шт.: "
                        f"{os.path.basename(management_out_file)}"
                    )
                stats_msg = chr(10).join([stats_msg] + missed_lines)

                batch_seconds = time.monotonic() - batch_started_at
                self.progress_end(f"Витягів: {total_extracts}")
                self.log(f"Збережено файл: {out_file}")
                self.log(
                    f"⏱️ Час генерації: {batch_seconds:.1f} с "
                    f"({batch_seconds / max(1, total_extracts):.1f} с на витяг)."
                )
                self.log(f"\n📊 СТАТИСТИКА ТА РОЗРАХУНОК ДРУКУ:\n{stats_msg}\n")
                self.show_layout_warnings(layout_warnings)

                if layout_warnings:
                    self.log("\nУВАГА: потрібна ручна перевірка макета:")
                    for warning in layout_warnings:
                        self.log(f"• {warning}")
                    messagebox.showwarning(
                        "Перевірте макет витягів",
                        f"Сформовано {total_extracts} витягів, але деякі витяги потребують ручного коригування.\n\n"
                        f"{stats_msg}\n\nДеталі у журналі.",
                    )
                else:
                    messagebox.showinfo(
                        "Успіх",
                        f"Успішно сформовано {total_extracts} витягів!\n\n"
                        f"{stats_msg}\n\n"
                        f"Файл збережено:\n{os.path.basename(out_file)}"
                    )
            else:
                # У наказі лише зміни по управлінню: загального файла немає,
                # і це не помилка — розрахунок розсилки для них не ведеться.
                total_extracts = 0
                self.progress_end(f"До управління: {len(management_temp_files)}")
                self.log("Загальний файл адресних витягів у цьому запуску не створювався.")
                messagebox.showinfo(
                    "Успіх",
                    f"Сформовано витягів до управління: {len(management_temp_files)}.\n\n"
                    "Загальний файл адресних витягів у цьому запуску не створювався.\n\n"
                    f"Файл збережено:\n{os.path.basename(management_out_file)}",
                )

            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir, ignore_errors=True)

            # Налаштування повертаємо ДО показу вікна: з вимкненим ScreenUpdating
            # видимий Word не перемальовується й виглядає зависшим.
            restore_word_settings(word, word_settings)
            word_settings = None

            self.log("Відкриваємо згенерований документ...")
            # Якщо загального файла немає (у наказі лише зміни по управлінню),
            # показуємо той єдиний документ, який справді створено.
            final_doc = word.Documents.Open(
                os.path.abspath(out_file if temp_files else management_out_file)
            )
            word.Visible = True
            final_doc.Activate()
            word = None

        finally:
            self.progress_end()
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir, ignore_errors=True)
            if word:
                restore_word_settings(word, word_settings)
                force_quit_word(word)
