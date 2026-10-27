"""Примірники 2/3.

Винесено з `generate_extracts.py` без змін логіки (26.10.2026).
"""
from __future__ import annotations

import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.constants import DISABLED, NORMAL
import win32com.client
from nodeautomationtoolkit.builtin_nodes.copy_generator import (
    PreviewSteps,
    build_copy_document,
    retry_on_busy_word,
)
from nodeautomationtoolkit.builtin_nodes.template_tags import (
    SIGNER_TAGS, certifier_tags,
)

from .errors import explain_error
from .naming import _slash_to_lines, back_page_tag_values, build_copy_two_filename, extract_metadata_from_filename, sanitize_filename
from .printing import copy_sheet_count
from .word_doc import copy_template_for_editing, force_quit_word


class CopiesMixin:
    """Примірники 2/3."""

    def _on_p2_source_mode_changed(self):
        """Сумісність із Drag-and-Drop: після зміни джерела оновлюємо список."""
        self._refresh_p2_orders_view()

    def show_p2_info(self):
        messagebox.showinfo(
            "Правила формування примірника № 2",
            "Примірники № 2/3 збираються з DOCX-заготовки. У {{зміст}} переноситься "
            "тіло наказу, а службова остання сторінка лишається із заготовки.\n\n"
            "Основні теги заготовки:\n"
            "• {{зміст}} — тіло наказу (обов’язковий тег)\n"
            "• {{підписант}} або окремі теги {{підписант_посада}}, "
            "{{підписант_звання}}, {{підписант_імя}}\n"
            "• {{згідно_з_оригіналом}} або {{засвідчення}}\n"
            "• {{номер_наказу}} або {{номер}} — номер із назви файла\n"
            "• {{дата_наказу}} або {{дата}} — дата з назви файла\n"
            "• {{примірник}} або {{примірник_номер}} — номер примірника\n\n"
            "Якщо номер або дату не знайдено в назві файла, відповідний тег лишається "
            "для ручного заповнення. Заготовка може бути багатосторінковою.",
        )

    def select_p2_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.p2_orders_folder.set(path)
            self.p2_source_mode.set("folder")
            # Папку результату переобчислюємо під нове джерело, доки
            # користувач не вибрав її вручну.
            if not self.p2_out_folder_manual.get():
                self.p2_out_folder.set(os.path.join(path, "Примірники_2"))
            self._refresh_p2_orders_view()
            self.save_config()

    def select_p2_file(self):
        paths = filedialog.askopenfilenames(filetypes=[("Word files", "*.docx *.doc")])
        if paths:
            self._set_p2_orders(list(paths))
            self.log_p2(f"Обрано наказів для примірників: {len(self.p2_manual_order_paths)}")

    def select_p2_back_page(self):
        path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx *.doc")])
        if path:
            self.p2_back_page_path.set(path)
            self.save_config()

    def edit_p2_back_page(self):
        path = self.p2_back_page_path.get()
        if not path or not os.path.exists(path):
            messagebox.showwarning("Помилка", "Спочатку виберіть файл шаблону задньої сторінки!")
            return
        try:
            self.log_p2("Відкриваємо задню сторінку у Word для редагування...")
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
            self.log_p2(f"Помилка відкриття задньої сторінки: {str(e)}")

    def select_p2_out_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.p2_out_folder.set(path)
            # Явний вибір користувача — більше не переобчислюємо автоматично.
            self.p2_out_folder_manual.set(True)
            self.save_config()

    def _set_copy_two_sources(self, paths: list[str], selected: bool = True):
        """Показує примірники № 2 останнього проходу як позначуваний список."""
        self.last_copy_two_paths = [os.path.abspath(path) for path in paths if os.path.isfile(path)]
        for item in self.copy_two_tree.get_children():
            self.copy_two_tree.delete(item)
        if not self.last_copy_two_paths:
            self.copy_two_tree.insert(
                "", tk.END,
                values=("—", "Спершу сформуйте примірники № 2 у вкладці «Примірники 2/3»."),
            )
            self._refresh_source_summary()
            return
        mark = "☑" if selected else "☐"
        for index, path in enumerate(self.last_copy_two_paths):
            self.copy_two_tree.insert("", tk.END, iid=f"copy2_{index}", values=(mark, path))
        self._select_order_path(self.copy_two_tree, self.doc_path.get())
        self._refresh_source_summary()

    def _refresh_p2_orders_view(self):
        """Список наказів для примірників: окремі файли або вміст теки."""
        tree = getattr(self, "p2_orders_tree", None)
        if tree is None:
            return
        if self.p2_source_mode.get() == "folder":
            paths = self._orders_in_folder(self.p2_orders_folder.get())
        else:
            paths = self._p2_manual_paths()
        self._fill_orders_tree(
            tree, paths,
            "Накази не обрано — натисніть «Обрати накази» або перетягніть файли сюди.",
        )
        self._refresh_p2_source_summary()

    def _refresh_p2_source_summary(self):
        if not hasattr(self, "p2_source_summary"):
            return
        paths = self._selected_p2_order_paths()
        if not paths:
            self.p2_source_summary.set("Накази не обрано")
            return
        source = (
            f"тека {os.path.basename(self.p2_orders_folder.get().rstrip(os.sep))}"
            if self.p2_source_mode.get() == "folder"
            else "обрані файли"
        )
        self.p2_source_summary.set(f"Наказів для примірників: {len(paths)} ({source})")

    def _p2_manual_paths(self) -> list[str]:
        """Окремо обрані накази для примірників, що досі існують на диску."""
        manual = [path for path in self.p2_manual_order_paths if os.path.isfile(path)]
        single = self.p2_single_file.get()
        if not manual and single and os.path.isfile(single):
            # Старий «окремий файл» із конфігу, записаного до появи списку.
            manual = [os.path.abspath(single)]
        return manual

    def _selected_p2_order_paths(self) -> list[str]:
        """Накази, позначені галочкою на вкладці примірників."""
        tree = getattr(self, "p2_orders_tree", None)
        if tree is not None:
            marked = self._marked_paths(tree)
            if marked:
                return marked
        if self.p2_source_mode.get() == "folder":
            return self._orders_in_folder(self.p2_orders_folder.get())
        return self._p2_manual_paths()

    def _set_p2_orders(self, paths: list[str]):
        self.p2_manual_order_paths = [
            os.path.abspath(path) for path in paths if os.path.isfile(path)
        ]
        self.p2_source_mode.set("file")
        if self.p2_manual_order_paths:
            self.p2_single_file.set(self.p2_manual_order_paths[0])
            if not self.p2_out_folder_manual.get():
                self.p2_out_folder.set(
                    os.path.join(os.path.dirname(self.p2_manual_order_paths[0]), "Примірники_2")
                )
        self._refresh_p2_orders_view()
        self.save_config()

    def clear_p2_orders(self):
        self.p2_manual_order_paths = []
        self.p2_single_file.set("")
        self._refresh_p2_orders_view()
        self.save_config()

    # =========================================================================
    # ДІЇ: ГЕНЕРАЦІЯ ПРИМІРНИКІВ 2/3
    # =========================================================================
    def run_generate_copies(self):
        self.save_config()
        back_page = self.p2_back_page_path.get()
        if not back_page or not os.path.exists(back_page):
            messagebox.showwarning("Помилка", "Виберіть дійсний файл шаблону «Задня сторінка»!")
            return

        # Беремо рівно те, що видно в списку на вкладці: і тека, і окремо
        # обрані файли проходять той самий шлях, тож пакет завжди збігається
        # з тим, що бачить користувач.
        order_files = [
            path for path in self._selected_p2_order_paths()
            if os.path.normcase(os.path.abspath(path)) != os.path.normcase(os.path.abspath(back_page))
        ]
        if not order_files:
            messagebox.showwarning("Немає файлів", "Не знайдено файлів наказів DOCX для обробки!")
            return

        copy_title = "Примірник № 2"
        back_page_abs = os.path.abspath(back_page)

        out_root = self.p2_out_folder.get()
        if not out_root or not self.p2_out_folder_manual.get():
            # Поки папку не обрано вручну, вона завжди відповідає ПОТОЧНОМУ
            # джерелу: інакше примірники писались би в папку попереднього наказу.
            base_dir = (
                self.p2_orders_folder.get()
                if self.p2_source_mode.get() == "folder"
                else os.path.dirname(order_files[0])
            )
            out_root = os.path.join(base_dir, "Примірники_2")
            self.p2_out_folder.set(out_root)
        os.makedirs(out_root, exist_ok=True)
        self.save_config()

        self.btn_run_p2.config(state=DISABLED)
        self.p2_log_text.delete(1.0, tk.END)
        for item in self.p2_tree.get_children():
            self.p2_tree.delete(item)

        self.log_p2(f"Початок пакетного формування примірників № 2: {len(order_files)} наказ(ів)...")

        import traceback

        created_records = []
        failed_orders: list[tuple[str, str]] = []
        # У режимі превʼю Word показується — інакше дивитись нема на що.
        preview_on = bool(self.p2_preview.get())
        preview_delay = self.p2_preview_delay.get()
        if preview_on:
            self.log_p2(
                f"🐢 Режим превʼю увімкнено: Word буде видимим, пауза після "
                f"кожного кроку — {preview_delay} сек. Для великого пакета це довго."
            )

        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = preview_on
        word.DisplayAlerts = 0  # wdAlertsNone: не показувати блокуючі діалоги (напр. "Зберегти зміни?")

        try:
            # Заготовка примірника може бути багатосторінковою.
            test_tmpl = word.Documents.Open(back_page_abs, ReadOnly=True)
            try:
                test_tmpl.Repaginate()
                tmpl_pages = test_tmpl.ComputeStatistics(2)
                template_text = test_tmpl.Content.Text
            finally:
                test_tmpl.Close(False)

            # Заготовка більше не є «останньою сторінкою», яку підставляють у
            # наказ: тепер це повноцінний шаблон примірника (шапка + теги +
            # {{зміст}}), тому обмеження в 1 сторінку зняте.
            if tmpl_pages != 1:
                self.log_p2(f"  ℹ️ Заготовка займає {tmpl_pages} стор.")

            if "{{зміст}}" not in template_text:
                self.log_p2("❌ ПОМИЛКА: у заготовці примірника не знайдено тег {{зміст}}!")
                messagebox.showerror(
                    "Помилка шаблону",
                    "У заготовці примірника не знайдено тег {{зміст}}.\n\n"
                    f"Файл: {os.path.basename(back_page_abs)}\n\n"
                    "Саме в це місце вставляються пункти наказу разом із підписантом.",
                )
                return

            # Якщо у заготовці є окремий тег підписанта — блок підписанта йде
            # в нього, а зміст завершується перед ним. Інакше підписант
            # лишається частиною змісту (сумісність зі старими заготовками).
            signer_tag_in_template = any(tag in template_text.casefold() for tag in SIGNER_TAGS)
            self.log_p2(
                "  Підписант: окремий тег {{підписант}}."
                if signer_tag_in_template
                else "  Підписант: у складі {{зміст}} (тегу {{підписант}} у заготовці немає)."
            )

            required_tags = tuple(back_page_tag_values("номер", "01.01.2000"))
            missing_tags = [tag for tag in required_tags if tag not in template_text]
            if missing_tags:
                self.log_p2(
                    "УВАГА: у шаблоні не знайдено тегів: " + ", ".join(missing_tags) + ". "
                    "Їхні значення не буде підставлено автоматично."
                )

            for idx, order_path in enumerate(order_files, start=1):
                fname = os.path.basename(order_path)
                self.log_p2(f"\n[{idx}/{len(order_files)}] Обробка: {fname}")
                order_num, order_date = extract_metadata_from_filename(fname)

                # Номер наказу може містити «/» (напр. «б/н», «123/45»), який
                # ОС трактує як роздільник шляху й створює вкладені папки,
                # тому назву папки треба чистити так само, як назву файлу.
                sub_folder_name = (
                    f"Наказ № {sanitize_filename(order_num)}"
                    if order_num
                    else sanitize_filename(os.path.splitext(fname)[0])
                )
                target_dir = os.path.join(out_root, sub_folder_name)
                os.makedirs(target_dir, exist_ok=True)

                out_copy_name = build_copy_two_filename(order_num, order_date, fname)
                target_file = os.path.join(target_dir, out_copy_name)

                doc = None
                working_copy = ""
                try:
                    self.log_p2("  Режим: збірка примірника із заготовки (без колонтитулів).")
                    values = back_page_tag_values(order_num, order_date)
                    # Зворотна сумісність для вже створених шаблонів.
                    if order_num:
                        values["{{номер}}"] = order_num
                    if order_date:
                        values["{{дата}}"] = order_date
                    values["{{примірник_номер}}"] = copy_title
                    values["{{засвідчення}}"] = "Згідно з оригіналом"

                    # Виконавець примірників; якщо не заповнено — беремо
                    # виконавця витягів, щоб не змушувати вводити двічі.
                    executor_value = (
                        self.p2_executor.get().strip() or self.executor.get().strip()
                    )
                    if executor_value:
                        values["{{виконавець}}"] = _slash_to_lines(executor_value)

                    # Засвідчувач примірників — окремі поля. За потреби
                    # користувач переносить значення з витягів окремою кнопкою.
                    cert_pos = _slash_to_lines(self.p2_certifier_position.get().strip())
                    cert_rank = self.p2_certifier_rank.get().strip()
                    cert_name = self.p2_certifier_name.get().strip()
                    values.update(certifier_tags(cert_pos, cert_rank, cert_name))
                    if cert_pos:
                        values["{{засвідчувач_посада}}"] = cert_pos
                        values["{{згідно_з_оригіналом_посада}}"] = cert_pos
                    if cert_rank:
                        values["{{засвідчувач_звання}}"] = cert_rank
                        values["{{згідно_з_оригіналом_звання}}"] = cert_rank
                    if cert_name:
                        values["{{засвідчувач_піб}}"] = cert_name
                        values["{{згідно_з_оригіналом_піб}}"] = cert_name
                        values["{{засвідчувач}}"] = cert_name

                    # Тег «Згідно з оригіналом» розгортається в СУЦІЛЬНИЙ блок
                    # без порожніх абзаців: заголовок, посада, а останнім
                    # рядком — звання та прізвище.
                    certifier_block = ["Згідно з оригіналом"]
                    if cert_pos:
                        certifier_block.append(cert_pos)
                    signature_line = " ".join(part for part in (cert_rank, cert_name) if part)
                    if signature_line:
                        certifier_block.append(signature_line)
                    values["{{згідно_з_оригіналом}}"] = "\r".join(certifier_block)
                    values["{{засвідчення}}"] = values["{{згідно_з_оригіналом}}"]

                    # Заготовка може бути у форматі Word 97-2003, тож робоча
                    # копія зберігає реальне розширення, а .docx дає SaveAs2.
                    working_copy = copy_template_for_editing(
                        back_page_abs,
                        os.path.join(target_dir, "_nat_tmpl.docx"),
                        label="заготовка примірника",
                    )
                    # Під кінець великого пакета Word може тимчасово відхиляти
                    # виклики — такий збій не є помилкою даних, тому повторюємо.
                    # Показник кроків СВІЙ на кожен наказ: інакше нумерація
                    # «Крок N/8» тривала б наскрізно через увесь пакет.
                    steps = PreviewSteps(
                        log=self.log_p2,
                        delay=preview_delay,
                        enabled=preview_on,
                        sleeper=self._preview_pause,
                    )
                    final_pages = retry_on_busy_word(
                        lambda: build_copy_document(
                            word,
                            os.path.abspath(order_path),
                            os.path.abspath(working_copy),
                            os.path.abspath(target_file),
                            values,
                            resolve_span=lambda source: self._order_body_context(
                                source, signer_as_tag=signer_tag_in_template
                            ),
                            log=self.log_p2,
                            preview=steps,
                        ),
                        log=self.log_p2,
                    )

                    # Друк односторонній, лише задня сторінка — на звороті (26.10.2026).
                    sheets_1_copy = copy_sheet_count(final_pages)
                    self.log_p2(
                        f"  ✅ Створено: {out_copy_name} (сторінок: {final_pages} | "
                        f"аркушів на примірник: {sheets_1_copy}, на № 2 і № 3: {sheets_1_copy * 2})"
                    )
                    created_records.append((
                        idx,
                        out_copy_name,
                        order_num or "—",
                        order_date or "—",
                        final_pages,
                        sheets_1_copy,
                        target_file,
                    ))
                except Exception as order_error:
                    # Збій одного наказу не має зривати весь пакет.
                    traceback.print_exc()
                    explanation = explain_error(order_error)
                    self.log_p2(f"  ПОМИЛКА під час обробки «{fname}»:\n    "
                                + explanation.replace("\n", "\n    "))
                    failed_orders.append((fname, explanation))
                    try:
                        if doc is not None:
                            doc.Close(False)
                    except Exception:
                        pass
                finally:
                    if working_copy and os.path.exists(working_copy):
                        try:
                            os.remove(working_copy)
                        except OSError:
                            pass

            for rec in created_records:
                self.p2_tree.insert("", tk.END, values=rec)
            self._set_copy_two_sources([record[-1] for record in created_records])

            total_orders = len(created_records)
            total_pages = sum(r[4] for r in created_records if isinstance(r[4], int))
            total_sheets_1_copy = sum(r[5] for r in created_records if isinstance(r[5], int))

            p2_stats_msg = (
                f"📑 Кількість опрацьованих наказів: {total_orders} шт.\n"
                f"📄 Сторінок в 1 примірнику (сумарно): {total_pages} стор.\n"
                f"🖨️ Аркушів А4 (друк односторонній, задня сторінка — на звороті передостанньої):\n"
                f"   • на примірники № 2: {total_sheets_1_copy} арк.\n"
                f"   • на примірники № 2 і № 3 разом: {total_sheets_1_copy * 2} арк."
            )

            self.log_p2(f"\n📊 СТАТИСТИКА ТА РОЗРАХУНОК ДРУКУ:\n{p2_stats_msg}")
            self.log_p2(f"\n🎉 Завершено! Успішно сформовано {total_orders} примірник(ів) № 2.")

            if failed_orders:
                details = "\n\n".join(
                    f"Наказ «{name}» — примірник не створено.\n{error}"
                    for name, error in failed_orders[:10]
                )
                if len(failed_orders) > 10:
                    details += (
                        f"\n\n… і ще {len(failed_orders) - 10} наказ(ів) "
                        "з такими самими збоями."
                    )
                self.log_p2(f"\nНе вдалося сформувати примірники: {len(failed_orders)} шт.")
                messagebox.showwarning(
                    "Примірники сформовано частково",
                    f"Сформовано {total_orders} з {len(order_files)} примірник(ів).\n\n"
                    f"Не вдалося ({len(failed_orders)}):\n{details}\n\n{p2_stats_msg}",
                )
            else:
                messagebox.showinfo(
                    "Успіх",
                    f"Успішно сформовано {total_orders} примірник(ів) № 2!\n\n"
                    f"{p2_stats_msg}"
                )

        except Exception as error:
            # Раніше винятку не було де перехопити: таблиця лишалась порожньою,
            # а користувач бачив лише traceback у консолі.
            traceback.print_exc()
            explanation = explain_error(error)
            self.log_p2(f"\n❌ ПОМИЛКА пакетної генерації:\n{explanation}")
            for rec in created_records:
                self.p2_tree.insert("", tk.END, values=rec)
            if created_records:
                self._set_copy_two_sources([record[-1] for record in created_records])
            messagebox.showerror(
                "Примірники сформовано не до кінця",
                f"{explanation}\n\n"
                f"Встигли сформувати: {len(created_records)} примірник(ів).",
            )
        finally:
            if word:
                force_quit_word(word)
            self.btn_run_p2.config(state=NORMAL)

    def select_all_copies(self):
        for item in self.p2_tree.get_children():
            self.p2_tree.selection_add(item)

    def deselect_all_copies(self):
        for item in self.p2_tree.selection():
            self.p2_tree.selection_remove(item)

    def transfer_selected_copy_to_extracts(self):
        selected = self.p2_tree.selection()
        if not selected:
            messagebox.showinfo("Вибір", "Виберіть сформований примірник у таблиці!")
            return

        valid_paths = []
        for sel_id in selected:
            item_vals = self.p2_tree.item(sel_id, "values")
            if item_vals and len(item_vals) >= 2:
                target_path = item_vals[-1]
                if os.path.exists(target_path):
                    valid_paths.append(target_path)

        if not valid_paths:
            messagebox.showwarning("Помилка", "Жоден із вибраних файлів не знайдено на диску.")
            return

        self._set_copy_two_sources(valid_paths)
        target_path = valid_paths[0]
        self._set_processing_order(target_path)
        self._refresh_order_signer()
        self.save_config()
        self.notebook.select(self.tab_extracts_page)
        self.log(f"\n📥 Передано {len(valid_paths)} примірників № 2 для розрахунку та витягів.")
        messagebox.showinfo(
            "Передано до Витягів",
            f"Позначено {len(valid_paths)} примірників № 2. У вкладці 1 можна зняти позначки з непотрібних файлів."
        )

    def open_p2_output_folder(self):
        folder = self.p2_out_folder.get()
        if folder and os.path.exists(folder):
            os.startfile(folder)
        else:
            messagebox.showwarning("Помилка", "Папка результату ще не створена.")

    def open_selected_copy_in_word(self):
        selected = self.p2_tree.selection()
        if not selected:
            messagebox.showinfo("Вибір", "Виберіть файл у таблиці!")
            return
        item_vals = self.p2_tree.item(selected[0], "values")
        if item_vals and len(item_vals) >= 2:
            file_path = item_vals[-1]
            if os.path.exists(file_path):
                os.startfile(file_path)
            else:
                messagebox.showwarning("Помилка", f"Файл не знайдено: {file_path}")
