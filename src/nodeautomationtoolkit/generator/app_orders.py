"""Список наказів, вибір, читання тексту з Word.

Винесено з `generate_extracts.py` без змін логіки (26.10.2026).
"""
from __future__ import annotations

import os
import tkinter as tk
from tkinter import filedialog, messagebox
import win32com.client

from .naming import is_generated_copy_filename
from .order_text import text_before_order_signer
from .reports import load_saved_analysis_reports
from .word_doc import force_quit_word, read_document_text


class OrdersMixin:
    """Список наказів, вибір, читання тексту з Word."""

    @staticmethod
    def _word_cache_key(doc_path: str, normalize: bool = True):
        try:
            stat = os.stat(doc_path)
        except OSError:
            return None
        return (os.path.abspath(doc_path), stat.st_mtime_ns, stat.st_size, normalize)

    def _word_bio_lines(self, doc_path: str, normalize: bool = True) -> frozenset:
        """Рядки біографії (відступ від 4 см) з останнього `_read_word_text`.

        Word тут не запускається: береться те, що зібрало саме читання тексту.
        Немає (файл змінився, текст підмінено в тестах) — порожньо, і
        маршрутизація спирається лише на текстове правило.
        """
        key = self._word_cache_key(doc_path, normalize)
        return getattr(self, "_word_bio_cache", {}).get(key, frozenset())

    def _read_word_text(self, doc_path: str, normalize: bool = True) -> str:
        """Текст наказу через Word, з кешем за файлом і часом його зміни.

        Запуск окремого екземпляра Word коштує близько секунди, а той самий наказ
        читається кілька разів за сеанс: розрахунок розсилки, оновлення підписанта,
        потім генерація витягів. Поки файл не змінився, вдруге Word не запускаємо.
        `normalize=False` — сирий текст для перевірки наказу (`read_document_text`).
        """
        cache = getattr(self, "_word_text_cache", None)
        if cache is None:
            cache = self._word_text_cache = {}
        bio_cache = getattr(self, "_word_bio_cache", None)
        if bio_cache is None:
            bio_cache = self._word_bio_cache = {}
        cache_key = self._word_cache_key(doc_path, normalize)
        if cache_key is not None and cache_key in cache:
            return cache[cache_key]

        word = None
        doc = None
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0  # wdAlertsNone: не показувати блокуючі діалоги (напр. "Зберегти зміни?")
            doc = word.Documents.Open(os.path.abspath(doc_path), ReadOnly=True)
            bio_lines: set[int] = set()
            text = read_document_text(doc, normalize=normalize, bio_lines=bio_lines)
            if cache_key is not None:
                # Кеш свідомо маленький: тексти наказів важкі, а користь дає
                # лише останній оброблюваний файл (і сусідні в пакеті).
                if len(cache) >= 4:
                    cache.clear()
                    bio_cache.clear()
                cache[cache_key] = text
                bio_cache[cache_key] = frozenset(bio_lines)
            return text
        finally:
            try:
                if doc:
                    doc.Close(False)
            except Exception:
                pass
            if word:
                force_quit_word(word)

    def _refresh_order_signer(self, text: str | None = None) -> tuple[str, dict[str, str]]:
        if text is None:
            try:
                text = self._read_word_text(self.doc_path.get())
            except Exception as error:
                self.log(f"УВАГА: не вдалося зчитати підписанта наказу: {error}")
                text = ""
        usable_text, signer = text_before_order_signer(text)
        self.order_signer_position.set(signer["position"].replace("\n", " / "))
        self.order_signer_rank.set(signer["rank"])
        self.order_signer_name.set(signer["name"])
        if signer["position"] or signer["rank"] or signer["name"]:
            self.log("Підписанта наказу зчитано; текст після нього буде проігноровано.")
        else:
            self.log("УВАГА: підписанта наказу не розпізнано; текст наказу не обрізано.")
        return usable_text, signer

    @staticmethod
    def _marked_paths(tree) -> list[str]:
        """Позначені галочкою й досі наявні файли зі списку джерел."""
        return [
            values[1]
            for item in tree.get_children()
            if (values := tree.item(item, "values"))
            and len(values) == 2 and values[0] == "☑" and os.path.isfile(values[1])
        ]

    @staticmethod
    def _orders_in_folder(folder: str) -> list[str]:
        """Накази DOCX у теці, без службових і вже згенерованих файлів."""
        if not folder or not os.path.isdir(folder):
            return []
        return [
            os.path.join(folder, name)
            for name in sorted(os.listdir(folder))
            if name.lower().endswith(".docx")
            and not name.startswith("~$")
            and not is_generated_copy_filename(name)
        ]

    def _fill_orders_tree(self, tree, paths: list[str], empty_hint: str):
        for item in tree.get_children():
            tree.delete(item)
        if not paths:
            tree.insert("", tk.END, values=("—", empty_hint))
            return
        for index, path in enumerate(paths):
            tree.insert("", tk.END, iid=f"order_{index}", values=("☑", path))

    @staticmethod
    def _select_order_path(tree, preferred_path: str = ""):
        """Позначає активний рядок окремо від галочки пакетної обробки."""
        children = tree.get_children()
        if not children:
            return
        preferred = os.path.normcase(os.path.abspath(preferred_path)) if preferred_path else ""
        chosen = ""
        for item in children:
            values = tree.item(item, "values")
            if len(values) < 2 or not os.path.isfile(values[1]):
                continue
            if not chosen:
                chosen = item
            if preferred and os.path.normcase(os.path.abspath(values[1])) == preferred:
                chosen = item
                break
        if chosen:
            tree.selection_set(chosen)
            tree.see(chosen)

    def _activate_order_from_tree(self, tree, alongside: bool = False):
        """Перемикає активний наказ і показує його збережені контрольні XLSX."""
        if self._batch_running or getattr(self, "_busy_depth", 0):
            return
        selected = tree.selection()
        if not selected:
            return
        values = tree.item(selected[0], "values")
        if len(values) < 2 or not os.path.isfile(values[1]):
            return
        order_path = os.path.abspath(values[1])
        output_folder = (
            os.path.dirname(order_path)
            if alongside
            else os.path.join(os.path.dirname(order_path), "Extracts_Output")
        )
        self.doc_path.set(order_path)
        self.out_folder.set(output_folder)
        self._refresh_source_summary()
        self._load_saved_analysis_into_ui(order_path, output_folder)

    def _load_saved_analysis_into_ui(self, order_path: str, output_folder: str):
        """Заповнює три вкладки з диска; Word і маршрутизатор не запускаються."""
        if not hasattr(self, "result_views"):
            return
        saved = load_saved_analysis_reports(order_path, output_folder)
        placeholders = {
            "calculation": ("— Для цього наказу ще немає збереженого розрахунку —", "—", "—"),
            "unmatched": ("—", "— Збережений контроль пропущених ще не створено —", "—"),
            "routing": ("—", "—", "— Збережений контроль маршрутизації ще не створено —", "—"),
        }
        columns = {
            "calculation": ["Військова частина / Відправник", "Пункти витягу", "Кількість"],
            "unmatched": ["Пункт", "Текст пункту", "Причина"],
            "routing": ["Пункт", "Збіги з таблиці", "Застосовані правила", "Підсумкові адресати"],
        }
        for key in ("calculation", "unmatched", "routing"):
            rows = saved[key]
            if not rows:
                if saved["found"][key] and key == "unmatched":
                    rows = [("—", "Пропущених пунктів немає.", "—")]
                else:
                    rows = [placeholders[key]]
            self._populate_result_tab(key, columns[key], rows)

        found_count = sum(saved["found"].values())
        name = os.path.basename(order_path)
        if found_count:
            self.log(
                f"Підвантажено збережені результати для {name}: "
                f"{found_count}/3 контрольних файлів; повторний розрахунок не запускався."
            )
        else:
            self.log(f"Для {name} збережених результатів ще немає.")
        for error in saved["errors"]:
            self.log(f"УВАГА: не вдалося відкрити збережений результат: {error}")
        self.results_notebook.select(self.result_tabs["calculation"])

    def _refresh_orders_view(self):
        """Оновлює список наказів вкладки витягів і підсумковий рядок."""
        tree = getattr(self, "orders_tree", None)
        if tree is None:
            return
        self._fill_orders_tree(
            tree, self.manual_order_paths,
            "Накази не обрано — натисніть «Обрати накази» або перетягніть файли сюди.",
        )
        self._select_order_path(tree, self.doc_path.get())
        self._refresh_source_summary()

    def _refresh_source_summary(self):
        """Один рядок, який чесно каже, що саме піде в обробку."""
        if not hasattr(self, "source_summary"):
            return
        paths, from_copies = self._selected_order_paths()
        if not paths:
            self.source_summary.set("Джерело не обрано — оберіть наказ або кілька наказів")
            self.message_source_summary.set("Наказ для повідомлень не обрано")
            return
        names = ", ".join(os.path.basename(path) for path in paths[:3])
        if len(paths) > 3:
            names += f" … ще {len(paths) - 3}"
        prefix = "Примірників № 2 в обробці" if from_copies else "Наказів в обробці"
        self.source_summary.set(f"{prefix}: {len(paths)} — {names}")
        active = self.doc_path.get()
        self.message_source_summary.set(
            "Один наказ: " + os.path.basename(active) if active and os.path.isfile(active)
            else "Наказ для повідомлень не обрано"
        )

    def _set_orders(self, paths: list[str]):
        """Приймає накази для вкладки витягів і готує папку результату."""
        self.manual_order_paths = [
            os.path.abspath(path) for path in paths if os.path.isfile(path)
        ]
        if self.manual_order_paths:
            # Поточним лишається ПЕРШИЙ наказ: його реквізити показуються в
            # інтерфейсі, з нього ж працюють повідомлення й порівняння.
            self.doc_path.set(self.manual_order_paths[0])
            self.out_folder.set(
                os.path.join(os.path.dirname(self.manual_order_paths[0]), "Extracts_Output")
            )
            self._refresh_order_signer()
        self._refresh_orders_view()
        self.save_config()

    def select_orders(self):
        """Один або кілька наказів — головний (і єдиний) вибір у вікні витягів."""
        paths = filedialog.askopenfilenames(filetypes=[("Word files", "*.docx *.doc")])
        if not paths:
            return
        self._set_orders(list(paths))
        self.log(f"Обрано наказів: {len(self.manual_order_paths)}")

    def select_message_order(self):
        """Повідомлення навмисно одинарні: один наказ — два вихідні файли."""
        path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx *.doc")])
        if path:
            self._set_orders([path])
            self.log(f"Обрано наказ для повідомлень: {os.path.basename(path)}")

    def select_orders_folder(self):
        """Уся тека наказів одним пакетом."""
        folder = filedialog.askdirectory()
        if not folder:
            return
        orders = self._orders_in_folder(folder)
        if not orders:
            messagebox.showwarning("Немає файлів", "У цій теці немає наказів DOCX для обробки.")
            return
        self._set_orders(orders)
        self.log(f"З теки {folder} обрано наказів: {len(orders)}")

    def clear_orders(self):
        self.manual_order_paths = []
        self._refresh_orders_view()
        self.save_config()

    def copy_certifier_from_extracts(self):
        """Переносить засвідчувача з витягів у примірники — коли він той самий."""
        self.p2_certifier_position.set(self.certifier_position.get())
        self.p2_certifier_rank.set(self.certifier_rank.get())
        self.p2_certifier_name.set(self.certifier_name.get())
        self.save_config()
        self.log_p2("Засвідчувача примірників скопійовано з витягів.")

    def _selected_order_paths(self) -> tuple[list[str], bool]:
        """Повертає `(шляхи, чи це примірники № 2)`.

        Прапорець потрібен для вибору папки результату: для примірників,
        створених пакетом, витяги кладемо поряд із ними у папку наказу.
        """
        if self.last_copy_two_paths:
            paths = self._marked_paths(self.copy_two_tree)
            if paths:
                return paths, True
            # Галочки зняті з усіх примірників — працюємо з наказами,
            # обраними вручну, замість того щоб відмовляти користувачу.
        tree = getattr(self, "orders_tree", None)
        if tree is not None:
            marked = self._marked_paths(tree)
            if marked:
                return marked, False
        manual_orders = [path for path in self.manual_order_paths if os.path.isfile(path)]
        if manual_orders:
            return manual_orders, False
        manual_order = self.doc_path.get()
        if manual_order and os.path.isfile(manual_order):
            return [manual_order], False
        return [], False

    def _set_processing_order(self, doc_path: str, alongside: bool = False):
        """Готує наказ до обробки й визначає папку результату.

        `alongside=True` — це примірник № 2 із пакетної обробки: витяги та
        звіти розрахунку кладемо ПОРЯД із ним, у ту саму папку наказу.
        Для вручну обраного наказу лишається окрема `Extracts_Output`,
        щоб не засмічувати папку, де лежать самі накази.
        """
        self.doc_path.set(doc_path)
        order_dir = os.path.dirname(os.path.abspath(doc_path))
        self.out_folder.set(order_dir if alongside else os.path.join(order_dir, "Extracts_Output"))
        os.makedirs(self.out_folder.get(), exist_ok=True)

    # =========================================================================
    # ДІЇ: РОЗРАХУНОК ТА ВИТЯГИ
    # =========================================================================
    def _missing_selected_files(self, *entries: tuple[str, str]) -> list[str]:
        """Перелік «підпис: шлях» для обраних файлів, яких немає на диску.

        Шляхи зберігаються в config.json, тож між запусками файл можуть
        перейменувати, перемістити або лишити на відключеному диску. Без цієї
        перевірки генерація доходить аж до копіювання зразка й падає з
        «[WinError 2] The system cannot find the file specified» — уже після
        розбору наказу й запису контрольних таблиць.
        """
        return [
            f"{label}: {path}"
            for label, path in entries
            if path and not os.path.isfile(path)
        ]

    def _report_missing_files(self, missing: list[str]) -> None:
        """Називає конкретні відсутні файли — і в журналі, і у вікні."""
        details = "\n".join(f"• {item}" for item in missing)
        self.log("ПОМИЛКА: не знайдено обраних файлів:\n" + details)
        messagebox.showwarning(
            "Немає потрібного файлу",
            "Що сталося: файл(и), вибрані для роботи, не лежать там, де вказано:\n\n"
            f"{details}\n\n"
            "Найчастіше файл перейменували, перемістили або він на диску, "
            "який зараз не підключений.\n\n"
            "Що зробити:\n"
            "1. Відкрийте «Зразки та реквізити».\n"
            "2. Натисніть «Вибрати» біля рядка, названого вище.\n"
            "3. Вкажіть потрібний файл і збережіть вікно.\n"
            "4. Натисніть кнопку генерації ще раз.",
        )
