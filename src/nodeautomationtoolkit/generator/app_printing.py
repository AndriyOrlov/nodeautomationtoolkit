"""Друк пакета (розділ «Друк»): черга, пауза після кожного кроку, пробний аркуш.

Правила й порядок — у `printing.py`. Тут — дії кнопок: зібрати чергу з теки
пакета, друкувати крок за кроком із паузою (оболонка показує свою паузу з
Пробілом/Enter, без неї — `askokcancel`), перевірка у PDF і пробний аркуш
для ручної подачі.
"""
from __future__ import annotations

import os
from tkinter import messagebox

import win32com.client

from .printing import (
    default_printer,
    page_count,
    pdf_name,
    plan_steps,
    print_step,
    printer_duplex,
    printer_trays,
    reload_printer,
    select_printer,
    scan_batch_folders,
    set_paper_tray,
    word_is_gone,
)
from .word_doc import force_quit_word

PDF_CHECK_FOLDER = "Друк_перевірка"


class _Empty:
    @staticmethod
    def get() -> str:
        return ""


class PrintingMixin:
    """Друк пакета наказів."""

    def _print_printer(self) -> str:
        """Вибраний принтер; порожньо — принтер Windows за замовчуванням."""
        variable = getattr(self, "print_printer", None)
        return (variable.get() if variable is not None else "").strip()

    def _print_tray(self, name: str) -> int:
        variable = getattr(self, name, None)
        value = (variable.get() if variable is not None else "") or ""
        return int(value) if str(value).strip().lstrip("-").isdigit() else 0

    def _tray_name(self, printer: str, tray: int) -> str:
        return next((title for number, title in printer_trays(printer) if number == tray), "") if tray else ""

    def _print_auto_duplex(self) -> bool:
        variable = getattr(self, "print_auto_duplex", None)
        return bool(variable.get()) if variable is not None else False

    def _print_folder(self) -> str:
        folder = (self.print_folder.get() or self.p2_out_folder.get() or "").strip()
        # Шлях зі скісними «/» (так його дає діалог Qt) Word сприймає як адресу:
        # назви PDF виходили з «%20», а вивантаження окремої сторінки падало.
        return os.path.normpath(folder) if folder else ""

    def _print_folders(self) -> list[str]:
        """Тека примірників № 2/3 і (якщо вказана) окрема тека витягів."""
        folders = [self._print_folder()]
        extracts = (getattr(self, "print_extracts_folder", None) or _Empty()).get().strip()
        if extracts:
            folders.append(os.path.normpath(extracts))
        return [folder for folder in folders if folder]

    def _print_flag(self, name: str) -> bool:
        """Галочка «що друкувати» (витяги / примірники); без неї — так."""
        var = getattr(self, name, None)
        return True if var is None else bool(var.get())

    def _print_excluded(self) -> set[str]:
        """Накази, з яких знято галочку в черзі (ключі `OrderFiles.key`); решта друкуються."""
        if not hasattr(self, "_print_excluded_keys"):
            self._print_excluded_keys = set()
        return self._print_excluded_keys

    def set_print_order_marked(self, key: str, marked: bool) -> None:
        (self._print_excluded().discard if marked else self._print_excluded().add)(key)

    def print_queue_orders(self) -> list[tuple[bool, str, str, str, str]]:
        """Черга по наказах: (позначено, наказ, що друкується, файли, ключ)."""
        grouped: dict[str, list] = {}
        titles: dict[str, str] = {}
        for title, caption, name, key in self._print_queue_entries():
            entry = grouped.setdefault(key, [[], []])
            titles[key] = title
            entry[0].append(caption)
            if name not in entry[1]:
                entry[1].append(name)
        excluded = self._print_excluded()
        return [
            (key not in excluded, titles[key], "; ".join(captions), "; ".join(names), key)
            for key, (captions, names) in grouped.items()
        ]

    def print_queue_rows(self) -> list[tuple[str, str, str]]:
        """Рядки черги без Word: (наказ, що друкується, файл)."""
        return [(title, caption, name) for title, caption, name, _key in self._print_queue_entries()]

    def _print_queue_entries(self) -> list[tuple[str, str, str, str]]:
        folders = [folder for folder in self._print_folders() if os.path.isdir(folder)]
        if not folders:
            return []
        rows = []
        auto = self._print_auto_duplex()

        def copy_caption(number: int) -> str:
            return (
                f"Примірник № {number}: останній аркуш двобічно (автоматично)"
                if auto else f"Примірник № {number} + задня сторінка"
            )

        for order in scan_batch_folders(folders):
            for caption, path, wanted in (
                ("Витяги — 2 сторінки на аркуш", order.extracts, self._print_flag("print_extracts")),
                ("Витяги до управління — 2 примірники", order.management, self.print_management.get()),
                (copy_caption(2), order.copy_two, self._print_flag("print_copies")),
                (copy_caption(3) + " (лише друк)", order.copy_two, self._print_flag("print_copies")),
            ):
                if path and wanted:
                    rows.append((order.title, caption, os.path.basename(path), order.key))
        return rows

    def _print_pause(self, done: str, after: str, upcoming: str) -> bool:
        """Пауза після кроку. True — далі, False — зупинити друк."""
        text = f"Надруковано: {done}.\n\n"
        if after:
            text += after + "\n\n"
        text += f"Відкладіть надруковане.\n\nДалі: {upcoming}" if upcoming else "Це був останній крок."
        return bool(messagebox.askokcancel("Друк", text))

    def run_print_action(self, pdf_check: bool = False):
        folders = [folder for folder in self._print_folders() if os.path.isdir(folder)]
        if not folders:
            messagebox.showwarning(
                "Друк", "Оберіть теку примірників № 2/3 і, якщо витяги лежать окремо, теку витягів."
            )
            return None
        folder = folders[0]
        orders = scan_batch_folders(folders)
        if not orders:
            messagebox.showwarning(
                "Друк", "У теках не знайдено ні витягів, ні примірників із номером і датою наказу в назві."
            )
            return None
        # Лише накази з галочкою в черзі (05.11.2026).
        excluded = self._print_excluded()
        skipped = [order.title for order in orders if order.key in excluded]
        orders = [order for order in orders if order.key not in excluded]
        if skipped:
            self.log("Без галочки, не друкуються: " + ", ".join(skipped))
        if not orders:
            messagebox.showwarning("Друк", "Не позначено жодного наказу для друку.")
            return None
        pdf_dir = os.path.join(folder, PDF_CHECK_FOLDER) if pdf_check else ""
        if pdf_dir:
            os.makedirs(pdf_dir, exist_ok=True)
        self.save_config()
        self.log(
            ("🧪 Перевірка друку у PDF: " if pdf_check else "🖨 Друк пакета: ")
            + f"{len(orders)} наказ(ів) з тек: {'; '.join(folders)}"
        )
        word = None
        printed = 0
        auto_duplex = self._print_auto_duplex() and not pdf_dir
        printer = self._print_printer() or default_printer()
        duplex = printer_duplex(printer) if auto_duplex and printer else None
        duplex_on: bool | None = None
        front_tray = 0 if pdf_dir else self._print_tray("print_front_tray")
        back_tray = 0 if pdf_dir else self._print_tray("print_back_tray")
        back_tray_name = self._tray_name(printer, back_tray)
        if not pdf_dir:
            self.log(
                f"Принтер: {printer or 'за замовчуванням'}"
                + ("; двобічний друк — автоматичний." if auto_duplex else "; ручна подача.")
            )
            if front_tray or back_tray:
                self.log(
                    f"Лотки: лицьові — {self._tray_name(printer, front_tray) or 'як у принтера'}; "
                    f"зворот — {back_tray_name or 'як у принтера'}."
                )
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            if not pdf_dir:
                # Порожнє поле — явно принтер Windows за замовчуванням: інакше Word
                # друкує на свій останній вибраний, а програма пише в журнал інший (18.11.2026).
                select_printer(word, printer)
            all_steps = []
            include_extracts = self._print_flag("print_extracts")
            include_copies = self._print_flag("print_copies")
            for order in orders:
                copy_pages = 0
                if include_copies and order.copy_two:
                    doc = word.Documents.Open(
                        os.path.abspath(order.copy_two), ReadOnly=True, AddToRecentFiles=False
                    )
                    try:
                        copy_pages = page_count(doc)
                    finally:
                        doc.Close(False)
                all_steps.extend(
                    plan_steps(
                        order, copy_pages, self.print_management.get(),
                        self.print_feed_hint.get().strip(), auto_duplex=self._print_auto_duplex(),
                        back_tray_name=back_tray_name,
                        include_extracts=include_extracts, include_copies=include_copies,
                    )
                )
            if not all_steps:
                self.log("Нічого друкувати: для позначених наказів не вибрано ні витягів, ні примірників.")
                messagebox.showwarning(
                    "Друк", "Нічого друкувати: увімкніть «Витяги», «Витяги до управління» або «Примірники № 2 і № 3»."
                )
                return 0
            skip_copy_three: set[str] = set()
            for index, step in enumerate(all_steps, start=1):
                if step.copy_three and step.order in skip_copy_three:
                    continue
                self.log(
                    f"  {index}/{len(all_steps)} {step.order}: {step.what}"
                    + (f" [стор. {step.pages}]" if step.pages else " [увесь файл]")
                    + ("" if step.pause_after else " → без паузи")
                )
                if duplex is not None and step.duplex != duplex_on:
                    # Двобічність — лише на останній аркуш примірника; Word бачить
                    # зміну після повторного вибору принтера.
                    duplex.set(step.duplex)
                    reload_printer(word, printer)
                    duplex_on = step.duplex
                try:
                    print_step(
                        word, step, os.path.join(pdf_dir, pdf_name(index, step)) if pdf_dir else "",
                        tray=back_tray if step.back_side else front_tray,
                    )
                    printed += 1
                except Exception as error:
                    self.log(f"  ПОМИЛКА: {step.order}, {step.what}: {error}")
                    if word_is_gone(error):
                        raise RuntimeError(
                            "Word закрився посеред друку — решту кроків не надруковано. "
                            f"Останній вдалий крок: {printed} з {len(all_steps)}."
                        ) from error
                    if step.copy_three:
                        skip_copy_three.add(step.order)
                    continue
                if pdf_dir or not step.pause_after:
                    continue
                upcoming = next(
                    (
                        f"{later.order} — {later.what}"
                        for later in all_steps[index:]
                        if not (later.copy_three and later.order in skip_copy_three)
                    ),
                    "",
                )
                if not self._print_pause(f"{step.order} — {step.what}", step.after, upcoming):
                    self.log("  Друк зупинено користувачем.")
                    break
        except Exception as error:
            self.log(f"ПОМИЛКА друку: {error}")
            messagebox.showerror("Друк", f"Друк перервано:\n{error}")
            return printed
        finally:
            if word is not None:
                force_quit_word(word)
            if duplex is not None:
                try:
                    duplex.restore()
                except Exception as error:
                    self.log(f"УВАГА: не вдалося повернути налаштування двобічного друку принтера: {error}")
        if pdf_dir:
            self.log(f"Готово: {printed} PDF у теці {pdf_dir} (2 сторінки на аркуш у PDF не видно).")
            messagebox.showinfo("Перевірка друку", f"Створено PDF кроків: {printed}.\n\nТека: {pdf_dir}")
        else:
            self.log(f"Готово: надруковано кроків {printed}.")
        return printed

    def run_print_check_pdf(self):
        return self.run_print_action(pdf_check=True)

    def run_print_test_sheet(self):
        """Один аркуш: «ЛИЦЕ», пауза для ручної подачі, «ЗВОРОТ» — як класти папір."""
        if self._print_auto_duplex():
            return self._print_duplex_test_sheet()
        word = None
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            printer = self._print_printer() or default_printer()
            select_printer(word, printer)
            back_tray_name = self._tray_name(printer, self._print_tray("print_back_tray"))
            for side, tray in (
                ("ЛИЦЕ", self._print_tray("print_front_tray")),
                ("ЗВОРОТ", self._print_tray("print_back_tray")),
            ):
                doc = word.Documents.Add()
                try:
                    doc.Content.Text = f"{side}\r↑ ВЕРХ АРКУША ↑"
                    doc.Content.Font.Size = 48
                    doc.Content.Font.Bold = True
                    doc.Content.ParagraphFormat.Alignment = 1  # по центру
                    set_paper_tray(doc, tray)
                    doc.PrintOut(Background=False, Copies=1)
                finally:
                    doc.Close(False)
                if side == "ЛИЦЕ" and not self._print_pause(
                    "пробний аркуш, лице",
                    "Візьміть цей аркуш і покладіть "
                    + (f"у лоток «{back_tray_name}»" if back_tray_name else "у ручну подачу")
                    + " так, як вважаєте правильним.",
                    "пробний аркуш, зворот",
                ):
                    return
            messagebox.showinfo(
                "Пробний аркуш",
                "Подивіться на аркуш: якщо «ЗВОРОТ» на звороті й обидві стрілки дивляться в один бік, "
                "так і кладіть. Запишіть спосіб у поле «Як класти в ручну подачу» — "
                "програма нагадуватиме його в паузах.",
            )
        except Exception as error:
            self.log(f"ПОМИЛКА пробного аркуша: {error}")
            messagebox.showerror("Пробний аркуш", str(error))
        finally:
            if word is not None:
                force_quit_word(word)

    def _print_duplex_test_sheet(self):
        """Автоматичний дуплекс: один аркуш, «ЛИЦЕ» і «ЗВОРОТ» з двох боків самим принтером."""
        printer = self._print_printer() or default_printer()
        duplex = printer_duplex(printer) if printer else None
        word = None
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            if duplex is not None:
                duplex.set(True)
            reload_printer(word, printer)
            doc = word.Documents.Add()
            try:
                doc.Content.Text = "ЛИЦЕ\r↑ ВЕРХ АРКУША ↑\r"
                end = doc.Content
                end.Collapse(0)  # wdCollapseEnd
                end.InsertBreak(7)  # wdPageBreak
                doc.Content.InsertAfter("ЗВОРОТ\r↑ ВЕРХ АРКУША ↑")
                doc.Content.Font.Size = 48
                doc.Content.Font.Bold = True
                doc.Content.ParagraphFormat.Alignment = 1  # по центру
                doc.PrintOut(Background=False, Copies=1)
            finally:
                doc.Close(False)
            messagebox.showinfo(
                "Пробний аркуш",
                f"Принтер: {printer or 'за замовчуванням'}.\n\nМає вийти ОДИН аркуш: «ЛИЦЕ» з одного боку, "
                "«ЗВОРОТ» — з іншого, стрілки в один бік. Якщо вийшло два аркуші — принтер не "
                "друкує з двох боків сам: зніміть галочку автоматичного двобічного друку.",
            )
        except Exception as error:
            self.log(f"ПОМИЛКА пробного аркуша: {error}")
            messagebox.showerror("Пробний аркуш", str(error))
        finally:
            if word is not None:
                force_quit_word(word)
            if duplex is not None:
                try:
                    duplex.restore()
                except Exception as error:
                    self.log(f"УВАГА: не вдалося повернути налаштування двобічного друку принтера: {error}")

