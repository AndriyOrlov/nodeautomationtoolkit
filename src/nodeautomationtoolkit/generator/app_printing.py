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
    page_count,
    pdf_name,
    plan_steps,
    print_step,
    scan_batch_folder,
    word_is_gone,
)
from .word_doc import force_quit_word

PDF_CHECK_FOLDER = "Друк_перевірка"


class PrintingMixin:
    """Друк пакета наказів."""

    def _print_folder(self) -> str:
        folder = (self.print_folder.get() or self.p2_out_folder.get() or "").strip()
        # Шлях зі скісними «/» (так його дає діалог Qt) Word сприймає як адресу:
        # назви PDF виходили з «%20», а вивантаження окремої сторінки падало.
        return os.path.normpath(folder) if folder else ""

    def print_queue_rows(self) -> list[tuple[str, str, str]]:
        """Рядки черги без Word: (наказ, що друкується, файл)."""
        folder = self._print_folder()
        if not folder or not os.path.isdir(folder):
            return []
        rows = []
        for order in scan_batch_folder(folder):
            for caption, path, wanted in (
                ("Витяги — 2 сторінки на аркуш", order.extracts, True),
                ("Витяги до управління", order.management, self.print_management.get()),
                ("Примірник № 2 + задня сторінка", order.copy_two, True),
                ("Примірник № 3 + задня сторінка (лише друк)", order.copy_two, True),
            ):
                if path and wanted:
                    rows.append((order.title, caption, os.path.basename(path)))
        return rows

    def _print_pause(self, done: str, after: str, upcoming: str) -> bool:
        """Пауза після кроку. True — далі, False — зупинити друк."""
        text = f"Надруковано: {done}.\n\n"
        if after:
            text += after + "\n\n"
        text += f"Відкладіть надруковане.\n\nДалі: {upcoming}" if upcoming else "Це був останній крок."
        return bool(messagebox.askokcancel("Друк", text))

    def run_print_action(self, pdf_check: bool = False):
        folder = self._print_folder()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("Друк", "Оберіть теку пакета (там, де лежать примірники № 2 і витяги).")
            return None
        orders = scan_batch_folder(folder)
        if not orders:
            messagebox.showwarning(
                "Друк", "У теці не знайдено ні витягів, ні примірників із номером і датою наказу в назві."
            )
            return None
        pdf_dir = os.path.join(folder, PDF_CHECK_FOLDER) if pdf_check else ""
        if pdf_dir:
            os.makedirs(pdf_dir, exist_ok=True)
        self.save_config()
        self.log(
            ("🧪 Перевірка друку у PDF: " if pdf_check else "🖨 Друк пакета: ")
            + f"{len(orders)} наказ(ів) з теки {folder}"
        )
        word = None
        printed = 0
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            all_steps = []
            for order in orders:
                copy_pages = 0
                if order.copy_two:
                    doc = word.Documents.Open(
                        os.path.abspath(order.copy_two), ReadOnly=True, AddToRecentFiles=False
                    )
                    try:
                        copy_pages = page_count(doc)
                    finally:
                        doc.Close(False)
                all_steps.extend(
                    plan_steps(order, copy_pages, self.print_management.get(), self.print_feed_hint.get().strip())
                )
            skip_copy_three: set[str] = set()
            for index, step in enumerate(all_steps, start=1):
                if step.copy_three and step.order in skip_copy_three:
                    continue
                self.log(f"  {index}/{len(all_steps)} {step.order}: {step.what}")
                try:
                    print_step(word, step, os.path.join(pdf_dir, pdf_name(index, step)) if pdf_dir else "")
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
                if pdf_dir:
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
        word = None
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            for side in ("ЛИЦЕ", "ЗВОРОТ"):
                doc = word.Documents.Add()
                try:
                    doc.Content.Text = f"{side}\r↑ ВЕРХ АРКУША ↑"
                    doc.Content.Font.Size = 48
                    doc.Content.Font.Bold = True
                    doc.Content.ParagraphFormat.Alignment = 1  # по центру
                    doc.PrintOut(Background=False, Copies=1)
                finally:
                    doc.Close(False)
                if side == "ЛИЦЕ" and not self._print_pause(
                    "пробний аркуш, лице",
                    "Візьміть цей аркуш і вкладіть у ручну подачу так, як вважаєте правильним.",
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
