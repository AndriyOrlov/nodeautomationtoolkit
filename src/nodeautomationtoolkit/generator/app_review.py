"""Кнопка «Перевірити наказ».

Винесено з `generate_extracts.py` без змін логіки (26.10.2026).
"""
from __future__ import annotations

import os
from pathlib import Path
from tkinter import messagebox
import win32com.client
from nodeautomationtoolkit.builtin_nodes.recipient_mapping import (
    read_recipient_mapping,
)
from nodeautomationtoolkit.order_review.check import (
    default_index_folder,
    group_findings,
)
from nodeautomationtoolkit.order_review.marking import save_marked_copy

from .errors import explain_error
from .naming import sanitize_filename
from .reports import _save_table_to_excel
from .review_cli import review_order_text
from .word_doc import force_quit_word


class ReviewMixin:
    """Кнопка «Перевірити наказ»."""

    # =========================================================================
    # ПЕРЕВІРКА НАКАЗУ (наприклад, складеного студентом)
    # =========================================================================
    def run_order_review_action(self):
        """Перевіряє обрані накази без еталона: правила оформлення + індекс попередніх наказів.

        Нічого не генерує й не змінює оригінал: поруч із наказом у теці
        «Перевірка» з'являються звіт Excel і позначена копія з примітками Word.
        """
        self.save_config()
        order_paths, _from_copies = self._selected_order_paths()
        if not order_paths:
            messagebox.showwarning("Помилка", "Оберіть наказ (або кілька), який треба перевірити.")
            return

        mapping = {}
        excel = self.excel_path.get()
        if excel and os.path.isfile(excel):
            mapping = read_recipient_mapping(path=excel).get("mapping", {})
        else:
            self.log("УВАГА: таблицю відповідностей не вибрано — адресати пунктів не перевірялись.")
        index_folder = default_index_folder(self.new_order_index_folder.get())

        self.log("\n=== ПЕРЕВІРКА НАКАЗІВ ===")
        reports, failures = [], []
        for number, order_path in enumerate(order_paths, start=1):
            name = os.path.basename(order_path)
            self.log(f"\n[{number}/{len(order_paths)}] Перевірка: {name}")
            try:
                reports.append((name, *self._review_one_order(order_path, mapping, index_folder)))
            except Exception as error:
                explanation = explain_error(error)
                self.log(f"  ПОМИЛКА перевірки «{name}»:\n    " + explanation.replace("\n", "\n    "))
                failures.append((name, explanation))

        lines = [
            f"{name}: помилок {result.errors}, зауважень {result.warnings}"
            for name, result, _folder in reports
        ]
        lines += [f"{name}: не перевірено — {error.splitlines()[0]}" for name, error in failures]
        folders = sorted({folder for _name, _result, folder in reports})
        messagebox.showinfo(
            "Перевірка наказів",
            "\n".join(lines[:15])
            + (f"\n… ще {len(lines) - 15}" if len(lines) > 15 else "")
            + ("\n\nЗвіт і позначена копія:\n" + "\n".join(folders[:3]) if folders else ""),
        )

    def _review_one_order(self, order_path: str, mapping: dict, index_folder: str):
        """Перевіряє один наказ; повертає (результат, тека зі звітом)."""
        # Сирий текст: «1.Капітана» без пропуску перевірка має побачити таким, як є.
        result = review_order_text(
            self._read_word_text(order_path, normalize=False), order_path, mapping, index_folder
        )

        for note in result.notes:
            self.log(f"  {note}")
        self.log(f"  Помилок: {result.errors}; зауважень: {result.warnings}.")
        # У журнал — лише правило й номери пунктів, без ПІБ і РНОКПП.
        for line in group_findings(result.findings):
            self.log(f"  {line}")

        folder = os.path.join(os.path.dirname(os.path.abspath(order_path)), "Перевірка")
        os.makedirs(folder, exist_ok=True)
        base = sanitize_filename(Path(order_path).stem)
        ordered = sorted(result.findings, key=lambda finding: finding.level != "помилка")
        _save_table_to_excel(
            os.path.join(folder, f"Перевірка_{base}.xlsx"),
            ["Серйозність", "Що перевірялось", "Де", "Що не так", "Як виправити"],
            [(f.level, f.rule, f.where, f.what, f.how) for f in ordered]
            + [("інфо", "Перевірка", "", note, "") for note in result.notes],
        )

        word = None
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            marked, unplaced = save_marked_copy(
                word, order_path, os.path.join(folder, f"{base} — перевірка.docx"), ordered
            )
            self.log(
                f"  Позначена копія: {marked} позначок у тексті"
                + (f", {unplaced} приміток на початку документа" if unplaced else "") + "."
            )
        except Exception as error:
            self.log(f"  УВАГА: позначену копію не створено ({explain_error(error).splitlines()[0]}); "
                     "звіт Excel збережено.")
        finally:
            if word:
                force_quit_word(word)
        self.log(f"  Звіт: {folder}")
        return result, folder
