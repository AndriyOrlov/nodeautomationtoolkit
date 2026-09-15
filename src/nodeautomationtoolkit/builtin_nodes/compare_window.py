"""Візуальний інтерфейс режиму порівняння DOCX документів (Side-by-Side Diff Window)."""

from __future__ import annotations

import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Optional

try:
    import ttkbootstrap as tb
    from ttkbootstrap.constants import BOTH, DISABLED, HORIZONTAL, LEFT, NORMAL, RIGHT, VERTICAL, W, X, Y
    HAS_TTKBOOTSTRAP = True
except ImportError:
    tb = None
    HAS_TTKBOOTSTRAP = False
    from tkinter.constants import BOTH, DISABLED, HORIZONTAL, LEFT, NORMAL, RIGHT, VERTICAL, W, X, Y

from nodeautomationtoolkit.builtin_nodes.compare_documents import CompareResult, compare_docx_documents


class DocxCompareWindow:
    """Двопанельне вікно візуального порівняння еталонного та згенерованого DOCX файлів."""

    def __init__(
        self,
        parent: tk.Tk | tk.Toplevel,
        reference_path: Optional[str] = None,
        generated_path: Optional[str] = None,
        mode: str = "витяги",
    ):
        self.parent = parent
        self.window = tk.Toplevel(parent)
        self.window.title("🔍 Режим Порівняння документів (Compare Mode) — Side-by-Side Diff")
        self.window.geometry("1180x760")
        self.window.minsize(800, 500)

        self.ref_var = tk.StringVar(value=reference_path or "")
        self.gen_var = tk.StringVar(value=generated_path or "")
        self.mode_var = tk.StringVar(value=mode)
        self.summary_var = tk.StringVar(value="— Оберіть файли та натисніть «Порівняти» —")

        self.last_result: Optional[CompareResult] = None

        self._build_ui()

        if self.ref_var.get() and self.gen_var.get():
            self.run_comparison()

    def _build_ui(self):
        # Головний контейнер
        main_frame = ttk.Frame(self.window, padding=10)
        main_frame.pack(fill=BOTH, expand=True)

        # ── Панель вибору файлів (Top bar) ───────────────────────────────────
        top_card = ttk.LabelFrame(main_frame, text=" Вибір файлів для порівняння ", padding=8)
        top_card.pack(fill=X, pady=(0, 8))
        top_card.columnconfigure(1, weight=1)

        # Рядок 1: Еталон
        ttk.Label(top_card, text="📁 Еталон (ручний DOCX):", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky=W, padx=(0, 6), pady=2)
        ttk.Entry(top_card, textvariable=self.ref_var).grid(row=0, column=1, sticky="ew", padx=(0, 6), pady=2)
        ttk.Button(top_card, text="📂 Вибрати", command=self._select_ref_file).grid(row=0, column=2, pady=2)

        # Рядок 2: Згенерований файл
        ttk.Label(top_card, text="⚙️ Згенерований DOCX:", font=("Segoe UI", 9, "bold")).grid(row=1, column=0, sticky=W, padx=(0, 6), pady=2)
        ttk.Entry(top_card, textvariable=self.gen_var).grid(row=1, column=1, sticky="ew", padx=(0, 6), pady=2)
        ttk.Button(top_card, text="📂 Вибрати", command=self._select_gen_file).grid(row=1, column=2, pady=2)

        # Рядок 3: Кнопка дії та режим
        action_bar = ttk.Frame(top_card)
        action_bar.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(6, 0))

        ttk.Label(action_bar, text="Режим:").pack(side=LEFT, padx=(0, 6))
        mode_cb = ttk.Combobox(action_bar, textvariable=self.mode_var, values=["витяги", "примірник_2", "повідомлення_зміст", "повідомлення_супровід"], state="readonly", width=22)
        mode_cb.pack(side=LEFT, padx=(0, 14))

        self.btn_compare = ttk.Button(action_bar, text="🔄 Порівняти файли", command=self.run_comparison)
        self.btn_compare.pack(side=LEFT, padx=(0, 8))

        # Легенда кольорів
        legend_frame = ttk.Frame(action_bar)
        legend_frame.pack(side=RIGHT)
        ttk.Label(legend_frame, text="Легенда: ").pack(side=LEFT)
        l_del = tk.Label(legend_frame, text=" Пропущено / Видалено ", bg="#ffdce0", fg="#b30000", padx=4, relief="groove")
        l_del.pack(side=LEFT, padx=2)
        l_ins = tk.Label(legend_frame, text=" Зайве / Додано ", bg="#dcffe4", fg="#007020", padx=4, relief="groove")
        l_ins.pack(side=LEFT, padx=2)
        l_mod = tk.Label(legend_frame, text=" Змінено / Стиль ", bg="#fff5b1", fg="#8a6d00", padx=4, relief="groove")
        l_mod.pack(side=LEFT, padx=2)

        # ── Двопанельний Side-by-Side перегляд (Middle area) ───────────────────
        diff_container = ttk.Frame(main_frame)
        diff_container.pack(fill=BOTH, expand=True, pady=(0, 8))
        diff_container.columnconfigure(0, weight=1)
        diff_container.columnconfigure(1, weight=1)
        diff_container.rowconfigure(1, weight=1)

        # Заголовки панелей
        ttk.Label(diff_container, text="📄 ЕТАЛОННИЙ ДОКУМЕНТ (Очікувано)", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky=W, padx=4, pady=(0, 4))
        ttk.Label(diff_container, text="⚙️ ЗГЕНЕРОВАНИЙ ДОКУМЕНТ (Отримано)", font=("Segoe UI", 10, "bold")).grid(row=0, column=1, sticky=W, padx=4, pady=(0, 4))

        # Ліва панель (Еталон)
        left_frame = ttk.Frame(diff_container)
        left_frame.grid(row=1, column=0, sticky="nsew", padx=(0, 4))
        left_frame.columnconfigure(0, weight=1)
        left_frame.rowconfigure(0, weight=1)

        self.text_left = tk.Text(left_frame, wrap="none", font=("Consolas", 10), padx=6, pady=6)
        self.text_left.grid(row=0, column=0, sticky="nsew")

        # Права панель (Згенеровано)
        right_frame = ttk.Frame(diff_container)
        right_frame.grid(row=1, column=1, sticky="nsew", padx=(4, 0))
        right_frame.columnconfigure(0, weight=1)
        right_frame.rowconfigure(0, weight=1)

        self.text_right = tk.Text(right_frame, wrap="none", font=("Consolas", 10), padx=6, pady=6)
        self.text_right.grid(row=0, column=0, sticky="nsew")

        # Спільний вертикальний скролбар
        self.v_scroll = ttk.Scrollbar(diff_container, orient=VERTICAL, command=self._on_vertical_scroll)
        self.v_scroll.grid(row=1, column=2, sticky="ns")
        self.text_left.configure(yscrollcommand=self._on_text_scroll)
        self.text_right.configure(yscrollcommand=self._on_text_scroll)

        # Налаштування стилів виділення
        for txt in (self.text_left, self.text_right):
            txt.tag_configure("EQUAL", background="#ffffff")
            txt.tag_configure("DELETED", background="#ffdce0", foreground="#900000")
            txt.tag_configure("INSERTED", background="#dcffe4", foreground="#006010")
            txt.tag_configure("MODIFIED", background="#fff5b1", foreground="#705000")
            txt.tag_configure("HEADER", font=("Consolas", 10, "bold"), foreground="#003366")

        # ── Нижня панель статусів та дій (Bottom bar) ────────────────────────
        bottom_card = ttk.Frame(main_frame)
        bottom_card.pack(fill=X)

        self.lbl_status = ttk.Label(
            bottom_card,
            textvariable=self.summary_var,
            font=("Segoe UI", 10, "bold"),
        )
        self.lbl_status.pack(side=LEFT, fill=X, expand=True, padx=(0, 10))

        self.btn_copy_chat = ttk.Button(
            bottom_card,
            text="📋 Скопіювати опис для чату AI",
            command=self.copy_chat_report,
            state=DISABLED,
        )
        self.btn_copy_chat.pack(side=RIGHT, padx=(4, 0))

        self.btn_save_report = ttk.Button(
            bottom_card,
            text="💾 Зберегти звіт (.md)",
            command=self.save_report_file,
            state=DISABLED,
        )
        self.btn_save_report.pack(side=RIGHT, padx=(4, 0))

    def _select_ref_file(self):
        filename = filedialog.askopenfilename(
            title="Оберіть еталонний файл DOCX",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")],
            parent=self.window,
        )
        if filename:
            self.ref_var.set(filename)

    def _select_gen_file(self):
        filename = filedialog.askopenfilename(
            title="Оберіть згенерований файл DOCX",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")],
            parent=self.window,
        )
        if filename:
            self.gen_var.set(filename)

    def _on_vertical_scroll(self, *args):
        self.text_left.yview(*args)
        self.text_right.yview(*args)

    def _on_text_scroll(self, *args):
        self.v_scroll.set(*args)
        self.text_left.yview_moveto(args[0])
        self.text_right.yview_moveto(args[0])

    def run_comparison(self):
        ref_p = self.ref_var.get().strip()
        gen_p = self.gen_var.get().strip()

        if not (os.path.isfile(ref_p) and os.path.isfile(gen_p)):
            messagebox.showwarning(
                "Помилка вибору",
                "Будь ласка, вкажіть дійсні шляхи до обох DOCX файлів.",
                parent=self.window,
            )
            return

        try:
            result = compare_docx_documents(ref_p, gen_p, mode=self.mode_var.get())
            self.last_result = result

            self.text_left.config(state=NORMAL)
            self.text_right.config(state=NORMAL)
            self.text_left.delete("1.0", tk.END)
            self.text_right.delete("1.0", tk.END)

            for row in result.side_by_side_rows:
                status = row["status"]
                ref_line = row["ref_line"]
                gen_line = row["gen_line"]

                # Ліва панель
                start_l = self.text_left.index(tk.INSERT)
                self.text_left.insert(tk.END, (ref_line if ref_line else "[—]") + "\n")
                end_l = self.text_left.index(tk.INSERT)
                self.text_left.tag_add(status, start_l, end_l)

                # Права панель
                start_r = self.text_right.index(tk.INSERT)
                self.text_right.insert(tk.END, (gen_line if gen_line else "[—]") + "\n")
                end_r = self.text_right.index(tk.INSERT)
                self.text_right.tag_add(status, start_r, end_r)

            self.text_left.config(state=DISABLED)
            self.text_right.config(state=DISABLED)

            self.summary_var.set(result.summary_text)
            self.btn_copy_chat.config(state=NORMAL)
            self.btn_save_report.config(state=NORMAL)

        except Exception as err:
            messagebox.showerror("Помилка порівняння", f"Не вдалося порівняти файли:\n{err}", parent=self.window)

    def copy_chat_report(self):
        if not self.last_result:
            return
        report = self.last_result.ai_chat_report
        self.window.clipboard_clear()
        self.window.clipboard_append(report)
        messagebox.showinfo(
            "Скопійовано",
            "📋 Структурований опис розбіжностей успішно скопійовано в буфер обміну!\n"
            "Тепер просто вставте його (Ctrl+V) у чат AI.",
            parent=self.window,
        )

    def save_report_file(self):
        if not self.last_result:
            return
        save_path = filedialog.asksaveasfilename(
            title="Зберегти звіт порівняння",
            defaultextension=".md",
            filetypes=[("Markdown", "*.md"), ("Text Files", "*.txt")],
            parent=self.window,
        )
        if save_path:
            with open(save_path, "w", encoding="utf-8") as f:
                f.write(self.last_result.ai_chat_report)
            messagebox.showinfo("Збережено", f"Звіт успішно збережено:\n{save_path}", parent=self.window)
