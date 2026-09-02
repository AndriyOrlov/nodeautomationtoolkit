"""Наскрізний прогін СПРАВЖНЬОГО App.run_generate_messages на фікстурах (C:).

GUI не будується (правило 8.4: фолбек без ttkbootstrap падає), тому App
створюється через __new__ — так само, як у run_extracts_e2e.py. Логіка
генерації справжня, підміняються лише вікна повідомлень і кнопка.
"""
import glob
import os
import sys

sys.stdout.reconfigure(encoding="utf-8")

E2E = sys.argv[1]
PROJECT = sys.argv[2] if len(sys.argv) > 2 else "."
PROJECT = os.path.abspath(PROJECT)
os.chdir(PROJECT)
sys.path.insert(0, os.path.join(PROJECT, "src"))

from importlib import util  # noqa: E402

spec = util.spec_from_file_location("gen_e2e", os.path.join(PROJECT, "generate_extracts.py"))
gen = util.module_from_spec(spec)
spec.loader.exec_module(gen)


class _Silent:
    @staticmethod
    def showinfo(*a, **k): print("[info]", (a[1] if len(a) > 1 else "")[:400])
    @staticmethod
    def showwarning(*a, **k): print("[warning]", (a[1] if len(a) > 1 else "")[:400])
    @staticmethod
    def showerror(*a, **k): print("[error]", (a[1] if len(a) > 1 else "")[:800])
    @staticmethod
    def askyesno(*a, **k): return False


gen.messagebox = _Silent
gen.os.startfile = lambda *a, **k: None


class Var:
    def __init__(self, value=""): self._v = value
    def get(self): return self._v
    def set(self, value): self._v = value


class Button:
    def config(self, **kwargs): pass


order = glob.glob(os.path.join(E2E, "Наказ*.docx"))[0]
out_folder = os.path.join(E2E, "Messages_Output")

app = gen.App.__new__(gen.App)
app.log = lambda message="": print(message)
app.save_config = lambda *a, **k: None
app.excel_path = Var(os.path.join(E2E, "mapping.xlsx"))
app.doc_path = Var(order)
app.message_cover_template_path = Var(os.path.join(E2E, "message_cover.docx"))
app.message_content_template_path = Var(os.path.join(E2E, "message_content.docx"))
app.message_out_folder = Var(out_folder)
app.message_executor = Var("Тест Тестенко 00-000")
app.btn_generate_messages = Button()


def word_pids():
    import subprocess
    out = subprocess.run(
        ["powershell", "-NoProfile", "-Command",
         "Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }"],
        capture_output=True, text=True)
    return {line.strip() for line in out.stdout.splitlines() if line.strip()}


before = word_pids()
print("=" * 72)
try:
    app.run_generate_messages()
except Exception as exc:
    import traceback
    traceback.print_exc()
    print("ЗБІЙ:", exc)

import subprocess  # noqa: E402
for pid in word_pids() - before:
    subprocess.run(["powershell", "-NoProfile", "-Command",
                    f"(Get-Process -Id {pid}).CloseMainWindow()"], capture_output=True)
    print(f"[harness] закрито Word, відкритий прогоном: PID {pid}")

print("=" * 72)
print("ФАЙЛИ В РЕЗУЛЬТАТІ:")
for path in sorted(glob.glob(os.path.join(out_folder, "*"))):
    print("  ", os.path.basename(path), os.path.getsize(path), "байт")
