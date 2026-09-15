"""Наскрізний прогін СПРАВЖНЬОГО App.run_extracts на локальних фікстурах (C:).

GUI не будується: у цьому Python немає ttkbootstrap, а фолбек на голий
tkinter.ttk падає на `bootstyle=` (окрема знахідка). Тому App створюється
через __new__, і йому підставляються лише ті атрибути, які run_extracts
реально читає. Логіка генерації — справжня, не імітація.
"""
import glob
import os
import sys

sys.stdout.reconfigure(encoding="utf-8")

E2E = sys.argv[1]
PROJECT = sys.argv[2]
os.chdir(PROJECT)
sys.path.insert(0, os.path.join(PROJECT, "src"))

from importlib import util  # noqa: E402

spec = util.spec_from_file_location("gen_e2e", os.path.join(PROJECT, "generate_extracts.py"))
gen = util.module_from_spec(spec)
spec.loader.exec_module(gen)


class _Silent:
    @staticmethod
    def showinfo(*a, **k): print("[info]", (a[1] if len(a) > 1 else "")[:300])
    @staticmethod
    def showwarning(*a, **k): print("[warning]", (a[1] if len(a) > 1 else "")[:300])
    @staticmethod
    def showerror(*a, **k): print("[error]", (a[1] if len(a) > 1 else "")[:600])
    @staticmethod
    def askyesno(*a, **k): return False


gen.messagebox = _Silent
gen.os.startfile = lambda *a, **k: None


class Var:
    def __init__(self, value=""): self._v = value
    def get(self): return self._v
    def set(self, value): self._v = value


order = glob.glob(os.path.join(E2E, "Наказ*.docx"))[0]
out_folder = os.path.join(E2E, "Extracts_Output")

app = gen.App.__new__(gen.App)
app.log = lambda message="": print(message)
app.show_analysis_results = lambda *a, **k: None
app.show_layout_warnings = lambda *a, **k: [print("[layout]", w) for w in (a[0] if a else [])]
app.excel_path = Var(os.path.join(E2E, "mapping.xlsx"))
app.doc_path = Var(order)
app.template_path = Var(os.path.join(E2E, "template.docx"))
app.out_folder = Var(out_folder)
app.executor = Var("Тест Тестенко 00-000")
app.certifier_position = Var("Начальник відділу /штабу військової частини А0001")
app.certifier_rank = Var("підполковник")
app.certifier_name = Var("Іван ЗАСВІДЧУВАЧ")
app.order_signer_position = Var("")
app.order_signer_rank = Var("")
app.order_signer_name = Var("")
app.group_corps_var = Var(True)
app.duplex_2up_layout = Var(True)

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
    app.run_extracts()
except Exception as exc:
    import traceback
    traceback.print_exc()
    print("ЗБІЙ:", exc)
# Наприкінці run_extracts відкриває результат у Word. Закриваємо ЛИШЕ ті
# екземпляри, які зʼявилися під час цього прогону — чужий Word не чіпаємо.
import subprocess
for pid in word_pids() - before:
    subprocess.run(["powershell", "-NoProfile", "-Command",
                    f"(Get-Process -Id {pid}).CloseMainWindow()"],
                   capture_output=True)
    print(f"[harness] закрито Word, відкритий прогоном: PID {pid}")

print("=" * 72)
print("ФАЙЛИ В РЕЗУЛЬТАТІ:")
for path in sorted(glob.glob(os.path.join(out_folder, "*"))):
    print("  ", os.path.basename(path), os.path.getsize(path), "байт")
