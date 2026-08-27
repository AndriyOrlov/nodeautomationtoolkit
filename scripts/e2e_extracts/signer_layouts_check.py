import sys
sys.stdout.reconfigure(encoding="utf-8")
from importlib import util
spec = util.spec_from_file_location("g", "generate_extracts.py"); g = util.module_from_spec(spec); spec.loader.exec_module(g)

BODY = ["§ 1", "Відповідно до пункту 1 ПРИЗНАЧИТИ:", "", "1. Майора ТЕСТЕНКА А.А., офіцера.", ""]
TAIL = ["Розрахунок розсилки витягів із наказу:", "1. Військова частина А0002\tп. 1."]
POS = ["Тимчасово виконуючий обов'язки", "командувача військ", "оперативного командування «Тест»"]

VARIANTS = {
 "один блок": POS + ["", "полковник   Іван ПЕТРЕНКО", ""],
 "ПОГОДЖЕНО + підписант": ["ПОГОДЖЕНО"] + POS + ["", "полковник   Петро ЗГОДІН", ""] + [""] + POS + ["", "полковник   Іван ПЕТРЕНКО", ""],
 "посада двічі, звання лише в другій": POS + ["", ""] + POS + ["", "полковник   Іван ПЕТРЕНКО", ""],
 "звання ПЕРЕД посадою (інший порядок)": ["полковник   Іван ПЕТРЕНКО"] + POS + [""],
}
for name, sig in VARIANTS.items():
    text = "\n".join(BODY + sig + TAIL)
    s = g._find_order_signer(text)
    body, _ = g.text_before_order_signer(text)
    leak = [m for m in ("Тимчасово виконуючий", "ПОГОДЖЕНО", "полковник") if m in body]
    exp = len(BODY) if name != "ПОГОДЖЕНО + підписант" else len(BODY)  # блок має починатися з першого
    start = s["start_line"] if s else None
    print(f"  {name:<38} start={start} (перший рядок блоку {len(BODY)}) протікає={leak}")
