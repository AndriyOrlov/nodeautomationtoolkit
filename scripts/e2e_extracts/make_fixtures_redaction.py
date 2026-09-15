"""Наказ із пунктом «ВИКЛАСТИ В ТАКІЙ РЕДАКЦІЇ» та великою цитатою всередині.

Усередині лапок цитується цілий § з преамбулою й власним пунктом, тому
маршрутизація бачить там і «шапки», і окремий пункт — а їхні діапазони рядків
ПЕРЕКРИВАЮТЬСЯ зі вступним абзацом. Саме на цьому у витягу двічі друкувався
пункт «11.» (див. `check_no_duplicate_lines.py`).

Усе вигадане, усе на C:. Далі — `check_no_duplicate_lines.py <тека> .`
"""
import os, sys
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt
from openpyxl import Workbook

OUT = sys.argv[1]; os.makedirs(OUT, exist_ok=True)
C = WD_ALIGN_PARAGRAPH.CENTER; J = WD_ALIGN_PARAGRAPH.JUSTIFY

def para(doc, text="", align=None, indent=None, bold=False, size=14, left=None):
    p = doc.add_paragraph()
    if align is not None: p.alignment = align
    if indent is not None: p.paragraph_format.first_line_indent = Cm(indent)
    if left is not None:
        p.paragraph_format.left_indent = Cm(left)
        p.paragraph_format.first_line_indent = Cm(0)
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = p.add_run(text); r.font.name = "Times New Roman"; r.font.size = Pt(size); r.bold = bold
    return p

def bio(doc, *lines):
    p = para(doc, lines[0], left=8.0)
    for extra in lines[1:]:
        p.add_run().add_break()
        r = p.add_run(extra); r.font.name = "Times New Roman"; r.font.size = Pt(14)
    return p

d = docx.Document()
s = d.sections[0]
s.left_margin, s.right_margin = Cm(2.0), Cm(1.0); s.top_margin, s.bottom_margin = Cm(1.0), Cm(1.0)
para(d, "МІНІСТЕРСТВО ОБОРОНИ УКРАЇНИ", C, bold=True); para(d)
para(d, "НАКАЗ", C, bold=True, size=22)
para(d, "КОМАНДУВАЧА ВІЙСЬК ОПЕРАТИВНОГО КОМАНДУВАННЯ «ТЕСТ»", C, bold=True); para(d)
para(d, "(по особовому складу)", C); para(d)
para(d, "«15» серпня 2026 року          м. Тестове          № 555", bold=True); para(d)

para(d, "§ 7", C); para(d)
para(d, "11. Пункт 6 наказу командувача військ оперативного командування “Тест” "
        "(по особовому складу) від 07 липня 2026 року № 123 про звільнення старшого "
        "лейтенанта КЛИМЕНКА Андрія Андрійовича, офіцера адміністративного відділення "
        "3 окремого тестового загону, ВИКЛАСТИ В ТАКІЙ РЕДАКЦІЇ:", J, indent=1.25)
para(d)
para(d, "“§1", C); para(d)
para(d, "Відповідно до пункту 2 частини четвертої статті 26 Закону України "
        "“Про військовий обов’язок і військову службу” нижчепойменованих осіб "
        "офіцерського складу ЗВІЛЬНИТИ з військової служби:", J, indent=1.25)
para(d)
para(d, "У ЗАПАС ЗА ПІДПУНКТОМ “г” (через сімейні обставини або з інших поважних "
        "причин, перелік яких визначається частиною дванадцятою цієї статті):", J, indent=1.25)
para(d)
para(d, "1. Старшого лейтенанта КЛИМЕНКА Андрія Андрійовича, офіцера "
        "адміністративного відділення 3 окремого тестового загону.", J, indent=1.25)
bio(d, "1990 р. н., освіта: ТВІ у 2012 р.,", "у ЗС - із 08.2008.")
bio(d, "1234567890.”")
para(d)

# Ще один пункт того самого § і тієї самої частини ПІСЛЯ редакції: перевіряємо,
# чи не причепиться до нього цитована «шапка» з середини пункту 11.
para(d, "12. Майора ПРИКЛАДЕНКА Богдана Богдановича, начальника служби "
        "3 окремого тестового загону, призначити начальником іншої служби "
        "цього самого загону.", J, indent=1.25)
bio(d, "1985 р. н., освіта: ТВІ у 2007 р.,", "у ЗС - із 09.2003.")
bio(d, "2345678901.")
para(d); para(d)
para(d, "Командир військової частини А0001"); para(d)
para(d, "полковник                                        Петро ТЕСТОВИЙ"); para(d)
para(d, "Розрахунок розсилки витягів із наказу:", C)
para(d, "1. Військова частина А0003\tп. 11.")
para(d, "Надр. 2 прим.")

order_path = os.path.join(OUT, "Наказ № 558 від 15.08.2026.docx")
d.save(order_path)

wb = Workbook(); ws = wb.active
ws.append(["Найменування", "Шифр", "Скорочення", "Корпус", "Кому", "Куди"])
ws.append(["3 окремий тестовий загін", "А0003", "3 отз", "",
           "Командиру військової частини А0003", "м. Третє"])
wb.save(os.path.join(OUT, "mapping.xlsx"))

# Зразок витягу — найпростіший: тут перевіряється вставка змісту, а не верстка
# підписанта (для неї є `make_template_table_signer.py`).
tpl = docx.Document()
ts = tpl.sections[0]
ts.left_margin, ts.right_margin = Cm(2.0), Cm(1.0)
ts.top_margin, ts.bottom_margin = Cm(1.0), Cm(1.0)
R = WD_ALIGN_PARAGRAPH.RIGHT
para(tpl, "{{кому}}", R)
para(tpl, "{{куди}}", R)
para(tpl)
para(tpl, "ВИТЯГ З НАКАЗУ", C, bold=True)
para(tpl, "КОМАНДИРА ВІЙСЬКОВОЇ ЧАСТИНИ А0001", C, bold=True)
para(tpl)
para(tpl, "{{дата_наказу}}          м. Тестове          {{номер_наказу}}")
para(tpl)
para(tpl, "{{зміст}}")
para(tpl)
para(tpl, "{{засвідчення}}")
para(tpl)
para(tpl, "{{виконавець}}", size=8)
tpl.save(os.path.join(OUT, "template.docx"))

print("готово:", order_path)
