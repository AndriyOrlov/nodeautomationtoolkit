"""Ізольований дослід: шапка + порожній абзац + пункт на межі сторінки (правило 5.4).

СТАРЕ правило: KeepWithNext лише на абзаці шапки.
НОВЕ правило: KeepWithNext на всьому діапазоні «шапка → перший пункт».

Синтетичний документ, жодного наказу: у нього доливається наповнення рядок за
рядком, доки шапка не опиниться на самій межі сторінки. Саме у вузькому вікні
(один рядок наповнення з тринадцяти) старе правило лишає шапку внизу сторінки
без її першого пункту — бо порожній абзац між ними не був ні до чого зчеплений.

    python scripts/e2e_extracts/check_heading_keep_rule.py

Очікується рядок «САМЕ ЦЕЙ ВИПАДОК»: старе — ВІДІРВАНА, нове — разом.
"""
import sys
import win32com.client
sys.stdout.reconfigure(encoding="utf-8")

word = win32com.client.DispatchEx("Word.Application")
word.Visible = False; word.DisplayAlerts = 0

def build(rule, lines):
    doc = word.Documents.Add()
    doc.PageSetup.TopMargin = 28; doc.PageSetup.BottomMargin = 28
    # Наповнювач, щоб шапка опинилась унизу першої сторінки
    filler = "\r".join(f"Рядок наповнення {i}" for i in range(1, lines))
    doc.Content.Text = filler + "\r§ 2\r\rПерший пункт розділу: текст пункту, який має\rбути разом із шапкою.\r"
    doc.Repaginate()
    paragraphs = list(doc.Paragraphs)
    heading = next(p for p in paragraphs if str(p.Range.Text).startswith("§ 2"))
    blank = heading.Next()
    item = blank.Next()
    # У справжньому генераторі пункт ЗАВЖДИ цілісний (правило 5.4),
    # тому він переїжджає на наступну сторінку блоком, а не розривається.
    item.Range.ParagraphFormat.KeepTogether = True
    if rule == "old":
        pf = heading.Range.ParagraphFormat
        pf.KeepTogether = True; pf.KeepWithNext = True
    else:
        span = doc.Range(heading.Range.Start, item.Range.Start)
        pf = span.ParagraphFormat
        pf.KeepTogether = True; pf.KeepWithNext = True
    doc.Repaginate()
    result = (heading.Range.Information(3), blank.Range.Information(3), item.Range.Information(3))
    doc.Close(False)
    return result

try:
    for lines in range(24, 37):
        old = build("old", lines)
        new = build("new", lines)
        flag_old = "ВІДІРВАНА" if old[0] != old[2] else "разом"
        flag_new = "ВІДІРВАНА" if new[0] != new[2] else "разом"
        mark = "  <<< САМЕ ЦЕЙ ВИПАДОК" if old[0] != old[2] and new[0] == new[2] else ""
        print(f"наповнення {lines-1:2d} рядків | старе: шапка стор.{old[0]} пункт стор.{old[2]} {flag_old}"
              f" | нове: шапка стор.{new[0]} пункт стор.{new[2]} {flag_new}{mark}")
finally:
    word.Quit()
