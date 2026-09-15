# -*- coding: utf-8 -*-
"""Примірник на СПРАВЖНЬОМУ Word: шрифт, нерозривність, порожні білі зображення.

Відтворює три скарги на пакетне формування примірників:

* шрифт «стрибає»: у наказі гарнітура й розмір задані СТИЛЕМ `Normal`, а не
  прямо, і після `FormattedText` текст отримував `Normal` заготовки;
* нерозривність «іноді не працює»: між шапкою та пунктом стоять порожні
  абзаци, а пункти пронумеровані АВТОНУМЕРАЦІЄЮ Word (номера немає в тексті) —
  тоді біографія одного пункту чіплялась до наступного пункту;
* порожні білі зображення з наказу потрапляють у примірник.

Наказ і заготовка створюються тут-таки, дані ВИГАДАНІ: жодного справжнього
наказу, жодного шляху на `E:`. Запуск:

    python scripts/e2e_copies/check_blank_images_and_fonts.py
"""

from __future__ import annotations

import os
import shutil
import sys
import tempfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "..", "src"))

for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

import win32com.client  # noqa: E402
from PIL import Image, ImageDraw  # noqa: E402

from nodeautomationtoolkit.builtin_nodes.copy_generator import build_copy_document  # noqa: E402

_WD_FORMAT_XML = 16
_WD_STYLE_NORMAL = -1
_WD_CHARACTER = 1

ORDER_FONT = ("Arial", 12.0)
TEMPLATE_FONT = ("Courier New", 16.0)
FLOATING_TEXT = "Текст біля плаваючої картинки."

failures: list[str] = []
checks = 0


def check(condition, description: str) -> None:
    global checks
    checks += 1
    print(f"  {'OK  ' if condition else 'ЗБІЙ'} {description}")
    if not condition:
        failures.append(description)


def _text(paragraph) -> str:
    # Вбудоване зображення Word показує в Range.Text символом «/».
    return (paragraph.Range.Text or "").strip("\r\x07 \t")


def make_png(path: str, dotted: bool) -> None:
    image = Image.new("RGB", (120, 60), "white")
    if dotted:
        ImageDraw.Draw(image).rectangle((50, 20, 70, 40), fill="black")
    image.save(path)


def _paragraph_with(doc, marker: str):
    for index in range(1, doc.Paragraphs.Count + 1):
        if marker in (doc.Paragraphs(index).Range.Text or ""):
            return doc.Paragraphs(index)
    raise AssertionError(f"у вигаданому наказі немає «{marker}»")


def _clear_marker(paragraph, marker: str):
    """Прибирає службову мітку з абзацу, лишаючи знак абзацу; повертає діапазон тексту.

    `Find.Execute(Replace=...)` через пізнє зв'язування pywin32 мітку не
    прибирав, тож текст абзацу задаємо прямо.
    """
    text_range = paragraph.Range
    text_range.MoveEnd(_WD_CHARACTER, -1)
    text_range.Text = (text_range.Text or "").replace(marker, "")
    return text_range


def build_order(word, path: str, white_png: str, dotted_png: str, auto_numbered: bool) -> None:
    doc = word.Documents.Add()
    try:
        normal = doc.Styles(_WD_STYLE_NORMAL)
        normal.Font.Name, normal.Font.Size = ORDER_FONT
        first_item = "Полковника ТЕСТОВОГО Тест Тестовича, командира підрозділу."
        second_item = "Майора ДРУГОГО Друг Друговича, заступника командира."
        if not auto_numbered:
            first_item, second_item = "1. " + first_item, "2. " + second_item
        for line in (
            "МІНІСТЕРСТВО ОБОРОНИ ТЕСТОВОЇ КРАЇНИ",
            "НАКАЗ",
            "",
            "§ 1",
            "",
            "Відповідно до статті 1 нижчепойменованих військовослужбовців ПРИЗНАЧИТИ:",
            "",
            first_item,
            "1970 р. н., освіта: вища у 2000 р.",
            "[WHITE]",
            "Призначається на вищу посаду.",
            "",
            second_item,
            "1980 р. н., освіта: вища у 2005 р.",
            "[DOTTED]",
            "[FLOATING]" + FLOATING_TEXT,
            "",
            "Командир тестової частини",
            "полковник                    Тест ТЕСТЕНКО",
        ):
            doc.Content.InsertAfter(line + "\r")

        doc.InlineShapes.AddPicture(white_png, False, True, _clear_marker(_paragraph_with(doc, "[WHITE]"), "[WHITE]"))
        doc.InlineShapes.AddPicture(dotted_png, False, True, _clear_marker(_paragraph_with(doc, "[DOTTED]"), "[DOTTED]"))
        floating_paragraph = _paragraph_with(doc, "[FLOATING]")
        _clear_marker(floating_paragraph, "[FLOATING]")
        doc.Shapes.AddPicture(white_png, False, True, 200, 10, 60, 30, floating_paragraph.Range)

        if auto_numbered:
            for marker in ("Полковника ТЕСТОВОГО", "Майора ДРУГОГО"):
                _paragraph_with(doc, marker).Range.ListFormat.ApplyNumberDefault()
        doc.SaveAs2(path, _WD_FORMAT_XML)
    finally:
        doc.Close(False)


def build_template(word, path: str) -> None:
    doc = word.Documents.Add()
    try:
        normal = doc.Styles(_WD_STYLE_NORMAL)
        normal.Font.Name, normal.Font.Size = TEMPLATE_FONT
        for line in ("НАКАЗ", "{{дата_наказу}}        {{номер_наказу}}", "", "{{зміст}}", "", "{{згідно_з_оригіналом}}"):
            doc.Content.InsertAfter(line + "\r")
        doc.SaveAs2(path, _WD_FORMAT_XML)
    finally:
        doc.Close(False)


def run_case(word, workdir: str, auto_numbered: bool) -> None:
    label = "автонумерація" if auto_numbered else "номери в тексті"
    print(f"\n=== Наказ: {label} ===")
    white_png = os.path.join(workdir, "white.png")
    dotted_png = os.path.join(workdir, "dotted.png")
    make_png(white_png, dotted=False)
    make_png(dotted_png, dotted=True)

    suffix = "auto" if auto_numbered else "typed"
    order = os.path.join(workdir, f"Наказ № 1 від 01.09.2026 {suffix}.docx")
    template = os.path.join(workdir, "template.docx")
    working = os.path.join(workdir, f"_work_{suffix}.docx")
    result = os.path.join(workdir, f"result_{suffix}.docx")
    build_order(word, order, white_png, dotted_png, auto_numbered)
    build_template(word, template)
    shutil.copy2(template, working)

    source = word.Documents.Open(order, ReadOnly=True)
    try:
        # Спершу — що сам вигаданий наказ зібрано правильно, інакше перевірки
        # нижче нічого не доводять.
        check(source.InlineShapes.Count == 2, f"наказ: два вбудовані зображення (є {source.InlineShapes.Count})")
        check(source.Shapes.Count == 1, f"наказ: одне плаваюче зображення (є {source.Shapes.Count})")
        texts = [_text(source.Paragraphs(i)) for i in range(1, source.Paragraphs.Count + 1)]
        check("[WHITE]" not in "".join(texts) and "[FLOATING]" not in "".join(texts), "наказ: службові мітки прибрано")
        if auto_numbered:
            item_paragraph = _paragraph_with(source, "Полковника ТЕСТОВОГО")
            list_string = str(item_paragraph.Range.ListFormat.ListString or "")
            check(bool(list_string) and not _text(item_paragraph)[:1].isdigit(), f"наказ: номер пункту лише в автонумерації («{list_string}»)")
        body_start = texts.index("§ 1") + 1
        body_end = max(i for i, text in enumerate(texts, 1) if text.endswith("Тест ТЕСТЕНКО"))
    finally:
        source.Close(False)

    pages = build_copy_document(
        word,
        order,
        working,
        result,
        {"{{дата_наказу}}": "“01” вересня 2026 року", "{{номер_наказу}}": "№1",
         "{{згідно_з_оригіналом}}": "Згідно з оригіналом\rНачальник штабу\rполковник Іван ІВАНЕНКО"},
        resolve_span=lambda _source: (body_start, body_end),
        log=print,
    )
    print(f"  (сторінок: {pages})")

    doc = word.Documents.Open(result, ReadOnly=True)
    try:
        paragraphs = [(i, _text(doc.Paragraphs(i))) for i in range(1, doc.Paragraphs.Count + 1)]

        def index_of(prefix: str):
            return next((i for i, text in paragraphs if text.startswith(prefix)), None)

        def keep_with_next(position) -> bool:
            return bool(doc.Paragraphs(position).Range.ParagraphFormat.KeepWithNext)

        # ---- шрифт наказу, а не заготовки ----
        for prefix in ("Відповідно до статті", "Полковника", "1. Полковника", "1970 р. н.", "Призначається"):
            position = index_of(prefix)
            if position is None:
                continue
            font = doc.Paragraphs(position).Range.Font
            check(
                (font.Name, float(font.Size)) == ORDER_FONT,
                f"«{prefix}…»: шрифт наказу {ORDER_FONT} (фактично {font.Name} {font.Size})",
            )

        # ---- порожні білі зображення ----
        check(doc.InlineShapes.Count == 1, f"лишилось лише зображення з крапкою (вбудованих: {doc.InlineShapes.Count})")
        check(doc.Shapes.Count == 0, f"біле плаваюче зображення не потрапило (плаваючих: {doc.Shapes.Count})")
        check(index_of(FLOATING_TEXT) is not None, "текст біля плаваючої картинки лишився")
        biography = index_of("1970 р. н.")
        if biography:
            following = paragraphs[biography][1] if biography < len(paragraphs) else ""
            check(
                following.startswith("Призначається"),
                f"на місці білого зображення немає порожнього рядка (далі «{following[:30]}»)",
            )

        # ---- нерозривність ----
        section = index_of("§ 1")
        heading = next((i for i, text in paragraphs if text.endswith("ПРИЗНАЧИТИ:")), None)
        for name, position in (("§ 1", section), ("«ПРИЗНАЧИТИ:»", heading)):
            if position is None:
                check(False, f"{name} на місці")
                continue
            check(keep_with_next(position), f"{name} зчеплена з наступним")
            if position < len(paragraphs) and not paragraphs[position][1]:
                check(keep_with_next(position + 1), f"порожній абзац після {name} теж зчеплений")

        first_item = index_of("Полковника") or index_of("1. Полковника")
        second_item = index_of("Майора") or index_of("2. Майора")
        last_biography = index_of("Призначається")
        if first_item and second_item and last_biography:
            item_format = doc.Paragraphs(first_item).Range.ParagraphFormat
            check(bool(item_format.KeepTogether), "перший пункт цілісний")
            check(keep_with_next(first_item), f"перший пункт зчеплений зі своєю біографією ({label})")
            check(keep_with_next(second_item), f"другий пункт зчеплений зі своєю біографією ({label})")
            check(
                not keep_with_next(last_biography),
                f"біографія першого пункту НЕ чіпляється до другого пункту ({label})",
            )
        else:
            check(False, "обидва пункти й біографія на місці")
    finally:
        doc.Close(False)


def main() -> int:
    workdir = tempfile.mkdtemp(prefix="nat_blank_fonts_")
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        run_case(word, workdir, auto_numbered=False)
        run_case(word, workdir, auto_numbered=True)
    finally:
        try:
            word.Quit()
        except Exception:
            pass

    print(f"\nПідсумок: {checks - len(failures)} з {checks}")
    if failures:
        print("Не пройшли:")
        for item in failures:
            print(f"  • {item}")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
