# -*- coding: utf-8 -*-
"""Наскрізна збірка примірника на СПРАВЖНЬОМУ Word (модуль 1).

Перевіряє те, що не видно з фейків:

* нерозривність (`KeepTogether` / `KeepWithNext`) у перенесеному змісті —
  зокрема коли підписант винесений в окремий тег і в змісті його НЕМАЄ;
* табличний варіант підписанта: `{{звання_підписанта}}` та
  `{{прізвище_підписанта}}` в окремих комірках, вирівняних до різних країв;
* перед «Згідно з оригіналом» немає порожнього абзаца;
* жирні ЦИФРИ дня в даті та номера наказу — і тільки вони.

Наказ і заготовка створюються тут-таки, дані ВИГАДАНІ: жодного справжнього
наказу, жодного шляху на `E:`. Запуск:

    .venv/Scripts/python.exe scripts/e2e_copies/check_copy_layout.py
"""

from __future__ import annotations

import os
import sys
import tempfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "..", "src"))
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", ".."))

for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

import win32com.client  # noqa: E402

from nodeautomationtoolkit.builtin_nodes.copy_generator import (  # noqa: E402
    build_copy_document,
)

_WD_FORMAT_XML = 16
_WD_PAGE_BREAK = 7
_WD_ALIGN_LEFT = 0
_WD_ALIGN_RIGHT = 2
_WD_COLOR_WHITE = 16777215
_WD_COLOR_BLACK = 0

failures: list[str] = []
checks = 0


def check(condition, description: str) -> None:
    global checks
    checks += 1
    if condition:
        print(f"  OK   {description}")
    else:
        print(f"  ЗБІЙ {description}")
        failures.append(description)


def _text(paragraph) -> str:
    return (paragraph.Range.Text or "").strip("\r\x07 \t")


def build_order(word, path: str) -> None:
    """Вигаданий наказ: шапка, два пункти з біографією, підписант, розсилка."""
    doc = word.Documents.Add()
    try:
        lines = [
            "МІНІСТЕРСТВО ОБОРОНИ ТЕСТОВОЇ КРАЇНИ",
            "НАКАЗ",
            "",
            "§ 1",
            "",
            "Відповідно до статті 1 нижчепойменованих військовослужбовців ПРИЗНАЧИТИ:",
            "",
            "1. Полковника ТЕСТОВОГО Тест Тестовича, командира підрозділу.",
            "1970 р. н., освіта: вища у 2000 р.",
            "1234567890.",
            "Призначається на вищу посаду.",
            "",
            "2. Майора ДРУГОГО Друг Другович, заступника командира.",
            "1980 р. н., освіта: вища у 2005 р.",
            "0987654321.",
            "Призначається на вищу посаду.",
            "",
            "Командир тестової частини",
            "полковник                    Тест ТЕСТЕНКО",
        ]
        for line in lines:
            doc.Content.InsertAfter(line + "\r")
        # У наказі трапляється БІЛИЙ текст — на екрані невидимий. У примірнику
        # він має стати звичайним чорним.
        white = doc.Content.Find
        white.Text = "1234567890."
        if white.Execute():
            white.Parent.Font.Color = _WD_COLOR_WHITE

        # Службовий хвіст на окремій сторінці — у примірник не має потрапити.
        doc.Paragraphs(doc.Paragraphs.Count).Range.InsertBreak(_WD_PAGE_BREAK)
        doc.Content.InsertAfter("Розрахунок розсилки витягів із наказу:\r")
        doc.Content.InsertAfter("1. Військова частина А0000   п. 1.\r")
        doc.SaveAs2(path, _WD_FORMAT_XML)
    finally:
        doc.Close(False)


def build_template(word, path: str) -> None:
    """Заготовка: реквізити, {{зміст}}, ТАБЛИЧНИЙ підписант, засвідчувач."""
    doc = word.Documents.Add()
    try:
        for line in (
            "МІНІСТЕРСТВО ОБОРОНИ ТЕСТОВОЇ КРАЇНИ",
            "НАКАЗ",
            "{{дата_наказу}}                    {{номер_наказу}}",
            "",
            "{{зміст}}",
            "",
            "Командир тестової частини",
        ):
            doc.Content.InsertAfter(line + "\r")

        # Табличний підписант: звання ліворуч, прізвище праворуч. Пробілами
        # двох різних вирівнювань не зробити — тому саме таблиця.
        anchor = doc.Paragraphs(doc.Paragraphs.Count).Range
        table = doc.Tables.Add(anchor, 1, 2)
        table.Borders.Enable = False
        table.Cell(1, 1).Range.Text = "{{звання_підписанта}}"
        table.Cell(1, 1).Range.ParagraphFormat.Alignment = _WD_ALIGN_LEFT
        table.Cell(1, 2).Range.Text = "{{прізвище_підписанта}}"
        table.Cell(1, 2).Range.ParagraphFormat.Alignment = _WD_ALIGN_RIGHT

        # Порожній абзац ПЕРЕД засвідчувачем — його має прибрати генератор.
        doc.Content.InsertAfter("\r")
        doc.Content.InsertAfter("{{згідно_з_оригіналом}}\r")
        doc.Content.InsertAfter("\r")
        doc.Content.InsertAfter("Виконавець: Тестовий Т.Т.\r")

        doc.Paragraphs(doc.Paragraphs.Count).Range.InsertBreak(_WD_PAGE_BREAK)
        doc.Content.InsertAfter("Відмітки служби діловодства\r")
        doc.SaveAs2(path, _WD_FORMAT_XML)
    finally:
        doc.Close(False)


def main() -> int:
    workdir = tempfile.mkdtemp(prefix="nat_layout_")
    order = os.path.join(workdir, "Наказ № 413 від 17.08.2026.docx")
    template = os.path.join(workdir, "template.docx")
    working = os.path.join(workdir, "_work.docx")
    result = os.path.join(workdir, "result.docx")

    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        build_order(word, order)
        build_template(word, template)

        import shutil

        shutil.copy2(template, working)

        # Межі тіла та реквізити підписанта — так само, як їх рахує App.
        source = word.Documents.Open(order, ReadOnly=True)
        try:
            texts = [
                (source.Paragraphs(i).Range.Text or "").replace(chr(7), "").strip()
                for i in range(1, source.Paragraphs.Count + 1)
            ]
            body_start = texts.index("§ 1") + 1
            signer_start = texts.index("Командир тестової частини") + 1
            last_paragraph = signer_start + 1
            body_end = signer_start - 1
            while body_end > body_start and not texts[body_end - 1]:
                body_end -= 1
            signer_lines = [
                texts[i - 1] for i in range(signer_start, last_paragraph + 1) if texts[i - 1]
            ]
        finally:
            source.Close(False)

        values = {
            "{{дата_наказу}}": "“17” серпня 2026 року",
            "{{номер_наказу}}": "№413",
            "{{згідно_з_оригіналом}}": "Згідно з оригіналом\rНачальник штабу\rполковник Іван ІВАНЕНКО",
        }

        def resolve(source_doc):
            return {
                "span": (body_start, body_end),
                "values": {
                    "{{підписант}}": "\r".join(signer_lines),
                    "{{звання_підписанта}}": "полковник",
                    "{{прізвище_підписанта}}": "Тест ТЕСТЕНКО",
                },
                "signature_line": "",
                "signer_in_tag": True,
            }

        pages = build_copy_document(
            word, order, working, result, values, resolve_span=resolve, log=print
        )
        print(f"\nПримірник зібрано, сторінок: {pages}\n")

        doc = word.Documents.Open(result, ReadOnly=False)
        try:
            paragraphs = [
                (i, _text(doc.Paragraphs(i))) for i in range(1, doc.Paragraphs.Count + 1)
            ]

            # ---- нерозривність ----
            items = [i for i, text in paragraphs if text.startswith(("1.", "2."))]
            check(len(items) == 2, f"обидва пункти на місці (знайдено {len(items)})")
            for index in items:
                fmt = doc.Paragraphs(index).Range.ParagraphFormat
                check(
                    bool(fmt.KeepTogether),
                    f"пункт у абзаці {index}: KeepTogether увімкнено",
                )
                check(
                    bool(fmt.KeepWithNext),
                    f"пункт у абзаці {index} зчеплений зі своєю біографією",
                )

            heading = next(
                (i for i, text in paragraphs if text.endswith("ПРИЗНАЧИТИ:")), None
            )
            check(heading is not None, "шапка розділу перенесена")
            if heading:
                check(
                    bool(doc.Paragraphs(heading).Range.ParagraphFormat.KeepWithNext),
                    "шапка зчеплена з наступним пунктом",
                )

            last_item = items[-1] if items else None
            if last_item:
                check(
                    bool(doc.Paragraphs(last_item).Range.ParagraphFormat.KeepWithNext),
                    "ОСТАННІЙ пункт не втратив зчеплення (підписант у тезі, "
                    "тож останнім підписантом він не є)",
                )

            # ---- табличний підписант ----
            check(doc.Tables.Count >= 1, f"таблиця підписанта є (таблиць: {doc.Tables.Count})")
            if doc.Tables.Count:
                table = doc.Tables(1)
                left = (table.Cell(1, 1).Range.Text or "").strip("\r\x07 ")
                right = (table.Cell(1, 2).Range.Text or "").strip("\r\x07 ")
                check(left == "полковник", f"звання в лівій комірці (там «{left}»)")
                check(
                    right == "Тест ТЕСТЕНКО", f"прізвище в правій комірці (там «{right}»)"
                )
                check(
                    table.Cell(1, 1).Range.ParagraphFormat.Alignment == _WD_ALIGN_LEFT,
                    "звання вирівняне ліворуч",
                )
                check(
                    table.Cell(1, 2).Range.ParagraphFormat.Alignment == _WD_ALIGN_RIGHT,
                    "прізвище вирівняне праворуч",
                )
                check(
                    "{{звання_підписанта}}" not in (doc.Content.Text or ""),
                    "теги табличного підписанта підставлено, а не лишено як текст",
                )

            # ---- засвідчувач ----
            certifier = next(
                (i for i, text in paragraphs if text.startswith("Згідно з оригіналом")), None
            )
            check(certifier is not None, "блок «Згідно з оригіналом» на місці")
            if certifier and certifier > 1:
                previous = doc.Paragraphs(certifier - 1).Range
                before = _text(doc.Paragraphs(certifier - 1))
                # Службовий абзац таблиці порожнім абзацом НЕ Є: це сама
                # структура рядка, і видаляти його не можна.
                in_table = chr(7) in (previous.Text or "")
                check(
                    bool(before) or in_table,
                    "перед «Згідно з оригіналом» немає порожнього абзаца "
                    f"(там {'кінець рядка таблиці' if in_table else f'«{before}»'})",
                )

            # ---- жирні цифри ----
            date_par = next((i for i, text in paragraphs if "серпня 2026" in text), None)
            check(date_par is not None, "рядок дати знайдено")
            if date_par:
                line = doc.Paragraphs(date_par).Range
                start = line.Start
                text = line.Text or ""
                day = text.find("17")
                bold_day = doc.Range(start + day, start + day + 2).Font.Bold
                check(bool(bold_day), "цифри дня в даті — жирні")

                month = text.find("серпня")
                bold_month = doc.Range(start + month, start + month + 6).Font.Bold
                check(not bold_month, "назва місяця — НЕ жирна")

                number = text.find("413")
                if number >= 0:
                    check(
                        bool(doc.Range(start + number, start + number + 3).Font.Bold),
                        "номер наказу — жирний",
                    )
                    sign = text.find("№")
                    check(
                        not doc.Range(start + sign, start + sign + 1).Font.Bold,
                        "знак «№» — НЕ жирний",
                    )

            # ---- білий текст ----
            hidden = next((i for i, t in paragraphs if t.startswith("1234567890")), None)
            check(hidden is not None, "білий рядок наказу перенесено")
            if hidden:
                check(
                    doc.Paragraphs(hidden).Range.Font.Color == _WD_COLOR_BLACK,
                    "білий текст наказу став ЧОРНИМ (фактично "
                    f"{doc.Paragraphs(hidden).Range.Font.Color})",
                )

            # ---- засвідчувач тримається підписанта ----
            if certifier and doc.Tables.Count:
                signer_table = doc.Tables(1)
                check(
                    not bool(signer_table.Rows.AllowBreakAcrossPages),
                    "рядок таблиці підписанта не розривається між сторінками",
                )
                check(
                    bool(signer_table.Range.Paragraphs(1).Range.ParagraphFormat.KeepWithNext),
                    "таблиця підписанта зчеплена з тим, що йде далі",
                )
                position = next(
                    (i for i, t in paragraphs if t == "Командир тестової частини"), None
                )
                if position:
                    check(
                        bool(doc.Paragraphs(position).Range.ParagraphFormat.KeepWithNext),
                        "посада підписанта зчеплена з наступним рядком",
                    )

            # ---- службовий хвіст ----
            check(
                "Розрахунок розсилки" not in (doc.Content.Text or ""),
                "таблиця розсилки наказу в примірник не потрапила",
            )
        finally:
            doc.Close(False)
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
