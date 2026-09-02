# -*- coding: utf-8 -*-
"""Перевірка колонтитулів примірника на СПРАВЖНЬОМУ Word (правило 11.7).

Навіщо: розкладка колонтитулів будується десятком властивостей Word COM, і
описка в назві будь-якої з них валить генерацію кожного наказу («Property ...
can not be set»). Фейки в `tests/test_copy_service_headers.py` перевіряють
логіку, але назв властивостей знати не можуть — це робить лише сам Word.

Дані тут ВИГАДАНІ й створюються на місці: жодного наказу, жодного витягу,
жодного шляху на `E:`. Запуск:

    .venv/Scripts/python.exe scripts/e2e_copies/check_service_headers.py
"""

from __future__ import annotations

import os
import sys
import tempfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "..", "src"))

# Консоль Windows за замовчуванням не cp65001, а вивід тут українською.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

import win32com.client  # noqa: E402

from nodeautomationtoolkit.builtin_nodes.copy_generator import (  # noqa: E402
    SERVICE_MARK_TEXT,
    apply_service_headers,
    clear_headers_and_footers,
)

_WD_HEADER_FOOTER_PRIMARY = 1
_WD_HEADER_FOOTER_FIRST_PAGE = 2
_WD_STATISTIC_PAGES = 2
_WD_FORMAT_XML = 16
_WD_PAGE_BREAK = 7  # wdPageBreak
_WD_ALIGN_CENTER = 1
_WD_COLOR_RED = 255
_WD_COLOR_BLACK = 0
# Стилі беремо КОНСТАНТАМИ, а не назвами: у локалізованому Word назва стилю
# українська («Верхній колонтитул»), і звернення за "Header" впало б.
_WD_STYLE_HEADER = -32
_WD_STYLE_FOOTER = -33

PAGES = 4

failures: list[str] = []
checks = 0


def check(condition: bool, description: str) -> None:
    global checks
    checks += 1
    if condition:
        print(f"  OK   {description}")
    else:
        print(f"  ЗБІЙ {description}")
        failures.append(description)


def build_source(word, path: str) -> None:
    """Створює документ на кілька сторінок із ручними розривами сторінок."""
    doc = word.Documents.Add()
    try:
        for page in range(1, PAGES + 1):
            doc.Content.InsertAfter(f"Сторінка вигаданого тексту номер {page}.\r")
            if page < PAGES:
                doc.Content.InsertParagraphAfter()
                doc.Paragraphs(doc.Paragraphs.Count).Range.InsertBreak(_WD_PAGE_BREAK)
        # Відтворюємо симптом користувача: у стилях колонтитулів ЧЕРВОНИЙ колір.
        # Колір є самостійною властивістю шрифту, тому без явного присвоєння
        # він протікав у готовий примірник.
        doc.Styles(_WD_STYLE_HEADER).Font.Color = _WD_COLOR_RED
        doc.Styles(_WD_STYLE_FOOTER).Font.Color = _WD_COLOR_RED
        # Колонтитул, який має бути стертий перед побудовою власного.
        doc.Sections(1).Headers(_WD_HEADER_FOOTER_PRIMARY).Range.Text = "СТАРИЙ КОЛОНТИТУЛ"
        doc.SaveAs2(path, _WD_FORMAT_XML)
    finally:
        doc.Close(False)


def main() -> int:
    workdir = tempfile.mkdtemp(prefix="nat_headers_")
    source = os.path.join(workdir, "synthetic.docx")

    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        build_source(word, source)

        doc = word.Documents.Open(source, ReadOnly=False)
        try:
            doc.Repaginate()
            pages_before = doc.ComputeStatistics(_WD_STATISTIC_PAGES)
            print(f"Синтетичний документ: {pages_before} стор.\n")

            clear_headers_and_footers(doc)
            applied = apply_service_headers(doc, log=print)
            print()

            check(applied is True, "колонтитули поставлено")

            doc.Repaginate()
            pages_after = doc.ComputeStatistics(_WD_STATISTIC_PAGES)
            check(
                pages_after == pages_before,
                f"кількість сторінок не змінилась ({pages_before} -> {pages_after}); "
                "зайва порожня сторінка від подвійного розриву — головний ризик",
            )

            total = doc.Sections.Count
            check(total >= 2, f"остання сторінка стала окремим розділом (розділів: {total})")

            first = doc.Sections(1)
            # УВАГА: `PageSetup` віддає ці прапорці як VARIANT_BOOL — int -1/0,
            # а не Python-bool, тому перевіряти можна лише істинність.
            # (`LinkToPrevious` і `RestartNumberingAtSection` при цьому віддають
            # справжні bool — однакового правила в Word немає.)
            check(
                bool(first.PageSetup.DifferentFirstPageHeaderFooter),
                "у першого розділу увімкнено «інша перша сторінка»",
            )
            check(
                not bool(first.PageSetup.OddAndEvenPagesHeaderFooter),
                "парні колонтитули вимкнено (правильна назва властивості)",
            )

            body_header = first.Headers(_WD_HEADER_FOOTER_PRIMARY)
            lines = [
                (body_header.Range.Paragraphs(i).Range.Text or "").strip("\r\x07 \t")
                for i in range(1, body_header.Range.Paragraphs.Count + 1)
            ]
            check(
                body_header.Range.Paragraphs.Count == 3,
                f"у верхньому колонтитулі три абзаци (фактично {body_header.Range.Paragraphs.Count}): {lines}",
            )
            check(
                body_header.Range.Fields.Count == 1,
                f"номер сторінки вставлено полем PAGE (полів: {body_header.Range.Fields.Count})",
            )
            check(
                SERVICE_MARK_TEXT in (body_header.Range.Text or ""),
                f"у верхньому колонтитулі є гриф «{SERVICE_MARK_TEXT}»",
            )
            if body_header.Range.Paragraphs.Count >= 3:
                check(not lines[2], "третій абзац верхнього колонтитула порожній")
                sizes = [
                    body_header.Range.Paragraphs(i).Range.Font.Size for i in (1, 2, 3)
                ]
                check(sizes[0] == 12, f"номер сторінки 12 пт (фактично {sizes[0]})")
                check(sizes[1] == 14, f"гриф 14 пт (фактично {sizes[1]})")
                names = {
                    body_header.Range.Paragraphs(i).Range.Font.Name for i in (1, 2, 3)
                }
                check(names == {"Times New Roman"}, f"шрифт Times New Roman (фактично {names})")
                colors = [
                    body_header.Range.Paragraphs(i).Range.Font.Color for i in (1, 2, 3)
                ]
                check(
                    set(colors) == {_WD_COLOR_BLACK},
                    f"верхній колонтитул ЧОРНИЙ, попри червоний стиль джерела "
                    f"(фактично {colors})",
                )
                aligns = [
                    body_header.Range.Paragraphs(i).Range.ParagraphFormat.Alignment
                    for i in (1, 2, 3)
                ]
                check(
                    set(aligns) == {_WD_ALIGN_CENTER},
                    f"усе по центру (фактично {aligns})",
                )

            body_footer = first.Footers(_WD_HEADER_FOOTER_PRIMARY)
            check(
                SERVICE_MARK_TEXT in (body_footer.Range.Text or ""),
                "у нижньому колонтитулі є гриф",
            )
            check(
                body_footer.Range.Fields.Count == 0,
                "у нижньому колонтитулі немає номера сторінки",
            )
            check(
                body_footer.Range.Paragraphs(1).Range.Font.Size == 14,
                "гриф у нижньому колонтитулі 14 пт",
            )
            check(
                body_footer.Range.Paragraphs(1).Range.Font.Color == _WD_COLOR_BLACK,
                "нижній колонтитул ЧОРНИЙ, попри червоний стиль джерела "
                f"(фактично {body_footer.Range.Paragraphs(1).Range.Font.Color})",
            )

            check(
                not (first.Headers(_WD_HEADER_FOOTER_FIRST_PAGE).Range.Text or "").strip("\r\x07 \t"),
                "на першій сторінці верхній колонтитул порожній",
            )
            check(
                not (first.Footers(_WD_HEADER_FOOTER_FIRST_PAGE).Range.Text or "").strip("\r\x07 \t"),
                "на першій сторінці нижній колонтитул порожній",
            )

            last = doc.Sections(total)
            check(
                not bool(last.Headers(_WD_HEADER_FOOTER_PRIMARY).LinkToPrevious),
                "розділ останньої сторінки відв'язано від попереднього",
            )
            check(
                not (last.Headers(_WD_HEADER_FOOTER_PRIMARY).Range.Text or "").strip("\r\x07 \t"),
                "на останній сторінці верхній колонтитул порожній",
            )
            check(
                not (last.Footers(_WD_HEADER_FOOTER_PRIMARY).Range.Text or "").strip("\r\x07 \t"),
                "на останній сторінці нижній колонтитул порожній",
            )
            check(
                not bool(last.Headers(_WD_HEADER_FOOTER_PRIMARY).PageNumbers.RestartNumberingAtSection),
                "остання сторінка не перезапускає нумерацію",
            )
            check(
                first.Headers(_WD_HEADER_FOOTER_PRIMARY).PageNumbers.StartingNumber == 1,
                "нумерація починається з одиниці",
            )
            check(
                "СТАРИЙ КОЛОНТИТУЛ" not in (body_header.Range.Text or ""),
                "старий колонтитул джерела стерто",
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
