"""Діагностика макета готового витягу: чи не відірвався § від того, що під ним.

Читає ГОТОВИЙ файл витягів (можна справжній) і друкує лише СТРУКТУРУ — номери
сторінок, прапорці нерозривності та мітки § . Жодного тексту наказу, ПІБ чи
номерів частин у вивід не потрапляє, тож результат безпечно показувати.

    python scripts/e2e_extracts/check_extract_layout.py "шлях\\до\\Витяги ….docx"

Для кожного § друкується:
  * сторінка самого § і сторінка НАСТУПНОГО непорожнього абзацу;
  * KeepWithNext / KeepTogether на абзаці § та на всьому проміжку до нього;
  * вердикт ВІДІРВАНО, якщо вони опинились на різних сторінках.

Якщо у вашого файлу § показує `keepwithnext=False` — зчеплення не застосувалось
(працює стара збірка або шапка не потрапила у список шапок витягу). Якщо
`keepwithnext=True`, а сторінки все одно різні — блок під § фізично не влазить,
і це вже випадок для ручної перевірки.

ВАЖЛИВО: файл не має бути відкритий у Word — інакше відкриття зависає без
помилки. Якщо він відкритий, закрийте його або зробіть копію.
"""
import os
import re
import sys

import win32com.client

sys.stdout.reconfigure(encoding="utf-8")

if len(sys.argv) < 2:
    print(__doc__)
    raise SystemExit(2)

path = os.path.abspath(sys.argv[1])
SECTION = re.compile(r"^\s*§")

word = win32com.client.DispatchEx("Word.Application")
word.Visible = False
word.DisplayAlerts = 0
try:
    doc = word.Documents.Open(path, ReadOnly=True)
    doc.Repaginate()
    rows = []
    for index, paragraph in enumerate(doc.Paragraphs, start=1):
        text = str(paragraph.Range.Text or "").strip("\r\x07 \t")
        paragraph_format = paragraph.Range.ParagraphFormat
        rows.append(
            {
                "index": index,
                "empty": not text,
                "section": bool(SECTION.match(text)),
                "label": text.split("\n")[0][:8] if SECTION.match(text) else "",
                "page": paragraph.Range.Information(3),  # wdActiveEndPageNumber
                "keep_next": bool(paragraph_format.KeepWithNext),
                "keep_together": bool(paragraph_format.KeepTogether),
            }
        )

    pages = doc.ComputeStatistics(2)
    print(f"файл: {os.path.basename(path)}")
    print(f"абзаців: {len(rows)}, сторінок: {pages}")

    problems = 0
    sections = [row for row in rows if row["section"]]
    if not sections:
        print("§ у документі не знайдено — перевіряти нічого.")
    for row in sections:
        following = next(
            (other for other in rows if other["index"] > row["index"] and not other["empty"]),
            None,
        )
        if following is None:
            continue
        gap = [
            other for other in rows
            if row["index"] <= other["index"] < following["index"]
        ]
        unbound = [other["index"] for other in gap if not other["keep_next"]]
        torn = row["page"] != following["page"]
        problems += bool(torn)
        print(
            f"{'ВІДІРВАНО' if torn else 'разом    '} {row['label']!r}: "
            f"§ стор.{row['page']} keepwithnext={row['keep_next']} "
            f"keeptogether={row['keep_together']} | наступний непорожній стор.{following['page']}"
            + (f" | без зчеплення абзаци: {unbound}" if unbound else "")
        )
    print("ПРОБЛЕМ:", problems)
    doc.Close(False)
finally:
    word.Quit()
