"""Позначена копія перевіреного наказу: виділення кольором і примітки Word.

Оригінал НЕ змінюється (PROJECT_RULES 1.4): документ відкривається лише для
читання, одразу зберігається під новою назвою, і всі позначки робляться в копії.
"""

from __future__ import annotations

import os

from .check import ERROR, Finding

#: WdColorIndex: червоний — помилка, жовтий — увага.
HIGHLIGHT = {ERROR: 6}
DEFAULT_HIGHLIGHT = 7
_FIND_LIMIT = 200  # Word шукає не довше 255 символів


def comment_text(finding: Finding) -> str:
    text = f"{finding.level.upper()} · {finding.rule}: {finding.what}."
    return f"{text} {finding.how}" if finding.how else text


def _find(document, quotes) -> object | None:
    for quote in quotes:
        needle = " ".join(str(quote or "").split())[:_FIND_LIMIT].strip()
        if len(needle) < 3:
            continue
        found = document.Content
        search = found.Find
        search.ClearFormatting()
        if search.Execute(
            FindText=needle.replace("^", "^^"),
            MatchCase=True,
            MatchWholeWord=False,
            MatchWildcards=False,
            Forward=True,
            Wrap=0,  # wdFindStop
        ):
            return found
    return None


def _narrow(found, focus: str):
    """Звужує знайдений фрагмент до `focus` (неправильне слово) у його межах.

    Фрагмент лише вказує місце в наказі — виділяти треба саме помилку, а не
    півабзацу. Не знайшлося — лишається весь фрагмент.
    """
    needle = " ".join(str(focus or "").split())[:_FIND_LIMIT].strip()
    if not needle:
        return found
    narrowed = found.Duplicate
    search = narrowed.Find
    search.ClearFormatting()
    if search.Execute(
        FindText=needle.replace("^", "^^"),
        MatchCase=False,
        MatchWholeWord=" " not in needle,
        MatchWildcards=False,
        Forward=True,
        Wrap=0,  # wdFindStop — лише в межах фрагмента
    ):
        return narrowed
    return found


def save_marked_copy(
    word, source_path: str, output_path: str, findings: list[Finding]
) -> tuple[int, int]:
    """Зберігає позначену копію. Повертає (позначено в тексті, примітки на початку документа)."""
    document = word.Documents.Open(
        os.path.abspath(source_path), ReadOnly=True, AddToRecentFiles=False
    )
    marked = unplaced = 0
    try:
        document.SaveAs2(os.path.abspath(output_path), 16)  # wdFormatXMLDocument
        # Помилки — останніми: на тому самому місці червоне перекриває жовте, а не навпаки.
        for finding in sorted(findings, key=lambda finding: finding.level == ERROR):
            found = _find(document, finding.quotes)
            if found is None:
                found = document.Range(0, 0)
                unplaced += 1
            else:
                found = _narrow(found, finding.focus)
                found.HighlightColorIndex = HIGHLIGHT.get(finding.level, DEFAULT_HIGHLIGHT)
                marked += 1
            document.Comments.Add(found, comment_text(finding))
        document.Save()
    finally:
        document.Close(False)
    return marked, unplaced
