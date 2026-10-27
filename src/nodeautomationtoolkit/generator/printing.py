"""Друк пакета: витяги, витяги до управління, примірники № 2 і № 3 по черзі.

Правила від користувача (26.10.2026):

- принтер друкує лише з одного боку, але має ручну подачу; аркуші виходять
  лицем донизу — останній надрукований аркуш лежить ЗВЕРХУ стосу;
- витяги — 2 сторінки на аркуш; витяги до управління — за бажанням, 1 на аркуш;
- примірник: усі аркуші односторонні, а задня сторінка (остання сторінка
  файлу) друкується на ЗВОРОТІ передостанньої сторінки. Тому: лицьові
  сторінки 1…n−1 → пауза «верхній аркуш у ручну подачу» → сторінка n;
- примірник № 3 ідентичний № 2, лише на першій сторінці замість
  «Для службового користування / Прим. № 2» — «Прим. № 3». Файл № 3 не
  створюється: у відкритому № 2 число міняється лише для друку, документ
  закривається без збереження;
- після КОЖНОГО кроку — пауза (Пробіл/Enter): надруковане треба відкласти.

Черга будується з файлів у теці пакета (номер і дата — з назви файлу, як
скрізь у генераторі). Друк іде на принтер Windows за замовчуванням через Word;
пробний режим — PDF-файли кроків замість паперу.
"""

from __future__ import annotations

import os
import re
from dataclasses import dataclass, field

from .naming import extract_metadata_from_filename, is_generated_copy_filename

EXTRACTS_PREFIX = "Витяги наказу"
MANAGEMENT_PREFIX = "Витяги до управління за наказом"

#: Word: wdPrintAllDocument, wdPrintRangeOfPages, wdExportFormatPDF, wdExportAllDocument, wdExportFromTo.
_WD_PRINT_ALL, _WD_PRINT_PAGES = 0, 4
_WD_EXPORT_PDF, _WD_EXPORT_ALL, _WD_EXPORT_FROM_TO = 17, 0, 3
_WD_STAT_PAGES = 2
_WD_GOTO_PAGE, _WD_GOTO_ABSOLUTE = 1, 1

_SERVICE_MARK_RE = re.compile(r"службового\s+користування", re.IGNORECASE)
_COPY_TWO_RE = re.compile(r"Прим\.?[ \t ]*№[ \t ]*2(?!\d)")


def copy_sheet_count(pages: int) -> int:
    """Аркушів на ОДИН примірник: усі аркуші односторонні, лише задня сторінка
    (остання) — на звороті передостанньої. 4 сторінки → 3 аркуші."""
    pages = int(pages or 0)
    return pages - 1 if pages >= 2 else max(pages, 0)


@dataclass
class OrderFiles:
    """Файли одного наказу в теці пакета."""

    number: str
    date: str
    extracts: str = ""
    management: str = ""
    copy_two: str = ""

    @property
    def title(self) -> str:
        parts = [f"№ {self.number}" if self.number else "", f"від {self.date}" if self.date else ""]
        return "Наказ " + " ".join(part for part in parts if part).strip() if any(parts) else "Наказ"


@dataclass
class PrintStep:
    """Один крок друку; після нього — пауза."""

    order: str
    what: str
    path: str
    pages: str = ""  # «1-4», «5»; порожньо — увесь документ
    two_per_sheet: bool = False
    copies: int = 1
    copy_three: bool = False
    #: Що сказати в паузі ПІСЛЯ кроку (крім загального «відкладіть»).
    after: str = ""
    notes: list[str] = field(default_factory=list)


def _order_key(path: str) -> tuple[str, str]:
    number, date = extract_metadata_from_filename(os.path.basename(path))
    return (number or "", date or "")


def scan_batch_folder(folder: str, depth: int = 2) -> list[OrderFiles]:
    """Знаходить у теці (і підтеках до `depth`) файли наказів для друку.

    Тимчасові файли Word («~$…») і файли без номера й дати пропускаються:
    без них неможливо звести витяги з примірником того самого наказу.
    """
    orders: dict[tuple[str, str], OrderFiles] = {}
    root_depth = os.path.abspath(folder).rstrip(os.sep).count(os.sep)
    for current, folders, files in os.walk(folder):
        if os.path.abspath(current).count(os.sep) - root_depth >= depth:
            folders[:] = []
        for name in sorted(files):
            if name.startswith("~$") or not name.lower().endswith(".docx"):
                continue
            path = os.path.join(current, name)
            key = _order_key(path)
            if not any(key):
                continue
            entry = orders.setdefault(key, OrderFiles(*key))
            if name.startswith(MANAGEMENT_PREFIX):
                entry.management = entry.management or path
            elif name.startswith(EXTRACTS_PREFIX):
                entry.extracts = entry.extracts or path
            elif is_generated_copy_filename(name):
                entry.copy_two = entry.copy_two or path
    return sorted(orders.values(), key=lambda order: _sort_key(order.number, order.date))


def _sort_key(number: str, date: str) -> tuple:
    day, month, year = (re.findall(r"\d+", date) + ["0", "0", "0"])[:3]
    digits = re.match(r"\d+", number or "")
    return (int(year or 0), int(month or 0), int(day or 0), int(digits.group(0)) if digits else 0, number)


def plan_steps(order: OrderFiles, copy_pages: int, include_management: bool, feed_hint: str = "") -> list[PrintStep]:
    """Кроки для одного наказу. `copy_pages` — сторінок у примірнику № 2."""
    title = order.title
    hint = f" ({feed_hint})" if feed_hint else ""
    steps: list[PrintStep] = []
    if order.extracts:
        steps.append(PrintStep(title, "Витяги — 2 сторінки на аркуш", order.extracts, two_per_sheet=True))
    if include_management and order.management:
        # Витяги до управління — 2 примірники, кожен витяг на окремому аркуші (26.10.2026).
        steps.append(PrintStep(title, "Витяги до управління — 2 примірники", order.management, copies=2))
    if order.copy_two:
        for copy_three, number in ((False, 2), (True, 3)):
            if copy_pages >= 2:
                fronts = "1" if copy_pages == 2 else f"1-{copy_pages - 1}"
                steps.append(
                    PrintStep(
                        title,
                        f"Примірник № {number} — лицьові сторінки 1–{copy_pages - 1}",
                        order.copy_two,
                        pages=fronts,
                        copy_three=copy_three,
                        after=f"Візьміть ВЕРХНІЙ аркуш з лотка виходу й вкладіть у ручну подачу{hint}.",
                    )
                )
                steps.append(
                    PrintStep(
                        title,
                        f"Примірник № {number} — задня сторінка (на зворот)",
                        order.copy_two,
                        pages=str(copy_pages),
                        copy_three=copy_three,
                    )
                )
            else:
                steps.append(
                    PrintStep(title, f"Примірник № {number}", order.copy_two, copy_three=copy_three)
                )
    return steps


# ─────────────────────────────────────────────────────────────────────────────
# Word
# ─────────────────────────────────────────────────────────────────────────────
def page_count(doc) -> int:
    doc.Repaginate()
    return int(doc.ComputeStatistics(_WD_STAT_PAGES))


def _first_page_range(doc):
    try:
        if page_count(doc) > 1:
            second = doc.GoTo(_WD_GOTO_PAGE, _WD_GOTO_ABSOLUTE, 2)
            return doc.Range(0, int(second.Start))
    except Exception:
        pass
    return doc.Content


def _replace_in_range(rng) -> bool:
    """«…службового користування … Прим. № 2» → «… № 3» у межах діапазону."""
    text = str(rng.Text or "")
    for match in _COPY_TWO_RE.finditer(text):
        if not _SERVICE_MARK_RE.search(text[max(0, match.start() - 300):match.start()]):
            continue
        find = rng.Duplicate.Find
        find.ClearFormatting()
        find.Text = match.group(0)
        find.MatchCase = True
        find.Forward = True
        find.Wrap = 0  # wdFindStop
        if find.Execute():
            found = find.Parent
            digit = found.Duplicate
            digit.SetRange(found.End - 1, found.End)
            if str(digit.Text) == "2":
                digit.Text = "3"
                return True
    return False


def mark_copy_three(doc) -> bool:
    """Міняє «Прим. № 2» на «Прим. № 3» на першій сторінці (текст або колонтитули).

    Повертає False, якщо напису поруч із «Для службового користування» не знайдено:
    тоді № 3 друкувати не можна — вийшов би другий № 2.
    """
    if _replace_in_range(_first_page_range(doc)):
        return True
    try:
        section = doc.Sections(1)
        for kind in (2, 1, 3):  # першої сторінки, основний, парних
            header = section.Headers(kind)
            if header.Exists and _replace_in_range(header.Range):
                return True
    except Exception:
        pass
    return False


def print_step(word, step: PrintStep, pdf_path: str = "") -> None:
    """Друкує (або вивантажує в PDF) один крок. Документ закривається без збереження."""
    # Без `Visible=False`: на документі без вікна Word падав під час експорту
    # чи друку діапазону сторінок. Сам Word і так прихований.
    doc = word.Documents.Open(os.path.abspath(step.path), ReadOnly=True, AddToRecentFiles=False)
    try:
        if step.copy_three and not mark_copy_three(doc):
            raise RuntimeError(
                "на першій сторінці примірника не знайдено «Для службового користування / Прим. № 2» — "
                "№ 3 не друкую, щоб не вийшов другий № 2"
            )
        first, last = _page_bounds(step.pages)
        if pdf_path:
            if first:
                doc.ExportAsFixedFormat(
                    OutputFileName=pdf_path, ExportFormat=_WD_EXPORT_PDF, OpenAfterExport=False,
                    Range=_WD_EXPORT_FROM_TO, From=first, To=last,
                )
            else:
                doc.ExportAsFixedFormat(
                    OutputFileName=pdf_path, ExportFormat=_WD_EXPORT_PDF, OpenAfterExport=False,
                    Range=_WD_EXPORT_ALL,
                )
            return
        options = {"Background": False, "Copies": max(1, int(step.copies)), "Collate": True}
        if step.pages:
            options.update(Range=_WD_PRINT_PAGES, Pages=step.pages)
        else:
            options.update(Range=_WD_PRINT_ALL)
        if step.two_per_sheet:
            options.update(PrintZoomColumn=2, PrintZoomRow=1)
        doc.PrintOut(**options)
    finally:
        doc.Close(False)


def _page_bounds(pages: str) -> tuple[int, int]:
    if not pages:
        return 0, 0
    numbers = [int(part) for part in re.findall(r"\d+", pages)]
    return numbers[0], numbers[-1]


def pdf_name(index: int, step: PrintStep) -> str:
    name = re.sub(r'[\/:*?"<>|]+', "_", f"{index:02d} {step.order} — {step.what}")
    return name.strip(" .")[:150] + ".pdf"


#: Word упав або закрився — далі друкувати нічим.
_WORD_GONE_CODES = (-2147023174, -2147023170, -2147417848)


def word_is_gone(error: BaseException) -> bool:
    return bool(getattr(error, "args", None)) and error.args[0] in _WORD_GONE_CODES


def default_printer() -> str:
    try:
        import win32print

        return str(win32print.GetDefaultPrinter())
    except Exception:
        return ""
