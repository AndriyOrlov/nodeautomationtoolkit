"""Друк пакета: витяги, витяги до управління, примірники № 2 і № 3 по черзі.

Правила від користувача (26.10.2026):

- принтер друкує лише з одного боку, але має ручну подачу; аркуші виходять
  лицем донизу — останній надрукований аркуш лежить ЗВЕРХУ стосу;
- витяги — 2 сторінки на аркуш; витяги до управління — за бажанням, 1 на аркуш;
- примірник: усі аркуші односторонні, лише останній аркуш має дві сторінки:
  передостанню і задню (останню сторінку файлу). Порядок (05.11.2026):
  сторінки 1…n−2 → задня сторінка n → пауза «верхній аркуш у ручну подачу» →
  сторінка n−1 на його звороті. Так стос після перевертання одразу лежить
  як треба;
- примірник № 3 ідентичний № 2, лише на першій сторінці замість
  «Для службового користування / Прим. № 2» — «Прим. № 3». Файл № 3 не
  створюється: у відкритому № 2 число міняється лише для друку, документ
  закривається без збереження;
- після КОЖНОГО кроку — пауза (Пробіл/Enter): надруковане треба відкласти.

Принтер з АВТОМАТИЧНИМ двобічним друком (05.11.2026, напр. KYOCERA FS-1025MFP):
останній аркуш примірника (передостання + задня сторінка) друкується двобічно
одним кроком, без паузи на ручну подачу; решта аркушів — однобічно, як і
раніше. Двобічність вмикається в налаштуваннях вибраного принтера лише на цей
крок і повертається назад (`printer_duplex`).

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
    def key(self) -> str:
        """Ключ наказу для вибору галочкою в черзі друку."""
        return f"{self.number}|{self.date}"

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
    #: Друкувати двобічно (лише з автоматичним дуплексом принтера).
    duplex: bool = False
    #: Пауза після кроку. Ні — коли наступний крок продовжує той самий примірник.
    pause_after: bool = True
    #: Друк на ЗВОРОТІ вже надрукованого аркуша — з лотка для звороту (05.11.2026).
    back_side: bool = False
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


def scan_batch_folders(folders: list[str], depth: int = 2) -> list[OrderFiles]:
    """Те саме, що `scan_batch_folder`, але з кількох тек (05.11.2026).

    Примірники № 2/3 і витяги можуть лежати в різних теках: файли одного наказу
    зводяться за номером і датою з назви. Перша знайдена тека з файлом виграє.
    """
    merged: dict[tuple[str, str], OrderFiles] = {}
    seen: set[str] = set()
    for folder in folders:
        key_folder = os.path.normcase(os.path.abspath(folder)) if folder else ""
        if not folder or not os.path.isdir(folder) or key_folder in seen:
            continue
        seen.add(key_folder)
        for found in scan_batch_folder(folder, depth):
            entry = merged.setdefault((found.number, found.date), OrderFiles(found.number, found.date))
            entry.extracts = entry.extracts or found.extracts
            entry.management = entry.management or found.management
            entry.copy_two = entry.copy_two or found.copy_two
    return sorted(merged.values(), key=lambda order: _sort_key(order.number, order.date))


def _sort_key(number: str, date: str) -> tuple:
    day, month, year = (re.findall(r"\d+", date) + ["0", "0", "0"])[:3]
    digits = re.match(r"\d+", number or "")
    return (int(year or 0), int(month or 0), int(day or 0), int(digits.group(0)) if digits else 0, number)


def _auto_duplex_copy_steps(title: str, path: str, pages: int, number: int, copy_three: bool) -> list[PrintStep]:
    """Примірник на принтері з автоматичним дуплексом: 1…n−2 однобічно, n−1/n — двобічно."""
    steps: list[PrintStep] = []
    if pages >= 3:
        shown = "сторінка 1" if pages == 3 else f"сторінки 1–{pages - 2}"
        steps.append(
            PrintStep(
                title, f"Примірник № {number} — {shown}", path,
                pages="1" if pages == 3 else f"1-{pages - 2}", copy_three=copy_three, pause_after=False,
            )
        )
    steps.append(
        PrintStep(
            title, f"Примірник № {number} — останній аркуш двобічно (сторінки {pages - 1} і {pages})", path,
            pages=f"{pages - 1}-{pages}", copy_three=copy_three, duplex=True,
        )
    )
    return steps


def plan_steps(
    order: OrderFiles,
    copy_pages: int,
    include_management: bool,
    feed_hint: str = "",
    auto_duplex: bool = False,
    back_tray_name: str = "",
    include_extracts: bool = True,
    include_copies: bool = True,
) -> list[PrintStep]:
    """Кроки для одного наказу. `copy_pages` — сторінок у примірнику № 2.

    `auto_duplex` — принтер сам друкує з двох боків: без ручної подачі.
    `back_tray_name` — лоток, у який кладеться аркуш для звороту (для підказки в паузі).
    """
    title = order.title
    hint = f" ({feed_hint})" if feed_hint else ""
    feed_place = f"у лоток «{back_tray_name}»" if back_tray_name else "у ручну подачу"
    steps: list[PrintStep] = []
    if include_extracts and order.extracts:
        steps.append(PrintStep(title, "Витяги — 2 сторінки на аркуш", order.extracts, two_per_sheet=True))
    if include_management and order.management:
        # Витяги до управління — 2 примірники, кожен витяг на окремому аркуші (26.10.2026).
        steps.append(PrintStep(title, "Витяги до управління — 2 примірники", order.management, copies=2))
    if include_copies and order.copy_two:
        for copy_three, number in ((False, 2), (True, 3)):
            if auto_duplex and copy_pages >= 2:
                steps.extend(_auto_duplex_copy_steps(title, order.copy_two, copy_pages, number, copy_three))
            elif copy_pages >= 2:
                # Користувач, 05.11.2026: задня сторінка друкується ПЕРЕД
                # передостанньою, а передостання — на її звороті.
                if copy_pages >= 3:
                    fronts = "1" if copy_pages == 3 else f"1-{copy_pages - 2}"
                    shown = "сторінка 1" if copy_pages == 3 else f"сторінки 1–{copy_pages - 2}"
                    steps.append(
                        PrintStep(
                            title,
                            f"Примірник № {number} — {shown}",
                            order.copy_two,
                            pages=fronts,
                            copy_three=copy_three,
                            # Задня сторінка йде одразу слідом з того самого лотка —
                            # пауза перед нею зайва (користувач, 06.11.2026).
                            pause_after=False,
                        )
                    )
                steps.append(
                    PrintStep(
                        title,
                        f"Примірник № {number} — задня сторінка {copy_pages}",
                        order.copy_two,
                        pages=str(copy_pages),
                        copy_three=copy_three,
                        after=f"Візьміть ВЕРХНІЙ аркуш з лотка виходу й покладіть {feed_place}{hint}.",
                    )
                )
                steps.append(
                    PrintStep(
                        title,
                        f"Примірник № {number} — сторінка {copy_pages - 1} на звороті задньої",
                        order.copy_two,
                        pages=str(copy_pages - 1),
                        copy_three=copy_three,
                        back_side=True,
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


def set_paper_tray(doc, tray: int) -> None:
    """Лоток паперу для друку цього документа (усі розділи); 0 — як у принтера.

    Документ відкрито лише для читання й закривається без збереження — файл не змінюється.
    """
    if not tray:
        return
    for section in doc.Sections:
        section.PageSetup.FirstPageTray = tray
        section.PageSetup.OtherPagesTray = tray


def printer_trays(name: str) -> list[tuple[int, str]]:
    """Лотки паперу принтера: (номер для Word, назва з драйвера)."""
    try:
        import win32print

        handle = win32print.OpenPrinter(name)
        try:
            port = win32print.GetPrinter(handle, 2)["pPortName"]
        finally:
            win32print.ClosePrinter(handle)
        bins = list(win32print.DeviceCapabilities(name, port, 6))  # DC_BINS
        names = list(win32print.DeviceCapabilities(name, port, 12))  # DC_BINNAMES
        return [(int(number), str(title).strip()) for number, title in zip(bins, names)]
    except Exception:
        return []


# ── «2 сторінки на аркуш» — складаємо самі (13.11.2026) ─────────────────────
#
# Раніше це робив Word масштабуванням під час друку (PrintZoomColumn=2). Драйвери
# типу V4 (Canon LBP6030 V4 тощо) це ігнорують — виходило 1 сторінка на аркуш.
# Тепер програма сама складає тимчасовий альбомний документ: на кожному аркуші
# дві сторінки витягу (векторні знімки сторінок Word, EMF) поруч. Він друкується
# 1:1 на будь-якому принтері. Документ ніде не зберігається.
_WD_VIEW_PRINT = 3
_WD_PAGE_BREAK = 7
_WD_RELATIVE_PAGE = 1
_WD_WRAP_FRONT = 3


def _page_images(doc, folder: str) -> list[str]:
    """Знімки всіх сторінок документа (EMF) у теці `folder`."""
    window = doc.Windows(1)
    window.View.Type = _WD_VIEW_PRINT
    pages = window.Panes(1).Pages
    paths = []
    for index in range(1, pages.Count + 1):
        path = os.path.join(folder, f"page{index:04d}.emf")
        with open(path, "wb") as handle:
            handle.write(bytes(pages(index).EnhMetaFileBits))
        paths.append(path)
    return paths


def build_two_up_document(word, doc, folder: str):
    """Новий документ Word: сторінки `doc` по дві на альбомному аркуші."""
    setup = doc.PageSetup
    page_width, page_height = float(setup.PageWidth), float(setup.PageHeight)
    images = _page_images(doc, folder)
    if not images:
        raise RuntimeError("у документі немає сторінок")
    sheet = word.Documents.Add()
    layout = sheet.PageSetup
    layout.Orientation = 1  # альбомна
    layout.PageWidth, layout.PageHeight = page_height, page_width
    layout.TopMargin = layout.BottomMargin = layout.LeftMargin = layout.RightMargin = 0
    layout.HeaderDistance = layout.FooterDistance = 0
    sheets = (len(images) + 1) // 2
    for _ in range(sheets - 1):
        tail = sheet.Content
        tail.Collapse(0)
        tail.InsertBreak(_WD_PAGE_BREAK)
    half = page_height / 2
    scale = min(half / page_width, page_width / page_height)
    width, height = page_width * scale, page_height * scale
    for index, image in enumerate(images):
        anchor = sheet.Paragraphs(index // 2 + 1).Range
        shape = sheet.Shapes.AddPicture(image, False, True, 0, 0, width, height, anchor)
        shape.WrapFormat.Type = _WD_WRAP_FRONT
        shape.RelativeHorizontalPosition = _WD_RELATIVE_PAGE
        shape.RelativeVerticalPosition = _WD_RELATIVE_PAGE
        shape.LockAspectRatio = 0
        shape.Width, shape.Height = width, height
        shape.Left = half * (index % 2) + (half - width) / 2
        shape.Top = (page_width - height) / 2
    return sheet


def _print_two_up(word, doc, step: PrintStep, pdf_path: str, tray: int) -> None:
    import shutil
    import tempfile

    folder = tempfile.mkdtemp(prefix="nat_2up_")
    sheet = None
    try:
        sheet = build_two_up_document(word, doc, folder)
        if pdf_path:
            sheet.ExportAsFixedFormat(
                OutputFileName=pdf_path, ExportFormat=_WD_EXPORT_PDF, OpenAfterExport=False,
                Range=_WD_EXPORT_ALL,
            )
            return
        set_paper_tray(sheet, tray)
        sheet.PrintOut(
            Background=False, Copies=max(1, int(step.copies)), Collate=True, Range=_WD_PRINT_ALL,
            Item=0, PageType=0, PrintToFile=False, ManualDuplexPrint=False,
        )
    finally:
        if sheet is not None:
            sheet.Close(False)
        shutil.rmtree(folder, ignore_errors=True)


def print_step(word, step: PrintStep, pdf_path: str = "", tray: int = 0) -> None:
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
        if step.two_per_sheet and not step.pages:
            _print_two_up(word, doc, step, pdf_path, tray)
            return
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
        set_paper_tray(doc, tray)
        # Явно все, що Word інакше бере зі своїх запамʼятованих налаштувань діалогу
        # друку: лише непарні/парні, ручний двобічний, у файл (18.11.2026).
        options = {
            "Background": False, "Copies": max(1, int(step.copies)), "Collate": True,
            "Item": 0, "PageType": 0, "PrintToFile": False, "ManualDuplexPrint": False,
        }
        if step.pages:
            options.update(Range=_WD_PRINT_PAGES, Pages=absolute_pages_spec(doc, step.pages))
        else:
            options.update(Range=_WD_PRINT_ALL)
        doc.PrintOut(**options)
    finally:
        doc.Close(False)


_WD_INFO_ADJUSTED_PAGE, _WD_INFO_SECTION = 1, 2


def absolute_pages_spec(doc, pages: str) -> str:
    """«1-2», «4» (фізичні сторінки) → «p1s1-p2s1», «p4s2» для `PrintOut(Pages=…)`.

    Word у `Pages` читає НАДРУКОВАНІ номери сторінок. У примірнику кілька розділів
    (остання сторінка — окремий), а в наказі бувають розділи з власною нумерацією:
    тоді «1» збігається з кількома сторінками й друкується двічі, а «3» — не та
    сторінка (18.11.2026). Адреса «p<номер>s<розділ>» однозначна.
    """
    first, last = _page_bounds(pages)
    if not first:
        return pages

    def one(number: int) -> str:
        rng = doc.GoTo(_WD_GOTO_PAGE, _WD_GOTO_ABSOLUTE, number)
        return f"p{int(rng.Information(_WD_INFO_ADJUSTED_PAGE))}s{int(rng.Information(_WD_INFO_SECTION))}"

    return one(first) if first == last else f"{one(first)}-{one(last)}"


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


def list_printers() -> list[str]:
    """Принтери Windows (локальні й мережеві), за абеткою."""
    try:
        import win32print

        flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
        return sorted({str(entry[2]) for entry in win32print.EnumPrinters(flags)}, key=str.casefold)
    except Exception:
        return []


def printer_supports_duplex(name: str) -> bool:
    """Чи драйвер заявляє двобічний друк. Довіряти не можна: так кажуть і принтери
    з ручним дуплексом, тому режим вмикає користувач, а це — лише підказка."""
    try:
        import win32print

        return win32print.DeviceCapabilities(name, "", 7) == 1  # DC_DUPLEX
    except Exception:
        return False


_DMDUP_SIMPLEX, _DMDUP_LONG_EDGE = 1, 2


class printer_duplex:
    """Тимчасово вмикає/вимикає двобічний друк у налаштуваннях КОРИСТУВАЧА для принтера.

    Змінюється лише власний (per-user, рівень 9) DEVMODE, не системний; після
    `restore()` значення повертається. Word підхоплює зміну після повторного
    вибору принтера (`select_printer`).
    """

    def __init__(self, name: str):
        self.name = name
        self._original: int | None = None

    def _devmode(self, handle):
        import win32print

        user = win32print.GetPrinter(handle, 9)
        devmode = user.get("pDevMode") if user else None
        return devmode or win32print.GetPrinter(handle, 2)["pDevMode"]

    def set(self, on: bool) -> None:
        import win32print

        handle = win32print.OpenPrinter(self.name, {"DesiredAccess": win32print.PRINTER_ACCESS_USE})
        try:
            devmode = self._devmode(handle)
            if self._original is None:
                self._original = int(devmode.Duplex or _DMDUP_SIMPLEX)
            devmode.Duplex = _DMDUP_LONG_EDGE if on else _DMDUP_SIMPLEX
            devmode.Fields |= 0x1000  # DM_DUPLEX
            win32print.SetPrinter(handle, 9, {"pDevMode": devmode}, 0)
        finally:
            win32print.ClosePrinter(handle)

    def restore(self) -> None:
        if self._original is None:
            return
        import win32print

        handle = win32print.OpenPrinter(self.name, {"DesiredAccess": win32print.PRINTER_ACCESS_USE})
        try:
            devmode = self._devmode(handle)
            devmode.Duplex = self._original
            win32print.SetPrinter(handle, 9, {"pDevMode": devmode}, 0)
        finally:
            win32print.ClosePrinter(handle)
        self._original = None


def select_printer(word, name: str) -> None:
    """Принтер для Word БЕЗ зміни принтера Windows за замовчуванням."""
    if not name:
        return
    # `WordBasic.FilePrintSetup(DoNotSetAsSysDefault=1)` через pywin32 не викликається
    # (іменованих аргументів у WordBasic немає), а `ActivePrinter` Word може зробити
    # принтером Windows за замовчуванням — тож одразу повертаємо попередній.
    import win32print

    before = win32print.GetDefaultPrinter()
    try:
        word.ActivePrinter = name
    finally:
        try:
            if before and win32print.GetDefaultPrinter() != before:
                win32print.SetDefaultPrinter(before)
        except Exception:
            pass


def reload_printer(word, name: str) -> None:
    """Змушує Word перечитати налаштування принтера (після зміни двобічного друку).

    Word тримає налаштування принтера весь сеанс; повторний вибір ТОГО САМОГО
    принтера їх не перечитує, тому — коротко на інший принтер і назад.
    """
    if not name:
        return
    other = next((printer for printer in list_printers() if printer != name), "")
    if other:
        try:
            select_printer(word, other)
        except Exception:
            pass
    select_printer(word, name)


def default_printer() -> str:
    try:
        import win32print

        return str(win32print.GetDefaultPrinter())
    except Exception:
        return ""
