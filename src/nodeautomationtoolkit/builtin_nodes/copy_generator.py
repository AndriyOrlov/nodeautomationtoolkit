"""Генерація примірників наказу (№ 2/3) — окремий ізольований модуль.

Підхід: примірник **не є копією файлу наказу**. Він збирається із заготовки
(шаблону), у якій уже є шапка наказу, теги реквізитів і тег `{{зміст}}`.
У `{{зміст}}` переноситься тіло наказу разом із підписантом — так само, як у
витягах, копіюванням `FormattedText` кожного абзацу.

Чому саме так: копіювання файлу наказу тягнуло за собою колонтитули та решту
службового оформлення оригіналу. Колонтитули примірника будуються тут з нуля
(`apply_service_headers`) — у заготовці їх немає.

Модуль працює лише з об'єктами Word COM: шляхи, назви папок і обробку помилок
тримає викликач.
"""

from __future__ import annotations

import re
import time

# Коди констант Word, щоб не тягнути залежність від win32com.client.constants
_WD_GO_TO_PAGE = 1
_WD_GO_TO_ABSOLUTE = 1
_WD_STATISTIC_LINES = 1
_WD_STATISTIC_PAGES = 2
_WD_UNDERLINE_SINGLE = 1
_WD_COLOR_BLACK = 0
_WD_ALIGN_PARAGRAPH_CENTER = 1
_WD_LINE_SPACE_SINGLE = 0
_WD_COLLAPSE_START = 1
_WD_SECTION_BREAK_NEXT_PAGE = 2
_WD_FIELD_PAGE = 33
_WD_HEADER_FOOTER_PRIMARY = 1
_WD_HEADER_FOOTER_FIRST_PAGE = 2
_WD_PAGE_BREAK_CHARACTER = ""

# Гриф обмеження доступу у колонтитулах примірника.
SERVICE_MARK_TEXT = "ДЛЯ СЛУЖБОВОГО КОРИСТУВАННЯ"
_SERVICE_MARK_SIZE = 14.0
_PAGE_NUMBER_SIZE = 12.0
_HEADER_FONT_NAME = "Times New Roman"

# Початок службової таблиці розсилки/відміток — у примірник не переноситься.
_DISTRIBUTION_MARKERS = (
    "розрахунок розсилки",
    "таблиця розсилки",
    "список розсилки",
    "розсилка:",
    "відмітки служби діловодства",
    "служба діловодства",
    "розіслано:",
    "відмітка про виконання",
)

_ORDER_BODY_KEYWORDS = (
    "НАКАЗУЮ",
    "ПРИЗНАЧИТИ",
    "НАПРАВИТИ",
    "ВІДРЯДИТИ",
    "ЗВІЛЬНИТИ",
    "ВІЙСЬКОВОСЛУЖБОВЦІВ",
)

# Прізвище наприкінці рядка підписанта. Форми: «Ім'я ПРІЗВИЩЕ»,
# «І.П.ПРІЗВИЩЕ» / «І. П. Прізвище» або просто «ПРІЗВИЩЕ».
# Гілка з ініціалами має допускати прізвище і ВЕЛИКИМИ, і малими літерами,
# інакше «С.М.ПОПКО» обрізається до «ПОПКО».
_SIGNATURE_NAME_RE = re.compile(
    # два ініціали: «С.М.ПОПКО»
    r"(?:[А-ЯІЇЄҐ]\.\s?[А-ЯІЇЄҐ]\.\s?[А-ЯІЇЄҐ][А-Яа-яІіЇїЄєҐґ'’-]+"
    # один ініціал: «І. ПЕТРЕНКО» — інакше ініціал лишався б за межами хвоста
    r"|[А-ЯІЇЄҐ]\.\s?[А-ЯІЇЄҐ][А-Яа-яІіЇїЄєҐґ'’-]+"
    # повне ім'я: «Петро ПЕТРЕНКО»
    r"|[А-ЯІЇЄҐ][а-яіїєґ'’-]+\s+[А-ЯІЇЄҐ][А-ЯІЇЄҐ'’-]+"
    # саме лише прізвище
    r"|[А-ЯІЇЄҐ][А-ЯІЇЄҐ'’-]{2,})\s*$"
)


# Коди та ознаки тимчасової зайнятості Word. Такі збої не є помилкою даних —
# Word просто не встиг обробити попередній виклик, тому їх варто повторити.
_TRANSIENT_COM_CODES = (
    -2147418111,  # RPC_E_CALL_REJECTED — «Call was rejected by callee»
    -2147417846,  # RPC_E_SERVERCALL_RETRYLATER — сервер зайнятий
)
_TRANSIENT_COM_MARKERS = (
    "call was rejected",
    "rejected by callee",
    "server is busy",
    "retrylater",
)


def is_transient_word_error(error: Exception) -> bool:
    """Чи є збій тимчасовою зайнятістю Word (можна повторити)."""
    code = getattr(error, "hresult", None)
    args = getattr(error, "args", ()) or ()
    if code in _TRANSIENT_COM_CODES or (args and args[0] in _TRANSIENT_COM_CODES):
        return True
    text = str(error).casefold()
    if any(marker in text for marker in _TRANSIENT_COM_MARKERS):
        return True
    # Коли Word відхиляє виклик, pywin32 інколи не встигає розібрати об'єкт
    # і повідомляє про відсутній атрибут («Open.Content»).
    return isinstance(error, AttributeError)


def retry_on_busy_word(action, attempts: int = 3, delay: float = 1.5, log=None):
    """Повторює дію, якщо Word тимчасово відхилив виклик."""
    for attempt in range(1, attempts + 1):
        try:
            return action()
        except Exception as error:
            if attempt >= attempts or not is_transient_word_error(error):
                raise
            if log:
                log(f"  Word зайнятий ({error}); повтор {attempt + 1} з {attempts}…")
            time.sleep(delay * attempt)


def _paragraph_text(paragraph) -> str:
    return (paragraph.Range.Text or "").rstrip("\r\x07")


def find_signer_paragraph_index(doc) -> int | None:
    """Абзац підписанта — останній абзац із текстом ПЕРЕД останньою сторінкою.

    Останню сторінку наказу займає службова таблиця розсилки/відміток, тому
    підписант — це останній осмислений текст перед нею. Визначення позиційне,
    без розбору тексту, як і у витягах.

    Додатково спрацьовує запобіжник за текстовими маркерами: якщо таблиця
    розсилки почалась раніше за останню сторінку, вміст обрізається по ній —
    у примірнику таблиці розсилки бути не повинно.
    """
    doc.Repaginate()
    total_pages = doc.ComputeStatistics(_WD_STATISTIC_PAGES)
    if total_pages < 1:
        return None

    if total_pages == 1:
        boundary = doc.Content.End
    else:
        boundary = doc.GoTo(_WD_GO_TO_PAGE, _WD_GO_TO_ABSOLUTE, total_pages).Start

    marker_start = find_distribution_start(doc, boundary)
    if marker_start is not None:
        boundary = marker_start

    for index in range(doc.Paragraphs.Count, 0, -1):
        paragraph = doc.Paragraphs(index)
        if paragraph.Range.Start >= boundary:
            continue
        if _paragraph_text(paragraph).strip():
            return index
    return None


def find_distribution_start(doc, boundary: int) -> int | None:
    """Позиція початку службової таблиці розсилки, якщо вона є до `boundary`."""
    for index in range(1, doc.Paragraphs.Count + 1):
        paragraph = doc.Paragraphs(index)
        if paragraph.Range.Start >= boundary:
            break
        clean = _paragraph_text(paragraph).strip().casefold()
        if not clean:
            continue
        if any(clean.startswith(marker) or clean == marker for marker in _DISTRIBUTION_MARKERS):
            return paragraph.Range.Start
    return None


def find_body_start_paragraph_index(doc) -> int:
    """Перший абзац тіла наказу (§, пронумерований пункт або розпорядча фраза).

    Шапка наказу вже є в заготовці, тому переносимо лише тіло.
    """
    for index in range(1, doc.Paragraphs.Count + 1):
        clean = _paragraph_text(doc.Paragraphs(index)).strip()
        if not clean:
            continue
        if clean.startswith("§") or re.match(r"^\d+[\.\)]", clean):
            return index
        upper = clean.upper()
        if any(keyword in upper for keyword in _ORDER_BODY_KEYWORDS):
            return index
    return 1


def clear_headers_and_footers(doc) -> None:
    """Прибирає колонтитули, що прийшли із заготовки чи наказу.

    Власні колонтитули примірника (номер сторінки та гриф) ставляться після
    цього окремо — `apply_service_headers`.
    """
    for section in doc.Sections:
        for collection in (section.Headers, section.Footers):
            for item in collection:
                try:
                    if (item.Range.Text or "").strip("\r\x07 \t"):
                        item.Range.Delete()
                except Exception:
                    continue


def _style_header_paragraph(paragraph_range, size: float) -> None:
    """Приводить абзац колонтитула до потрібного вигляду.

    Стиль `Header`/`Footer` у шаблоні може мати власне вирівнювання, відступи,
    шрифт І КОЛІР, тому все виставляється ЯВНО — той самий підхід, що й у
    правилі 5.2.1 для тіла документа.

    Колір задається окремо й обов'язково: без нього колонтитул успадковував
    колір стилю заготовки й виходив ЧЕРВОНИМ. Ані `Name`, ані `Size` цього не
    перекривають — колір є самостійною властивістю шрифту.
    """
    font = paragraph_range.Font
    font.Name = _HEADER_FONT_NAME
    font.Size = size
    font.Bold = False
    font.Italic = False
    font.Underline = 0
    font.Color = _WD_COLOR_BLACK

    fmt = paragraph_range.ParagraphFormat
    fmt.Alignment = _WD_ALIGN_PARAGRAPH_CENTER
    fmt.LeftIndent = 0
    fmt.RightIndent = 0
    fmt.FirstLineIndent = 0
    fmt.SpaceBefore = 0
    fmt.SpaceAfter = 0
    fmt.LineSpacingRule = _WD_LINE_SPACE_SINGLE


def _set_story_lines(story, lines) -> None:
    """Робить у колонтитулі рівно стільки абзаців, скільки рядків у `lines`.

    Останній абзац story Word видалити не дає, і точний результат присвоєння
    `Range.Text` залежить від версії, тому кількість абзаців вирівнюється
    окремо — вставкою або видаленням, із запобіжником від нескінченного циклу.
    """
    story.Range.Text = "\r".join(lines)

    guard = len(lines) + 4
    while story.Range.Paragraphs.Count > len(lines) and guard > 0:
        guard -= 1
        before = story.Range.Paragraphs.Count
        try:
            story.Range.Paragraphs(before).Range.Delete()
        except Exception:
            break  # останній абзац story Word видаляти не дає
        if story.Range.Paragraphs.Count >= before:
            break  # видалення нічого не змінило — далі не намагаємось
    while story.Range.Paragraphs.Count < len(lines) and guard > 0:
        guard -= 1
        story.Range.InsertParagraphAfter()


def _style_story(story, sizes: dict, default: float) -> None:
    """Проходить УСІ наявні абзаци колонтитула й задає їм кегль.

    `sizes` — розмір для конкретних номерів абзаців, решта отримує `default`.
    Перебір іде за фактичною кількістю абзаців: звертатися до третього абзацу
    наосліп не можна, бо точний результат присвоєння `Range.Text` залежить від
    версії Word, а падіння тут коштувало б цілого примірника.
    """
    for number in range(1, story.Range.Paragraphs.Count + 1):
        _style_header_paragraph(
            story.Range.Paragraphs(number).Range, sizes.get(number, default)
        )


def _clear_story(item) -> None:
    """Спорожняє колонтитул і відв'язує його від попереднього розділу."""
    try:
        item.LinkToPrevious = False
    except Exception:
        pass
    item.Range.Text = ""


def isolate_last_page_section(doc) -> bool:
    """Виносить ОСТАННЮ сторінку в окремий розділ.

    Word уміє «іншу першу сторінку» (`DifferentFirstPageHeaderFooter`), але
    «іншої останньої» в ньому НЕМАЄ ЗОВСІМ. Єдиний спосіб лишити останню
    сторінку без колонтитулів — зробити її самостійним розділом і відв'язати
    його від попереднього.

    Повертає `True`, якщо остання сторінка після виклику є окремим розділом.
    """
    doc.Repaginate()
    pages = doc.ComputeStatistics(_WD_STATISTIC_PAGES)
    if pages < 2:
        return False

    start = doc.GoTo(_WD_GO_TO_PAGE, _WD_GO_TO_ABSOLUTE, pages).Start
    sections = doc.Sections
    if sections.Count > 1 and sections(sections.Count).Range.Start >= start:
        return True  # заготовка вже має окремий розділ на останню сторінку

    # Якщо сторінка починається РУЧНИМ розривом сторінки, його треба ЗАМІНИТИ
    # розривом розділу, а не додавати другий: два розриви підряд дають зайву
    # порожню сторінку.
    if start > 0:
        previous = doc.Range(start - 1, start)
        if (previous.Text or "") == _WD_PAGE_BREAK_CHARACTER:
            previous.Delete()
            start -= 1

    doc.Range(start, start).InsertBreak(_WD_SECTION_BREAK_NEXT_PAGE)

    # Розрив розділу вже переносить текст на нову сторінку. Якщо в абзаца ще й
    # стоїть «з нової сторінки», сторінок стає дві — одна з них порожня.
    sections = doc.Sections
    if sections.Count > 1:
        try:
            sections(sections.Count).Range.Paragraphs(1).Format.PageBreakBefore = False
        except Exception:
            pass

    doc.Repaginate()
    return doc.Sections.Count > 1


def apply_service_headers(doc, log=None) -> bool:
    """Колонтитули примірника: номер сторінки та гриф обмеження доступу.

    Розкладка (усе по центру, Times New Roman):

    * **верхній** — три абзаци: номер сторінки (12 pt), гриф (14 pt) і
      ПОРОЖНІЙ абзац, щоб текст наказу не зливався з колонтитулом;
    * **нижній** — лише гриф (14 pt).

    Ані на ПЕРШІЙ, ані на ОСТАННІЙ сторінці колонтитулів немає. Нумерація
    наскрізна й рахується з першої сторінки, тож друга сторінка — «2».
    У заготовці цього нічого немає — усе будує цей код.

    Повертає `True`, якщо колонтитули поставлено.
    """
    def note(message: str) -> None:
        if log:
            log(message)

    doc.Repaginate()
    pages_before = doc.ComputeStatistics(_WD_STATISTIC_PAGES)
    if pages_before < 2:
        # Одна сторінка є водночас першою й останньою, тож на ній не має бути
        # ні номера, ні грифа — колонтитули просто не ставляться.
        note("  Колонтитули: примірник на одну сторінку — ні номера, ні грифа.")
        return False

    isolate_last_page_section(doc)

    sections = doc.Sections
    total = sections.Count
    for index in range(1, total + 1):
        section = sections(index)
        setup = section.PageSetup
        # Чиста ПЕРША сторінка — це вміє сам Word.
        setup.DifferentFirstPageHeaderFooter = index == 1
        # Властивість зветься саме `OddAndEvenPagesHeaderFooter`. Ім'я
        # `DifferentOddAndEvenPagesHeaderFooter` (за аналогією з першою
        # сторінкою) у Word НЕ ІСНУЄ — воно валило генерацію КОЖНОГО наказу
        # помилкою «Property ... can not be set». Сама вимкненість парних
        # колонтитулів — лише запобіжник проти налаштувань заготовки, тому
        # відмову тут переживаємо, але не мовчки.
        try:
            setup.OddAndEvenPagesHeaderFooter = False
        except Exception as error:
            note(f"  УВАГА: не вдалося вимкнути парні колонтитули: {error}")

        header = section.Headers(_WD_HEADER_FOOTER_PRIMARY)
        footer = section.Footers(_WD_HEADER_FOOTER_PRIMARY)
        try:
            header.LinkToPrevious = False
            footer.LinkToPrevious = False
        except Exception:
            pass

        # Нумерація наскрізна: перезапуск лише в першому розділі, з одиниці.
        # Кожна властивість окремо: для ПЕРШОГО розділу Word може відхилити
        # `RestartNumberingAtSection` (перед ним немає розділу), і спільний
        # `try` тоді проковтнув би ще й `StartingNumber`.
        numbers = header.PageNumbers
        try:
            numbers.RestartNumberingAtSection = index == 1
        except Exception as error:
            note(f"  УВАГА: нумерація розділу {index}: {error}")
        if index == 1:
            try:
                numbers.StartingNumber = 1
            except Exception as error:
                note(f"  УВАГА: початок нумерації: {error}")

        if index == total and total > 1:
            # Останній розділ — це остання сторінка, вона лишається чистою.
            _clear_story(header)
            _clear_story(footer)
            continue

        _set_story_lines(header, ["", SERVICE_MARK_TEXT, ""])
        number_range = header.Range.Paragraphs(1).Range
        number_range.Collapse(_WD_COLLAPSE_START)
        doc.Fields.Add(number_range, _WD_FIELD_PAGE, "", False)
        # Розміри — за НОМЕРОМ абзацу, а не наосліп: якщо Word віддав інакшу
        # кількість абзаців, зайвий рядок гірший за помилку, але падати через
        # звернення до неіснуючого абзацу примірник не має.
        _style_story(header, {1: _PAGE_NUMBER_SIZE}, default=_SERVICE_MARK_SIZE)

        _set_story_lines(footer, [SERVICE_MARK_TEXT])
        _style_story(footer, {}, default=_SERVICE_MARK_SIZE)

        if index == 1:
            _clear_story(section.Headers(_WD_HEADER_FOOTER_FIRST_PAGE))
            _clear_story(section.Footers(_WD_HEADER_FOOTER_FIRST_PAGE))

    doc.Repaginate()
    pages_after = doc.ComputeStatistics(_WD_STATISTIC_PAGES)
    if pages_after > pages_before:
        # Найгірший наслідок цієї операції — зайва порожня сторінка від
        # подвійного розриву, тому про розбіжність треба сказати вголос.
        note(
            f"  УВАГА: після поділу на розділи сторінок стало {pages_after} "
            f"замість {pages_before}."
        )
    note(
        f"  Колонтитули: номер сторінки (12 пт) і гриф «{SERVICE_MARK_TEXT}» "
        "(14 пт); перша й остання сторінки чисті."
    )
    return True


def find_signature_name_tail(text: str) -> str:
    """Повертає прізвище (з ініціалами чи ім'ям) наприкінці рядка підпису."""
    match = _SIGNATURE_NAME_RE.search((text or "").rstrip())
    return match.group(0).strip() if match else ""


def push_tail_to_right_edge(doc, paragraph_range, tail: str, limit: int = 400) -> int:
    """Відсуває хвіст рядка (прізвище) до самого правого краю пробілами.

    Таблицею цього зробити не вдається, тому підбирається максимальна
    кількість пробілів, за якої рядок ще не переноситься.

    Пошук ДВІЙКОВИЙ, а не послідовний: кожне вимірювання — це звернення до
    Word, і послідовне додавання по одному пробілу давало сотні звернень на
    кожен підпис. За великого пакета Word від такого навантаження починав
    відхиляти виклики («Call was rejected by callee»). Двійковий пошук
    вкладається приблизно у 9 вимірювань.

    Кількість рядків беремо з `ComputeStatistics`, а не з `Information`, яка
    в прихованому Word кидає помилку.

    Повертає кількість доданих пробілів.
    """
    text = (paragraph_range.Text or "").rstrip("\r\x07")
    position = text.rfind(tail)
    if not tail or position <= 0:
        return 0

    insert_at = paragraph_range.Start + position
    base_lines = paragraph_range.ComputeStatistics(_WD_STATISTIC_LINES)
    current = 0

    def set_spaces(count: int) -> int:
        """Лишає рівно `count` пробілів перед хвостом і повертає кількість рядків."""
        nonlocal current
        if count > current:
            doc.Range(insert_at, insert_at).InsertBefore(" " * (count - current))
        elif count < current:
            doc.Range(insert_at, insert_at + (current - count)).Delete()
        current = count
        return paragraph_range.ComputeStatistics(_WD_STATISTIC_LINES)

    low, high, best = 0, limit, 0
    while low <= high:
        middle = (low + high) // 2
        if set_spaces(middle) > base_lines:
            high = middle - 1          # рядок уже перенісся — забагато
        else:
            best = middle
            low = middle + 1

    if current != best:
        set_spaces(best)
    return best


def format_signature_line(doc, paragraph_range, underline: bool) -> None:
    """Ставить прізвище під правий край і за потреби підкреслює рядок.

    У таблиці прізвище вже стоїть у своїй комірці, тому пробілами його не
    відсуваємо — інакше текст поїхав би за межу комірки.
    """
    in_table = False
    try:
        in_table = bool(paragraph_range.Information(12))  # wdWithInTable
    except Exception:
        in_table = False

    if not in_table:
        text = (paragraph_range.Text or "").rstrip("\r\x07")
        tail = find_signature_name_tail(text)
        if tail:
            # У наказі звання та прізвище часто розділені табуляціями й
            # крапками-заповнювачами. Замінюємо цей проміжок одним пробілом,
            # інакше положення прізвища визначали б табуляції, а не пробіли.
            position = text.rfind(tail)
            head = text[:position].rstrip(" \t.·…_")
            gap_start = paragraph_range.Start + len(head)
            gap_end = paragraph_range.Start + position
            if gap_end > gap_start:
                doc.Range(gap_start, gap_end).Delete()
                doc.Range(gap_start, gap_start).InsertBefore(" ")
            push_tail_to_right_edge(doc, paragraph_range, tail)
    if underline:
        paragraph_range.Font.Underline = _WD_UNDERLINE_SINGLE


def apply_keep_together_rules(doc, content_start: int, content_end: int) -> None:
    """Нерозривність блоків — так само, як у витягах.

    Пункт разом зі своїм біографічним блоком не розривається між сторінками,
    а шапки зчеплені з наступним пунктом.

    Підписант (останній непорожній абзац) НЕ приклеюється до попереднього
    пункту: інакше ланцюг «останній пункт + біографія + підписант» стає надто
    довгим, не вміщується і Word переносить його цілком, лишаючи порожнє місце
    та відриваючи підписанта на окрему сторінку.
    """
    content_range = doc.Range(content_start, content_end)
    spans = [
        (content_range.Paragraphs(i).Range.Start, content_range.Paragraphs(i).Range.End)
        for i in range(1, content_range.Paragraphs.Count + 1)
    ]

    def kind_of(text: str) -> str:
        clean = (text or "").strip()
        if not clean:
            return "blank"
        if clean.startswith("§"):
            return "heading"
        if re.match(r"^\d{1,3}(?:\.\d{1,3})*[\.\)]\s", clean):
            return "item"
        if clean.endswith(":"):
            return "heading"
        return "continuation"

    kinds = [kind_of(doc.Range(start, end).Text) for start, end in spans]

    # Останній непорожній абзац — підписант наказу.
    signer_index = next(
        (i for i in range(len(kinds) - 1, -1, -1) if kinds[i] != "blank"), None
    )

    for index, (start, end) in enumerate(spans):
        if kinds[index] == "blank":
            continue
        paragraph_format = doc.Range(start, end).ParagraphFormat
        paragraph_format.KeepTogether = True

        next_meaningful = next(
            (j for j in range(index + 1, len(kinds)) if kinds[j] != "blank"), None
        )
        if next_meaningful is None:
            paragraph_format.KeepWithNext = False
        elif next_meaningful == signer_index:
            # Підписант починає власний блок і не тягне за собою весь
            # останній пункт із біографією.
            paragraph_format.KeepWithNext = False
        elif kinds[index] == "heading":
            paragraph_format.KeepWithNext = True
        else:
            paragraph_format.KeepWithNext = kinds[next_meaningful] == "continuation"


def copy_order_body(doc, tag_range, source_doc, first_para: int, last_para: int) -> tuple[int, int]:
    """Переносить абзаци наказу у примірник разом із форматуванням.

    Повертає межі вставленого змісту `(початок, кінець)`.
    """
    total = source_doc.Paragraphs.Count
    if not 1 <= first_para <= last_para <= total:
        raise ValueError(f"некоректний діапазон абзаців наказу {first_para}-{last_para} із {total}")

    insert_point = tag_range.Paragraphs(1).Range.Start
    tag_range.Paragraphs(1).Range.Delete()
    content_start = insert_point

    # Копіюємо діапазон ОДНІЄЮ дією, а не абзац за абзацом: поабзацне
    # копіювання не відтворює таблиці, і блок підписанта, оформлений
    # таблицею (звання ліворуч, прізвище праворуч), втрачався. Заодно це
    # один виклик COM замість сотень — Word менше навантажений.
    source_range = source_doc.Range(
        source_doc.Paragraphs(first_para).Range.Start,
        source_doc.Paragraphs(last_para).Range.End,
    )
    destination = doc.Range(insert_point, insert_point)
    destination.FormattedText = source_range.FormattedText
    content_end = destination.End

    # Ручні розриви сторінок з наказу у примірник не переносяться.
    copied = doc.Range(content_start, content_end)
    for index in range(1, copied.Paragraphs.Count + 1):
        try:
            copied.Paragraphs(index).Range.ParagraphFormat.PageBreakBefore = False
        except Exception:
            continue

    return content_start, content_end


def remove_trailing_empty_page(doc) -> bool:
    """Видаляє останню сторінку, якщо на ній немає тексту.

    Деякі накази мають порожню останню сторінку; тягнути її у примірник не
    потрібно. Сторінка з текстом (зокрема службова «остання сторінка»
    заготовки) не чіпається.
    """
    doc.Repaginate()
    pages_before = doc.ComputeStatistics(_WD_STATISTIC_PAGES)
    if pages_before < 2:
        return False

    last_page_start = doc.GoTo(_WD_GO_TO_PAGE, _WD_GO_TO_ABSOLUTE, pages_before).Start
    tail_text = doc.Range(last_page_start, doc.Content.End).Text or ""
    if tail_text.strip("\r\x07\x0c \t\n "):
        return False  # на останній сторінці є текст — лишаємо як є

    removed = False
    while doc.Paragraphs.Count > 1:
        last_paragraph = doc.Paragraphs(doc.Paragraphs.Count)
        if (last_paragraph.Range.Text or "").strip("\r\x07\x0c \t "):
            break
        last_paragraph.Range.Delete()
        removed = True
        doc.Repaginate()
        if doc.ComputeStatistics(_WD_STATISTIC_PAGES) < pages_before:
            break
    return removed


def format_certifier_block(doc) -> bool:
    """Форматує блок засвідчувача, що починається з «Згідно з оригіналом».

    Блок суцільний — порожніх абзаців усередині немає, тому його кінцем є
    перший порожній абзац. Останній рядок (звання та прізвище) підкреслюється
    повністю, а прізвище відсувається до правого краю.
    """
    finder = doc.Content.Find
    finder.Text = "Згідно з оригіналом"
    if not finder.Execute():
        return False

    start_index = doc.Range(0, finder.Parent.Start).Paragraphs.Count

    # Між підписантом і блоком засвідчувача порожніх абзаців бути не повинно.
    while start_index > 1:
        previous = doc.Paragraphs(start_index - 1).Range
        if (previous.Text or "").strip("\r\x07 \t"):
            break
        previous.Delete()
        start_index -= 1

    total = doc.Paragraphs.Count
    last_index = start_index
    for index in range(start_index, total + 1):
        if not (doc.Paragraphs(index).Range.Text or "").strip("\r\x07 \t"):
            break
        last_index = index

    format_signature_line(doc, doc.Paragraphs(last_index).Range, underline=True)

    # Підписант і «Згідно з оригіналом» — один неподільний блок: інакше
    # засвідчувач відривається на наступну сторінку (правило 5.4 AGENT.md).
    chain_start = start_index
    while chain_start > 1 and (doc.Paragraphs(chain_start - 1).Range.Text or "").strip("\r\x07 \t"):
        chain_start -= 1
    for index in range(chain_start, last_index + 1):
        paragraph_format = doc.Paragraphs(index).Range.ParagraphFormat
        paragraph_format.KeepTogether = True
        paragraph_format.KeepWithNext = index < last_index
    return True


def replace_tags(doc, values: dict[str, str]) -> None:
    """Підставляє значення тегів заготовки."""
    for tag, value in values.items():
        find_obj = doc.Content.Find
        find_obj.Text = tag
        iterations = 0
        while find_obj.Execute() and iterations < 100:
            iterations += 1
            find_obj.Parent.Text = value
            find_obj = doc.Content.Find
            find_obj.Text = tag


def build_copy_document(
    word,
    order_path: str,
    working_copy_path: str,
    target_file: str,
    values: dict[str, str],
    resolve_span=None,
    log=None,
) -> int:
    """Збирає примірник із заготовки та зберігає його у `target_file`.

    `working_copy_path` — уже створена копія заготовки, яку можна змінювати.
    `resolve_span(source_doc) -> (перший_абзац, останній_абзац)` визначає межі
    тіла наказу разом із підписантом. Її передає викликач, щоб використати ту
    саму логіку пошуку підписанта, що й у витягах, а не дублювати її тут.

    Повертає кількість сторінок готового примірника.
    """
    def note(message: str) -> None:
        if log:
            log(message)

    source_doc = None
    doc = None
    try:
        source_doc = word.Documents.Open(order_path, ReadOnly=True)

        signature_line = ""
        if resolve_span is not None:
            resolved = resolve_span(source_doc)
            if isinstance(resolved, dict):
                # Розширений варіант: межі + теги, похідні від наказу
                # (блок підписанта та його реквізити для заготовки).
                body_start, signer_index = resolved["span"]
                values = {**values, **resolved.get("values", {})}
                signature_line = resolved.get("signature_line", "")
            else:
                body_start, signer_index = resolved
        else:
            # Запасний варіант, якщо межі не передали.
            signer_index = find_signer_paragraph_index(source_doc)
            body_start = find_body_start_paragraph_index(source_doc)
        if not signer_index or not body_start:
            raise ValueError("не вдалося визначити межі тіла наказу")
        if body_start > signer_index:
            raise ValueError("тіло наказу не знайдено перед підписантом")
        tail_preview = _paragraph_text(source_doc.Paragraphs(signer_index))[:40]
        note(
            f"  Тіло наказу: абзаци {body_start}–{signer_index}; "
            f"останній рядок: «{tail_preview}…»"
        )

        doc = word.Documents.Open(working_copy_path, ReadOnly=False)
        replace_tags(doc, values)

        find_obj = doc.Content.Find
        find_obj.Text = "{{зміст}}"
        if not find_obj.Execute():
            raise ValueError("у заготовці немає тегу {{зміст}}")

        content_start, content_end = copy_order_body(
            doc, find_obj.Parent, source_doc, body_start, signer_index
        )
        apply_keep_together_rules(doc, content_start, content_end)

        if signature_line:
            # Підписант підставлений окремим тегом — форматуємо його останній
            # рядок (звання та прізвище), а не останній пункт наказу.
            finder = doc.Content.Find
            finder.Text = signature_line
            if finder.Execute():
                format_signature_line(doc, finder.Parent.Paragraphs(1).Range, underline=True)
                note("  Рядок підписанта оформлено.")
            else:
                note("  УВАГА: рядок підписанта у документі не знайдено.")
        else:
            # Підписант — частина змісту: форматуємо його останній абзац.
            content_range = doc.Range(content_start, content_end)
            for i in range(content_range.Paragraphs.Count, 0, -1):
                paragraph_range = content_range.Paragraphs(i).Range
                if (paragraph_range.Text or "").strip("\r\x07 \t"):
                    format_signature_line(doc, paragraph_range, underline=True)
                    break

        # Блок засвідчувача («Згідно з оригіналом» + посада + звання/прізвище)
        # іде суцільно, без порожніх абзаців. Форматуємо його ОСТАННІЙ рядок.
        if format_certifier_block(doc):
            note("  Блок «Згідно з оригіналом» оформлено.")
        else:
            note("  УВАГА: блок «Згідно з оригіналом» у документі не знайдено.")

        clear_headers_and_footers(doc)

        # Деякі накази мають порожню останню сторінку — у примірнику її не лишаємо.
        if remove_trailing_empty_page(doc):
            note("  Прибрано порожню останню сторінку.")

        # ТІЛЬКИ ПІСЛЯ видалення порожньої сторінки: колонтитули залежать від
        # того, яка сторінка є останньою, тож ставити їх раніше не можна.
        apply_service_headers(doc, log=log)

        doc.Repaginate()
        pages = doc.ComputeStatistics(_WD_STATISTIC_PAGES)
        doc.SaveAs2(target_file, 16)  # wdFormatXMLDocument
        return pages
    finally:
        for document in (doc, source_doc):
            if document is not None:
                try:
                    document.Close(False)
                except Exception:
                    pass
