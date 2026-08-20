"""Генерація примірників наказу (№ 2/3) — окремий ізольований модуль.

Підхід: примірник **не є копією файлу наказу**. Він збирається із заготовки
(шаблону), у якій уже є шапка наказу, теги реквізитів і тег `{{зміст}}`.
У `{{зміст}}` переноситься тіло наказу разом із підписантом — так само, як у
витягах, копіюванням `FormattedText` кожного абзацу.

Чому саме так: копіювання файлу наказу тягнуло за собою колонтитули та решту
службового оформлення, а в примірниках колонтитулів бути не повинно.

Модуль працює лише з об'єктами Word COM: шляхи, назви папок і обробку помилок
тримає викликач.
"""

from __future__ import annotations

import re

# Коди констант Word, щоб не тягнути залежність від win32com.client.constants
_WD_GO_TO_PAGE = 1
_WD_GO_TO_ABSOLUTE = 1
_WD_STATISTIC_LINES = 1
_WD_STATISTIC_PAGES = 2
_WD_UNDERLINE_SINGLE = 1

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
    r"(?:[А-ЯІЇЄҐ]\.\s?[А-ЯІЇЄҐ]\.\s?[А-ЯІЇЄҐ][А-Яа-яІіЇїЄєҐґ'’-]+"
    r"|[А-ЯІЇЄҐ][а-яіїєґ'’-]+\s+[А-ЯІЇЄҐ][А-ЯІЇЄҐ'’-]+"
    r"|[А-ЯІЇЄҐ][А-ЯІЇЄҐ'’-]{2,})\s*$"
)


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
    """Прибирає всі колонтитули: у примірниках їх бути не повинно."""
    for section in doc.Sections:
        for collection in (section.Headers, section.Footers):
            for item in collection:
                try:
                    if (item.Range.Text or "").strip("\r\x07 \t"):
                        item.Range.Delete()
                except Exception:
                    continue


def find_signature_name_tail(text: str) -> str:
    """Повертає прізвище (з ініціалами чи ім'ям) наприкінці рядка підпису."""
    match = _SIGNATURE_NAME_RE.search((text or "").rstrip())
    return match.group(0).strip() if match else ""


def push_tail_to_right_edge(doc, paragraph_range, tail: str, limit: int = 400) -> int:
    """Відсуває хвіст рядка (прізвище) до самого правого краю пробілами.

    Таблицею цього зробити не вдається, тому пробіли додаються, доки рядок не
    почне переноситись, після чого остання порція знімається. Кількість рядків
    беремо з `ComputeStatistics`, а не з `Information`, яка в прихованому Word
    кидає помилку.

    Повертає кількість доданих пробілів.
    """
    text = (paragraph_range.Text or "").rstrip("\r\x07")
    position = text.rfind(tail)
    if not tail or position <= 0:
        return 0

    insert_at = paragraph_range.Start + position
    base_lines = paragraph_range.ComputeStatistics(_WD_STATISTIC_LINES)
    added = 0

    for batch in (8, 1):
        while added + batch <= limit:
            doc.Range(insert_at, insert_at).InsertBefore(" " * batch)
            added += batch
            if paragraph_range.ComputeStatistics(_WD_STATISTIC_LINES) > base_lines:
                doc.Range(insert_at, insert_at + batch).Delete()
                added -= batch
                break
    return added


def format_signature_line(doc, paragraph_range, underline: bool) -> None:
    """Ставить прізвище під правий край і за потреби підкреслює рядок."""
    tail = find_signature_name_tail((paragraph_range.Text or "").rstrip("\r\x07"))
    if tail:
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

    for index in range(first_para, last_para + 1):
        source_range = source_doc.Paragraphs(index).Range.Duplicate
        if "\x0c" in (source_range.Text or ""):
            continue  # ручні розриви сторінок з наказу не переносимо
        destination = doc.Range(insert_point, insert_point)
        destination.FormattedText = source_range.FormattedText
        insert_point = destination.End

    return content_start, insert_point


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

        if resolve_span is not None:
            body_start, signer_index = resolve_span(source_doc)
        else:
            # Запасний варіант, якщо межі не передали.
            signer_index = find_signer_paragraph_index(source_doc)
            body_start = find_body_start_paragraph_index(source_doc)
        if not signer_index or not body_start:
            raise ValueError("не вдалося визначити межі тіла наказу")
        if body_start > signer_index:
            raise ValueError("тіло наказу не знайдено перед підписантом")
        note(f"  Тіло наказу: абзаци {body_start}–{signer_index} (останній — підписант).")

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

        # Останній абзац змісту — підписант наказу: підкреслюємо його
        # й відсуваємо прізвище до правого краю.
        content_range = doc.Range(content_start, content_end)
        paragraph_count = content_range.Paragraphs.Count
        for i in range(paragraph_count, 0, -1):
            paragraph_range = content_range.Paragraphs(i).Range
            if (paragraph_range.Text or "").strip("\r\x07 \t"):
                format_signature_line(doc, paragraph_range, underline=True)
                break

        # Прізвище засвідчувача теж має стояти під правим краєм.
        for tag in ("{{згідно_з_оригіналом}}", "{{засвідчення}}"):
            value = values.get(tag)
            if not value:
                continue
            finder = doc.Content.Find
            finder.Text = value
            if finder.Execute():
                format_signature_line(doc, finder.Parent.Paragraphs(1).Range, underline=False)

        clear_headers_and_footers(doc)

        # Деякі накази мають порожню останню сторінку — у примірнику її не лишаємо.
        if remove_trailing_empty_page(doc):
            note("  Прибрано порожню останню сторінку.")

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
