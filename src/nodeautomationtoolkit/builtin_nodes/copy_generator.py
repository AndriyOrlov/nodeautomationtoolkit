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
import time

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

        if resolve_span is not None:
            resolved = resolve_span(source_doc)
            if isinstance(resolved, dict):
                # Розширений варіант: межі + теги, похідні від наказу
                # (реквізити підписанта для заготовки).
                body_start, signer_index = resolved["span"]
                values = {**values, **resolved.get("values", {})}
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

        # Останній абзац змісту — підписант наказу: підкреслюємо його
        # й відсуваємо прізвище до правого краю.
        content_range = doc.Range(content_start, content_end)
        paragraph_count = content_range.Paragraphs.Count
        for i in range(paragraph_count, 0, -1):
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
