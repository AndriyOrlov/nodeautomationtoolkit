"""Робота з документом Word: текст, абзаци, формат, запуск і закриття Word.

Винесено з `generate_extracts.py` без змін логіки (26.10.2026).
"""
from __future__ import annotations

import os
import re
import shutil
from nodeautomationtoolkit.builtin_nodes.recipient_mapping import (
    normalize_item_numbering,
)

from .errors import UserError


#: Лівий відступ біографічного блоку — від 4 см (зазвичай 6, буває 5–7), як у
#: перевірці наказу (`word_macro/NATCheckCore.bas`). 1 см = 72 / 2.54 pt.
BIO_MIN_LEFT_INDENT_PT = 4 * 72 / 2.54


def paragraph_line_count(paragraph_text: str) -> int:
    """Скільки рядків дає абзац у тексті `"\\n".join(абзаци).splitlines()`.

    ЄДИНЕ місце цього підрахунку — для тексту наказу й для карти «рядок →
    абзац» витягів і повідомлень. `\\f` (розрив сторінки) і `\\v` (мʼякий
    перенос) теж розривають рядок, і в КІНЦІ абзацу дають ще один рядок, якого
    `paragraph_text.splitlines()` не бачить. Карта рахувала саме так, тож після
    кожного такого абзацу межі пунктів зсувались на абзац: пункт втрачав
    біографію, а сусідній пункт захоплював чужий абзац (дубль).
    """
    return len((paragraph_text.rstrip("\r\x07") + "\n").splitlines())


def is_page_break_only_text(paragraph_text: str) -> bool:
    """Абзац, у якому, крім розриву сторінки, нічого немає."""
    text = paragraph_text or ""
    return "\x0c" in text and not text.strip("\r\x07\v\f \t")


def remove_page_breaks_in_range(doc, start: int, end: int) -> int:
    """Прибирає розриви сторінки/розділу з вставленого діапазону `doc`.

    Пагінацію витягу чи повідомлення будуємо заново, тож розрив наказу тут
    зайвий — але текст абзацу, в якому він стояв, лишається. Раніше такий
    абзац пропускали цілком, і зникав його текст (біографія на межі сторінок).
    Повертає новий кінець діапазону.
    """
    # Кінець беремо з самого діапазону, а не «мінус один»: Word при видаленні
    # розриву буває вставляє пробіл («розумне вирізання»), і рахунок
    # збивався на символ — наступний абзац вставлявся ПЕРЕД знаком абзацу
    # попереднього, а той втрачав відступ і порожній рядок після себе.
    container = doc.Range(start, end)
    for _ in range(100):
        if "\x0c" not in (container.Text or ""):
            break
        found = False
        for code in ("^m", "^b"):
            search = container.Duplicate
            finder = search.Find
            finder.ClearFormatting()
            if finder.Execute(FindText=code, MatchWildcards=False, Forward=True, Wrap=0):
                search.Delete()
                found = True
                break
        if not found:
            break
    return container.End


def read_document_text(doc, normalize: bool = True, bio_lines: set | None = None) -> str:
    """Текст документа, зібраний З АБЗАЦІВ, а не з `Content.Text`.

    `normalize=False` — без `normalize_item_numbering`: перевірці наказу треба
    бачити «1.Капітана» таким, як його набрали.

    `bio_lines` (множина, заповнюється тут) — номери рядків тексту, чий абзац
    має лівий відступ від `BIO_MIN_LEFT_INDENT_PT`: біографічний блок пункту.
    Маршрутизація не обриває на них пункт (`map_military_units(bio_lines=)`).

    `Content.Text` склеює цілий рядок таблиці в один рядок тексту, тоді як
    `doc.Paragraphs` рахує кожну комірку окремим абзацом. Через це нумерація
    рядків розходилася з нумерацією абзаців, і після будь-якої таблиці в тілі
    наказу пункти зіставлялися не з тими абзацами — частина пунктів губилася.

    Збираючи текст саме з абзаців, ми отримуємо відповідність «рядок ↔ абзац»
    за побудовою: обидві сторони розбиваються однаково.

    Номер пункту з автонумерації Word у `Range.Text` НЕ входить: такий пункт
    виглядав звичайним текстом, його поглинала шапка, а частини з нього
    ставали контекстом для всіх наступних пунктів. Тому номер номерного списку
    дописуємо на початок абзацу — лише для номерів («1.», «2)»), не для
    маркерів. Кількість рядків не змінюється.
    """
    list_numbers = _list_item_numbers_by_start(doc)
    lines = []
    line_index = 0
    for paragraph in iter_paragraphs(doc):
        paragraph_range = paragraph.Range
        paragraph_text = (paragraph_range.Text or "").rstrip("\r\x07")
        if list_numbers:
            try:
                number = list_numbers.get(int(paragraph_range.Start))
            except Exception:
                number = None
            if number and not re.match(r"^\s*\d", paragraph_text):
                paragraph_text = f"{number} {paragraph_text}"
        lines.append(paragraph_text)
        line_count = paragraph_line_count(paragraph_text)
        if bio_lines is not None and paragraph_text.strip():
            try:
                indent = float(paragraph.LeftIndent)
            except Exception:
                indent = 0.0
            # 9999999 — «невизначено» Word, не відступ.
            if BIO_MIN_LEFT_INDENT_PT <= indent < 9999999:
                bio_lines.update(range(line_index, line_index + line_count))
        line_index += line_count
    text = "\n".join(lines)
    return normalize_item_numbering(text) if normalize else text


_LIST_ITEM_NUMBER_RE = re.compile(r"^\d{1,3}(?:\.\d{1,3})*[\.\)]$")


def _list_item_numbers_by_start(doc) -> dict[int, str]:
    """Номери автонумерації Word за початком абзацу: `{Range.Start: "1."}`.

    Один прохід `ListParagraphs` — лише абзаци списків, а не весь документ.
    Документ без списків (чи фейк у тестах) дає порожній словник.
    """
    numbers: dict[int, str] = {}
    try:
        list_paragraphs = doc.ListParagraphs
        if not int(list_paragraphs.Count):
            return numbers
        for paragraph in list_paragraphs:
            paragraph_range = paragraph.Range
            list_string = str(paragraph_range.ListFormat.ListString or "").strip()
            if _LIST_ITEM_NUMBER_RE.match(list_string):
                numbers[int(paragraph_range.Start)] = list_string
    except Exception:
        return numbers
    return numbers


#: Пункт про смерть / загибель: лише йому дозволено шрифт менший за 14, якщо
#: він не вміщується на сторінку (користувач, 29.09.2026). Розміри — по черзі.
DEATH_ITEM_FONT_SIZES = (13, 12)


# Додатковий прохід стиснення витягу, який після підбору інтервалу 16 → 14 пт
# усе ще займає більше сторінок, ніж міг би (користувач, 26.10.2026: «є моменти,
# коли можна втиснути все на 1 сторінку»). Спершу зменшуються лише ПОРОЖНІ рядки
# між пунктами, потім — і текст, але не нижче 12 пт (дозволив користувач
# 26.10.2026; менший точний інтервал при шрифті 14 вже помітно стискає рядки).
# Прохід приймається лише тоді, коли сторінок справді стає менше; інакше витяг
# лишається як був.
EXTRA_BLANK_SPACINGS = (13.0, 12.0, 11.0, 10.0, 9.0, 8.0, 7.0, 6.0)


EXTRA_TEXT_SPACINGS = (13.75, 13.5, 13.25, 13.0, 12.75, 12.5, 12.25, 12.0)


_DEATH_ITEM_RE = re.compile(r"загинув|загинула|загибел|помер(?:ла)?\b|смерт", re.IGNORECASE)


def is_death_item_text(text: str) -> bool:
    """Чи йдеться в пункті про смерть або загибель військовослужбовця."""
    return bool(_DEATH_ITEM_RE.search(text or ""))


def is_biographical_paragraph(p_text: str) -> bool:
    """Визначає, чи є абзац біографічним блоком (дата народження, освіта, служба, РНОКПП/ІПН)."""
    t = (p_text or "").strip().casefold()
    if not t:
        return False
    if t.startswith("відповідно до") or "звільнити" in t or "призначити:" in t or "наказую:" in t:
        return False
    if re.match(r"^\d{1,3}[\.\)]\s+\D", t):
        return False
    if re.fullmatch(r"\d{5,12}\.?", t):
        return True
    # «р. н.» пишуть і злитно, і з пробілами — в офіційному зразку саме
    # з пробілом. Через вузьку перевірку рядок р.н. не вважався
    # біографічним, і обов'язковий порожній абзац з'їжджав на ІПН.
    # У пунктах про смерть / загибель біографія починається з «Народився …».
    if t.startswith(("народився", "народилася", "народилась")):
        return True
    if re.search(r"\bр\s*\.\s*н\s*\.", t) or "року народження" in t:
        return True
    if "освіта:" in t or "освіта -" in t or "освіта –" in t or "закінчив у" in t:
        return True
    if "у зс із" in t or "у зс з" in t or "у зсу із" in t or "у зсу з" in t or re.search(r"у\s+зс\s+(?:із|з)\s+\d{2}", t):
        return True
    if "рнокпп" in t or "іпн" in t or "ідентифікаційний номер" in t:
        return True
    if re.fullmatch(r"вос\s*-\s*\d+[\w\s\.]*", t):
        return True
    return False


def detect_word_extension(path: str) -> str:
    """Визначає справжній формат файлу Word за сигнатурою, а не за розширенням.

    Word відмовляється відкривати файл, якщо його вміст не відповідає
    розширенню («формат і розширення файлу не збігаються»). Це стається,
    коли шаблон збережено у форматі Word 97-2003, а робоча копія отримує
    розширення `.docx`. Тому робочу копію треба створювати з тим
    розширенням, яке відповідає РЕАЛЬНОМУ вмісту файлу.
    """
    try:
        with open(path, "rb") as handle:
            signature = handle.read(8)
    except OSError:
        return os.path.splitext(path)[1] or ".docx"

    if signature.startswith(b"PK\x03\x04"):
        return ".docx"  # OOXML (zip-контейнер)
    if signature.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"):
        return ".doc"   # OLE2, Word 97-2003
    if signature.startswith(b"{\\rtf"):
        return ".rtf"
    return os.path.splitext(path)[1] or ".docx"


def is_path_writable(path: str) -> bool:
    """Чи можна створити або перезаписати файл за цим шляхом.

    Повертає False, якщо файл тримає інший процес (найчастіше — відкритий
    у Word). Це дозволяє показати зрозуміле повідомлення замість службової
    помилки COM «Не вдається зберегти файл, якщо він використовується
    іншим процесом».
    """
    if not os.path.exists(path):
        return True
    try:
        with open(path, "ab"):
            return True
    except OSError:
        return False


def copy_template_for_editing(
    template_path: str, output_path: str, label: str = "зразок"
) -> str:
    """Створює робочу копію шаблону з коректним для його вмісту розширенням.

    Повертає шлях робочої копії. Якщо він відрізняється від `output_path`,
    викликач має після `SaveAs2` прибрати проміжний файл.
    """
    if not os.path.isfile(template_path):
        # shutil.copy2 у Windows кидає «[WinError 2] The system cannot find the
        # file specified» БЕЗ назви файлу, а detect_word_extension мовчки
        # проковтує відсутній файл — через це причина збою була невидима.
        raise UserError(
            f"немає файлу, який вибраний як {label}:\n{template_path}",
            "Такого файлу на диску немає: його перейменували, перемістили або "
            "видалили. Відкрийте «Зразки та реквізити», натисніть «Вибрати» "
            f"біля рядка «{label}» і вкажіть файл заново.",
        )
    real_ext = detect_word_extension(template_path)
    if real_ext.lower() == os.path.splitext(output_path)[1].lower():
        working_path = output_path
    else:
        working_path = os.path.splitext(output_path)[0] + real_ext
    shutil.copy2(template_path, working_path)
    return working_path


def force_quit_word(word, timeout: float = 5.0) -> None:
    """Завершує Word.Application, а якщо конкретний процес (за PID) завис і
    не закрився самостійно протягом timeout секунд — примусово завершує
    ЛИШЕ цей процес, не чіпаючи інші відкриті вікна Word користувача."""
    pid = None
    try:
        import win32process
        _, pid = win32process.GetWindowThreadProcessId(word.Hwnd)
    except Exception:
        pid = None

    try:
        word.Quit()
    except Exception:
        pass

    if not pid:
        return

    try:
        import time
        import win32api
        import win32con
        import win32event
    except Exception:
        return

    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            handle = win32api.OpenProcess(win32con.SYNCHRONIZE, False, pid)
        except Exception:
            return  # процес уже завершився
        try:
            if win32event.WaitForSingleObject(handle, 200) == win32con.WAIT_OBJECT_0:
                return  # процес самостійно завершився
        finally:
            win32api.CloseHandle(handle)

    # Процес не завершився сам за timeout секунд — примусово завершуємо
    # виключно цей PID (саме той Word, який використовувався для генерації).
    try:
        handle = win32api.OpenProcess(win32con.PROCESS_TERMINATE, False, pid)
        win32api.TerminateProcess(handle, 0)
        win32api.CloseHandle(handle)
    except Exception:
        pass


#: Word повертає це значення замість властивості, яка в діапазоні НЕОДНАКОВА
#: (наприклад, вирівнювання, якщо абзаци діапазону вирівняні по-різному).
WD_UNDEFINED = 9999999


def iter_paragraphs(container) -> list:
    """Абзаци документа чи діапазону — через ІТЕРАТОР колекції, а не Paragraphs(i).

    Word не тримає абзаци масивом: кожне звернення `Paragraphs(i)` проходить
    історію з початку, тож звичний цикл `for i in range(1, Count + 1)` виходить
    квадратичним. Для документа на 60 абзаців це непомітно, але такі цикли тут
    виконуються десятки разів на кожен витяг. Ітератор колекції віддає ті самі
    об'єкти одним проходом.

    Фолбек на індексацію лишено для фейкових COM-об'єктів у тестах, які
    ітератора не мають.
    """
    paragraphs = getattr(container, "Paragraphs", container)
    try:
        return list(paragraphs)
    except TypeError:
        return [paragraphs(index) for index in range(1, paragraphs.Count + 1)]


def last_paragraph(document):
    """Останній абзац документа: `Paragraphs.Last` замість `Paragraphs(Count)`.

    Індексація останнім номером щоразу проходить увесь документ, а `Last`
    Word віддає одразу. Прибирання «висячих» порожніх абзаців у кінці —
    цикл саме з таких звернень.
    """
    paragraphs = document.Paragraphs
    direct = getattr(paragraphs, "Last", None)
    if direct is not None:
        return direct
    return paragraphs(paragraphs.Count)


#: Налаштування Word, які на час пакета вимикаються: (об'єкт, властивість, значення).
#: Фонова репагінація перебудовує макет після КОЖНОЇ правки абзацу, хоча генератор
#: і так кличе `Repaginate()` там, де вимір справді потрібен; перевірка орфографії
#: працює на файлах, які живуть кілька секунд.
_BATCH_WORD_SETTINGS = (
    ("application", "ScreenUpdating", False),
    ("options", "Pagination", False),
    ("options", "CheckSpellingAsYouType", False),
    ("options", "CheckGrammarAsYouType", False),
)


def _word_setting_target(word, scope):
    return word if scope == "application" else getattr(word, "Options", None)


def tune_word_for_batch(word) -> list:
    """Вимикає фонову роботу Word на час пакета; повертає знімок для відкату.

    Частина цих налаштувань у Word спільна для всієї програми й зберігається між
    запусками, тому початкові значення обов'язково запам'ятовуємо: користувач не
    має після генерації знайти в СВОЄМУ Word вимкнену перевірку орфографії.
    Кожна властивість окремо — набір `Options` різниться між версіями, і одна
    відсутня назва не має валити пакет.
    """
    snapshot = []
    for scope, prop, value in _BATCH_WORD_SETTINGS:
        target = _word_setting_target(word, scope)
        if target is None:
            continue
        try:
            snapshot.append((scope, prop, getattr(target, prop)))
        except Exception:
            continue
        try:
            setattr(target, prop, value)
        except Exception:
            pass
    return snapshot


def restore_word_settings(word, snapshot) -> None:
    """Повертає налаштування Word, змінені `tune_word_for_batch`."""
    for scope, prop, value in snapshot or ():
        target = _word_setting_target(word, scope)
        if target is None:
            continue
        try:
            setattr(target, prop, value)
        except Exception:
            pass


def _carry_source_formatting(doc, start: int, end: int, source_paragraph) -> None:
    """Переносить шрифт і геометрію абзацу наказу ЯВНО (правило 5.2.1).

    `FormattedText` не переносить властивість, яка в наказі дорівнює типовій:
    якщо в наказі шрифт заданий стилем `Normal` (Times New Roman 14), окремої
    ознаки шрифту в абзаці немає, і в документі-результаті такий абзац
    успадковує `Normal` ШАБЛОНА — звідти в повідомленні й брався чужий
    шрифт замість Times New Roman. `Font` та `ParagraphFormat` діапазону
    віддають ДІЮЧІ значення наказу, тому переносимо саме їх.

    Мішаний абзац Word віддає порожньою назвою шрифту (`""`) і розміром
    `9999999` — такі значення пропускаємо: у ньому шрифти вже задані явно
    й перенеслися разом із `FormattedText`.
    """
    font_values, geometry = _read_source_formatting(source_paragraph.Range)
    _apply_carried_formatting(doc, start, end, font_values, geometry)


def _read_source_formatting(source_range) -> tuple[dict, dict]:
    """Діючі шрифт і геометрія абзацу наказу (див. `_carry_source_formatting`)."""
    source_font = source_range.Font
    source_format = source_range.ParagraphFormat

    font_values = {}
    try:
        font_name = str(source_font.Name or "").strip()
        if font_name:
            font_values["Name"] = font_name
    except Exception:
        pass
    try:
        font_size = float(source_font.Size)
        if 0 < font_size < 1000:  # 9999999 = мішаний розмір
            font_values["Size"] = font_size
    except Exception:
        pass

    geometry = {}
    for prop in ("Alignment", "LeftIndent", "RightIndent", "FirstLineIndent"):
        try:
            geometry[prop] = getattr(source_format, prop)
        except Exception:
            pass
    return font_values, geometry


def _apply_carried_formatting(doc, start: int, end: int, font_values: dict, geometry: dict) -> None:
    """Записує прочитане `_read_source_formatting` у діапазон результату."""
    destination = doc.Range(start, end)
    for prop, value in font_values.items():
        try:
            setattr(destination.Font, prop, value)
        except Exception:
            pass
    # `ParagraphFormat` діапазону діє на ВСІ його абзаци — одне звернення до
    # Word замість `Paragraphs(i)` на кожен абзац (на слабкому ПК це й було
    # головним гальмом повідомлень).
    dest_format = destination.ParagraphFormat
    try:
        dest_format.PageBreakBefore = False
    except Exception:
        pass
    for prop, value in geometry.items():
        try:
            setattr(dest_format, prop, value)
        except Exception:
            pass
