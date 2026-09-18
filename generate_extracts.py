# Інтерфейс будує Qt-оболонка (`generate_extracts_qt.py`, `generator_qt`). Логіка
# написана під API tkinter (StringVar, messagebox, filedialog, константи), і оболонка
# перед запуском підміняє ці імена своєю реалізацією (`install_qt_bridge`).
from __future__ import annotations

import os
import sys
import json
import re
import shutil
import time
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.constants import *

import win32com.client
import openpyxl

# Додаємо src до sys.path для імпорту модулів nodeautomationtoolkit
project_root = os.path.dirname(os.path.abspath(__file__))
src_path = os.path.join(project_root, "src")
if src_path not in sys.path:
    sys.path.insert(0, src_path)

from nodeautomationtoolkit.builtin_nodes.recipient_mapping import (
    read_recipient_mapping,
    _build_unit_fuzzy_pattern,
    keep_open_names_of,
    _table_cipher,
    map_military_units,
    normalize_item_numbering,
    _format_full_closed_unit_text,
    _format_item_numbers_range,
)
from nodeautomationtoolkit.builtin_nodes.message_order import (
    generate_decision_order,
    cipher_unit_names,
    find_content_start_line,
    reflow_soft_breaks,
)
from nodeautomationtoolkit.builtin_nodes.copy_generator import (
    PREVIEW_DEFAULT_DELAY,
    PreviewSteps,
    build_copy_document,
    retry_on_busy_word,
    _ORDER_BODY_KEYWORDS,
)
from nodeautomationtoolkit.builtin_nodes.blank_images import remove_blank_images
from nodeautomationtoolkit.builtin_nodes.template_tags import (
    SIGNER_TAGS, tag_aliases, expand_common_tags, certifier_tags, signer_tags,
)


def _save_table_to_excel(filepath: str, headers: list[str], rows: list):
    """Зберігає таблицю в Excel (.xlsx) за допомогою openpyxl."""
    os.makedirs(os.path.dirname(os.path.abspath(filepath)), exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(headers)
    for row in rows:
        if isinstance(row, dict):
            ws.append([row.get(h, "") for h in headers])
        else:
            ws.append(list(row))
    try:
        wb.save(filepath)
    except PermissionError:
        raise PermissionError(
            f"Файл '{os.path.basename(filepath)}' зайнятий іншою програмою. "
            f"Будь ласка, закрийте його в Excel і спробуйте знову."
        )


def format_ukr_date(d_str: str) -> str:
    """Форматує дату у військовий стандарт: “15” серпня 2026 року."""
    try:
        d = datetime.strptime(d_str.strip(), "%d.%m.%Y")
    except (AttributeError, ValueError):
        return ""
    months = [
        "січня", "лютого", "березня", "квітня", "травня", "червня",
        "липня", "серпня", "вересня", "жовтня", "листопада", "грудня"
    ]
    m_str = months[d.month - 1]
    return f"“{d.strftime('%d')}” {m_str} {d.year} року"


def format_message_date(d_str: str) -> str:
    """Форматує дату для ПОВІДОМЛЕНЬ: 20.05.2025 року.

    У повідомленнях дата не розкривається словами: на відміну від витягів,
    де діє військовий стандарт `“20” травня 2025 року` (`format_ukr_date`),
    тут лишається числовий запис.
    """
    try:
        d = datetime.strptime(str(d_str).strip(), "%d.%m.%Y")
    except (AttributeError, ValueError):
        return ""
    return f"{d.strftime('%d.%m.%Y')} року"


def extract_metadata_from_filename(filename: str):
    """Витягує номер і дату наказу виключно з назви файлу."""
    order_num = ""
    order_date = ""
    m_num = re.search(r"№\s*([A-Za-zА-Яа-яІіЇїЄєҐґ0-9\-\/]+)", filename)
    if m_num:
        order_num = m_num.group(1).strip()
    m_date = re.search(r"від\s+([0-9]{2}\.[0-9]{2}\.[0-9]{4})", filename.lower())
    if m_date:
        order_date = m_date.group(1).strip()
    return order_num, order_date


def extract_metadata_from_text(text: str) -> tuple[str, str]:
    """Зчитує номер і дату з шапки наказу, якщо їх немає у назві файла."""
    header_text = (text or "")[:5000]
    order_num = ""
    order_date = ""

    m_num = re.search(r"№\s*([A-Za-zА-Яа-яІіЇїЄєҐґ0-9\-\/]+)", header_text)
    if m_num:
        order_num = m_num.group(1).strip()

    m_date = re.search(r"\b([0-9]{2}\.[0-9]{2}\.[0-9]{4})\b", header_text)
    if m_date:
        order_date = m_date.group(1)
    else:
        months = {
            "січня": "01", "лютого": "02", "березня": "03", "квітня": "04",
            "травня": "05", "червня": "06", "липня": "07", "серпня": "08",
            "вересня": "09", "жовтня": "10", "листопада": "11", "грудня": "12",
        }
        m_word_date = re.search(
            r"[«“‘\"]?\s*(\d{1,2})\s*[»”’\"]?\s+(" + "|".join(months) + r")\s+(\d{4})(?:\s+року|\s+р\.)?",
            header_text,
            re.IGNORECASE,
        )
        if m_word_date:
            day, month_name, year = m_word_date.groups()
            order_date = f"{int(day):02d}.{months[month_name.lower()]}.{year}"

    return order_num, order_date


def sanitize_filename(name: str, replacement: str = "_") -> str:
    """Очищає назву файлу від небезпечних та заборонених символів ОС Windows."""
    clean = re.sub(r'[\\/:*?"<>|\r\n\t\x00-\x1f]', replacement, str(name or ""))
    return clean.strip(". ")


def build_extracts_filename(order_num: str, order_date: str, prefix: str = "Витяги наказу") -> str:
    """Назва зведеного файла витягів без вигаданих реквізитів наказу."""
    parts = [prefix]
    if order_num:
        safe_number = sanitize_filename(order_num)
        parts.append(f"№ {safe_number}")
    if order_date:
        safe_date = sanitize_filename(order_date)
        parts.append(f"від {safe_date}")
    return " ".join(parts) + ".docx"


def select_extracts_for_scope(map_result: dict, scope: str = "all") -> tuple[dict, dict]:
    """Повертає адресні й управлінські витяги для вибраного режиму запуску.

    Маршрутизатор завжди аналізує наказ повністю. Розділення робимо вже після
    аналізу, щоб окрема кнопка управлінських витягів не створювала другий набір
    правил розпізнавання й не могла розійтися зі звичайною генерацією.
    """
    units = map_result.get("unit_paragraphs", {})
    management = map_result.get("management_paragraphs", {})
    if scope == "general":
        return units, {}
    if scope == "management":
        return {}, management
    if scope == "all":
        return units, management
    raise ValueError(f"Невідомий режим генерації витягів: {scope}")


MANAGEMENT_RESULT_LABEL = "Управління · окремий витяг"


def management_calculation_rows(map_result: dict) -> list[tuple[str, str, int]]:
    """Інформаційні рядки управління для таблиці результатів у програмі.

    Вони навмисно не додаються до XLSX розрахунку розсилки: це не адресати й
    не відправлення. Зелений рядок у UI лише показує, що пункт розпізнано та
    для нього можна сформувати окремий витяг до управління.
    """
    rows = []
    for data in map_result.get("management_paragraphs", {}).values():
        items = data.get("items", []) if isinstance(data, dict) else []
        labels = [item.get("label", "") for item in items if isinstance(item, dict)]
        rows.append((MANAGEMENT_RESULT_LABEL, _format_item_numbers_range(labels), len(items)))
    return rows


def _saved_xlsx_rows(path: str) -> list[tuple]:
    """Читає вже створений контрольний XLSX без будь-якого перерахунку."""
    if not os.path.isfile(path):
        return []
    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = workbook.active
        return [
            tuple("" if value is None else value for value in row)
            for row in sheet.iter_rows(min_row=2, values_only=True)
            if any(value not in (None, "") for value in row)
        ]
    finally:
        workbook.close()


def load_saved_analysis_reports(order_path: str, output_folder: str) -> dict:
    """Завантажує збережені розсилку, пропуски й маршрутизацію для наказу."""
    order_base = sanitize_filename(os.path.splitext(os.path.basename(order_path))[0])
    paths = {
        "calculation": os.path.join(output_folder, f"Розрахунок_розсилки_{order_base}.xlsx"),
        "unmatched": os.path.join(output_folder, f"Контроль_пропущених_пунктів_{order_base}.xlsx"),
        "routing": os.path.join(output_folder, f"Контроль_маршрутизації_{order_base}.xlsx"),
    }
    errors = []

    def read(key: str) -> list[tuple]:
        try:
            return _saved_xlsx_rows(paths[key])
        except Exception as error:
            errors.append(f"{os.path.basename(paths[key])}: {error}")
            return []

    calculation_source = read("calculation")
    unmatched_source = read("unmatched")
    routing_source = read("routing")
    calculation_rows = [tuple(row[:3]) for row in calculation_source]
    routing_rows = [
        (
            row[0] if len(row) > 0 else "",
            row[1] if len(row) > 1 else "",
            row[2] if len(row) > 2 else "",
            row[5] if len(row) > 5 else (row[-1] if row else ""),
        )
        for row in routing_source
    ]
    # Старі й нові контрольні файли містять правило «зміна до управління».
    # Саме з нього відновлюємо зелені інформаційні рядки, не змінюючи
    # еталонний XLSX розсилки й не запускаючи маршрутизатор повторно.
    management_rows = [
        (MANAGEMENT_RESULT_LABEL, str(row[0]), 1)
        for row in routing_source
        if len(row) > 2 and "управлін" in str(row[2]).casefold()
    ]
    calculation_rows.extend(management_rows)
    return {
        "calculation": calculation_rows,
        "unmatched": [tuple(row[:3]) for row in unmatched_source],
        "routing": routing_rows,
        "found": {key: os.path.isfile(path) for key, path in paths.items()},
        "paths": paths,
        "errors": errors,
    }


def read_document_text(doc) -> str:
    """Текст документа, зібраний З АБЗАЦІВ, а не з `Content.Text`.

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
    return normalize_item_numbering("\n".join(lines))


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


def _slash_to_lines(text_val: str) -> str:
    """Конвертує слеші ' / ' або '/' (крім 'в/ч' та дат) у переноси рядків для Word.

    Використовується і витягами, і примірниками, тому живе на рівні модуля:
    як вкладена функція вона була видима лише всередині `run_extracts`.
    """
    if not text_val:
        return ""
    t = re.sub(r"\b([вВ])\s*/\s*([чЧ])\b", r"\1_SLASH_TEMP_\2", str(text_val))
    t = re.sub(r"(\d)\s*/\s*(\d)", r"\1_NUMSLASH_TEMP_\2", t)
    t = re.sub(r"\s*/\s*", "\r", t)
    t = t.replace("_SLASH_TEMP_", "/")
    t = t.replace("_NUMSLASH_TEMP_", "/")
    return t.replace("\n", "\r")


def back_page_tag_values(order_num: str, order_date: str) -> dict[str, str]:
    """Значення офіційних тегів односторінкової «задньої сторінки».

    Номер і дата беруться лише з назви файла наказу. Порожні значення не
    повертаються: відповідний тег має лишитися у шаблоні для ручного внесення.
    """
    values = {
        "{{згідно_з_оригіналом}}": "Згідно з оригіналом",
        "{{примірник}}": "Примірник № 2",
    }
    if order_num:
        # Знак «№» є частиною значення — так само, як у витягах, інакше
        # у примірнику лишалося б саме лише число.
        values["{{номер_наказу}}"] = f"№{order_num}"
    if order_date:
        values["{{дата_наказу}}"] = format_ukr_date(order_date) or order_date
    return values


# Префікс назви примірника: «2,3» — номери примірників, які друкуються з
# одного файлу. Старий префікс лишається відомим програмі, бо вже згенеровані
# файли нікуди не діваються.
COPY_FILENAME_PREFIX = "2,3"
_LEGACY_COPY_PREFIX = "прим_"


def build_copy_two_filename(order_num: str, order_date: str, source_filename: str) -> str:
    """Назва примірника — така, щоб її прочитав МОДУЛЬ ВИТЯГІВ.

    За правилом 3.3 номер і дата наказу беруться з назви файлу, і шукаються
    там саме у вигляді «№ …» та «від …» (`extract_metadata_from_filename`).
    Стара назва `прим_2_17.08.2026_413.docx` не мала жодного з цих маркерів,
    тому переданий у витяги примірник приходив БЕЗ номера й дати — поля
    доводилось заповнювати руками. Тепер назва замикає це коло сама.

    Реквізити не вигадуються: якщо їх не було в назві наказу, лишається його
    власна назва з префіксом — те, що з неї читалось, читатиметься й далі.
    """
    if order_num and order_date:
        # Скісну риску («б/н», «123/45» — правило 10.5) Windows у назві файлу
        # не дозволяє взагалі, тож зберегти її неможливо. Замінюємо на ДЕФІС,
        # а не на підкреслення: підкреслення не входить у шаблон пошуку номера,
        # і назва читалась назад обрізаною — «б/н» ставало «б», «123/45» → «123».
        safe_number = sanitize_filename(order_num, replacement="-")
        safe_date = sanitize_filename(order_date)
        return f"{COPY_FILENAME_PREFIX}_№{safe_number} від {safe_date}.docx"
    return f"{COPY_FILENAME_PREFIX}_{os.path.basename(source_filename)}"


def is_generated_copy_filename(filename: str) -> bool:
    """Чи це вже згенерований примірник (брати його як наказ не можна).

    Знає й старий префікс `прим_`: файли, зроблені до перейменування, лежать
    у теках користувача й далі, і повторно обробляти їх так само не можна.
    """
    name = os.path.basename(filename or "").lower()
    return name.startswith(COPY_FILENAME_PREFIX.lower() + "_") or _LEGACY_COPY_PREFIX in name


#: Правила набору (нерозривні пробіли, повтори «військової частини», порожні
#: рядки) спільні з генератором наказів — лежать у пакеті, а не тут.
from nodeautomationtoolkit.builtin_nodes.typography import (  # noqa: E402
    ORDER_SIGNER_START_RE as _ORDER_SIGNER_START_RE,
    apply_ukrainian_typography,
    clean_duplicated_units,
    ensure_blank_line_before_items,
)


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


# Вид частини з номером: «169 батальйону резерву», «12 навчального центру»,
# «300 військового госпіталю». Без цього жовтим позначались лише бригади,
# полки й «окремі …», а решта нових частин лишалась у закритому повідомленні
# відкритою БЕЗ жодної позначки (розд. 9.5.7).
_NUMBERED_UNIT_KIND = (
    r"(?:бригад(?:а|и|і|у|ою)|полк(?:у|ом|і|ові)?|батальйон(?:у|і|ом|ові)?"
    r"|дивізіон(?:у|і|ом|ові)?|загін|загон(?:у|і|ом)|центр(?:у|і|ом)?"
    r"|госпітал(?:ь|ю|і|ем)|баз(?:а|и|і|у|ою)|вуз(?:ол|ла|лу|лі|лом)"
    r"|арсенал(?:у|і|ом)?|корпус(?:у|і|ом)?|дивізі(?:я|ї|ю|єю)|комендатур(?:а|и|і|у|ою)"
    # Види зі словника користувача: станція та вузол ФПЗ, ремонтна майстерня,
    # ремонтна майстерня. «Частина» до переліку НЕ входить: вона ловила б
    # «до пункту 2 частини четвертої статті» — у картографічної частини своє правило.
    r"|станці(?:я|ї|ю|єю)|майстерн(?:я|і|ю|ею))"
)
# Слово між номером і видом частини. Номер підрозділу («2 відділу», «3 взводу»)
# і номер статті/пункту частиною не є — такі слова проміжок не пропускає.
_NUMBERED_UNIT_GAP_WORD = (
    r"(?!(?:року|рік|років|відділ\w*|взвод\w*|рот[аиіу]|роті|батаре\w*|груп\w*|служб\w*"
    r"|пункт\w*|частин\w*|стат\w*|наказ\w*)\b)"
    r"[а-яіїєґʼ'’-]+"
)
_UNMATCHED_OPEN_UNIT_RE = re.compile(
    r"\b(?:"
    r"(?:\d{1,3}\s*(?:-?[а-яіїєґ]+)?\s*)?(?:окрем\w+\s+)+(?:механізован\w+|танков\w+|десантн\w+|артилерійськ\w+|піхотн\w+|єгерськ\w+|стрілецьк\w+|штурмов\w+|розвідувальн\w+|гірсько-штурмов\w+|десантно-штурмов\w+|аеромобільн\w+|повітряно-десантн\w+|зв['’]язку)?\s*(?:бригад\w*|полк\w*|батальйон\w*|дивізіон\w*|загін\w*|центр\w*)"
    r"|"
    r"\d{1,3}\s*(?:-?[а-яіїєґ]+)?\s*(?:механізован\w+|танков\w+|десантн\w+|артилерійськ\w+|піхотн\w+|стрілецьк\w+|штурмов\w+|десантно-штурмов\w+|аеромобільн\w+)?\s*(?:бригад\w*|полк\w*|армійськ\w+\s+корпус\w*|АК)"
    r"|"
    r"(?:армійськ\w+\s+корпус\w*|\b\d{1,3}\s*АК\b)"
    r"|"
    # «10 командного пункту протиповітряної оборони», «командно-розвідувальний пункт».
    # Саме слово «пункт» у наказі означає пункт наказу, тому поруч потрібне «командн…».
    r"(?<![\d./,:-])\d{1,4}\s+(?:[\w’'ʼ-]+\s+){0,3}?команд\w*[-\s]?\w*\s+пункт\w*"
    r"|"
    # «картографічна частина» — лише з цим прикметником, інакше ловилося б
    # «частини четвертої статті».
    r"(?:картографічн\w+\s+частин\w*)"
    r"|"
    # «окремої танкової Тестівської бригади» — топонім між словами назви (додаток 53)
    r"(?:окрем\w+\s+)(?:[\w’'ʼ-]+\s+){1,3}?(?:бригад\w*|полк\w*|батальйон\w*|дивізіон\w*|центр\w*)"
    r"|"
    r"(?<![\d./,:-])\d{1,4}(?:\s*-?\s*(?:й|го|му|ий|ого|ому|им))?\s+"
    rf"(?:{_NUMBERED_UNIT_GAP_WORD}\s+){{0,4}}?{_NUMBERED_UNIT_KIND}"
    r")\b",
    re.IGNORECASE | re.UNICODE,
)
# Після назви стоїть шифр або посилання «цієї самої …» — частину вже названо.
_UNIT_ALREADY_CLOSED_AFTER_RE = re.compile(
    r"^\s*(?:(?:військов\w+\s+частин\w*|в\s*/?\s*ч)\s+[АA]?\d+"
    r"|(?:ціє|цьо|тіє|то|цій|цим|цією|тим|тією)\w*\s+сам\w*)",
    re.IGNORECASE,
)

def find_unmatched_open_unit_spans(text: str) -> list[tuple[int, int]]:
    """Повертає діапазони відкритих назв частин, які лишилися після шифрування.

    Не підсвічує лінійні внутрішні батальйони/дивізіони, які вже належать закритій в/ч,
    і територіальні центри комплектування (ТЦК лишається відкритим, розд. 9.5.6).
    """
    if not text:
        return []
    spans = []
    for match in _UNMATCHED_OPEN_UNIT_RE.finditer(text):
        start, end = match.start(), match.end()
        following_text = text[end:end + 60]
        if _UNIT_ALREADY_CLOSED_AFTER_RE.match(following_text):
            continue
        if re.match(r"^\s*комплектуванн", following_text, re.IGNORECASE):
            continue
        spans.append((start, end + _quoted_name_length(text, end)))
    return spans


# Родові означення після виду частини належать до назви: «батальйону резерву»,
# «вузла зв'язку», «центру підготовки». Жовта позначка на них закінчується, а
# для заготовки в таблиці їх треба забрати разом із назвою.
_UNIT_NAME_TAIL_WORD_RE = re.compile(
    r"[ \t]+(?!(?:військов\w*|частин\w*|цієї|цього|того|тієї|та|і|й|до|з|із|зі|на|у|в|для|від)\b)"
    r"[а-яіїєґʼ'’-]+(?:у|ю|и|і|ї|ння|ня|ів|ей)(?=[\s,.;:)»”]|$)",
    re.IGNORECASE,  # «КУДИ» пишеться ВЕЛИКИМИ — хвіст там такий самий
)

# Номер і вид частини — щоб відрізнити справді нову частину від ЗГАДКИ вже
# наявної, яка в наказі написана не повністю («55 ОКРЕМОГО ПОЛКУ» при рядку
# «55 окремий полк радіотехнічного забезпечення»). Шифрувати таку згадку не
# можна (пошук лише за номером заборонений, 4.2.7), але й дописувати її в
# таблицю дублем — теж.
_UNIT_KIND_KEYS = (
    ("бригада", r"бригад"), ("полк", r"полк"), ("батальйон", r"батальйон"),
    ("дивізіон", r"дивізіон"), ("загін", r"заг[іо]н"), ("центр", r"центр"),
    ("госпіталь", r"госпітал"), ("база", r"\bбаз[аиіуо]"), ("вузол", r"вуз(?:ол|л)"),
    ("арсенал", r"арсенал"), ("корпус", r"корпус"), ("дивізія", r"дивізі[яїює]"),
    ("комендатура", r"комендатур"),
)


def _unit_number_and_kind(name: str) -> tuple[str, str] | None:
    text = str(name or "")
    number = re.match(r"\s*(\d{1,4})(?!\d)", text)
    if not number:
        return None
    found = []
    for key, pattern in _UNIT_KIND_KEYS:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            found.append((match.start(), key))
    return (number.group(1), min(found)[1]) if found else None


def _similar_table_row(name: str, mapping: dict) -> str:
    """Стовпець A рядка з тим самим номером і видом частини, або порожньо."""
    wanted = _unit_number_and_kind(name)
    if not wanted:
        return ""
    for key, value in (mapping or {}).items():
        column_a = str(value.get("open_name") or key) if isinstance(value, dict) else str(key)
        if _unit_number_and_kind(column_a) == wanted:
            return column_a
    return ""


# Назва в лапках після виду частини («командування «Тест»») — частина назви,
# тож жовта позначка має захоплювати і її.
_QUOTED_NAME_AFTER_RE = re.compile(r"[ \t]*[«“„\"][^«»“”„\"]{0,40}[»”\"]")


def _quoted_name_length(text: str, end: int) -> int:
    match = _QUOTED_NAME_AFTER_RE.match(text, end)
    return match.end() - end if match else 0

def _unit_name_tail_length(text: str, end: int, max_words: int = 3) -> int:
    position = end
    for _ in range(max_words):
        match = _UNIT_NAME_TAIL_WORD_RE.match(text, position)
        if not match:
            break
        position = match.end()
    return position - end


def unmatched_open_unit_spans(text: str, mapping=None) -> list[tuple[int, int]]:
    """`find_unmatched_open_unit_spans` без назв рядків таблиці з позначкою «$».

    Такі назви лишаються відкритими навмисно, тож жовтої позначки не отримують.
    """
    spans = find_unmatched_open_unit_spans(text)
    kept_names = keep_open_names_of(mapping)
    if not spans or not kept_names:
        return spans
    from nodeautomationtoolkit.builtin_nodes.message_order import soften_unit_text

    kept_patterns = [_build_unit_fuzzy_pattern(soften_unit_text(name)) for name in kept_names]
    result = []
    for start, end in spans:
        span_text = soften_unit_text(text[start:end]).casefold()
        window = soften_unit_text(text[max(0, start - 200):end + 200])
        if any(
            span_text in match.group(0).casefold()
            for pattern in kept_patterns
            for match in pattern.finditer(window)
        ):
            continue
        result.append((start, end))
    return result


def collect_new_unit_names(text: str, mapping: dict, similar: list | None = None) -> list[str]:
    """Відкриті назви частин із наказу, яких немає в таблиці (стовпець A).

    Шукається так само, як жовта позначка в повідомленні: текст шифрується, і
    все, що лишилося відкритим, — нові частини. Рядок таблиці з порожнім
    шифром сюди не потрапляє: частина в таблиці вже є, бракує лише шифру.
    """
    ciphered, _, _ = cipher_unit_names(text or "", mapping)
    # Той самий пом'якшувальний фільтр, що й у шифруванні: апострофи, дефіси.
    from nodeautomationtoolkit.builtin_nodes.message_order import soften_unit_text

    known_patterns = [
        _build_unit_fuzzy_pattern(soften_unit_text(str(name)))
        for name in (*(mapping or {}), *keep_open_names_of(mapping))
        if str(name).strip()
    ]
    names: list[str] = []
    seen: set[str] = set()
    for start, end in unmatched_open_unit_spans(ciphered, mapping):
        tail = _unit_name_tail_length(ciphered, end)
        name = re.sub(r"\s+", " ", ciphered[start:end + tail]).strip()
        key = name.casefold()
        if not name or key in seen:
            continue
        seen.add(key)
        softened_name = soften_unit_text(name)
        if any(pattern.search(softened_name) for pattern in known_patterns):
            continue
        # Частина з таким номером і видом у таблиці вже є — це неповна згадка,
        # а не нова частина: у `similar` (для журналу), але не в заготовки.
        close_row = _similar_table_row(name, mapping)
        if close_row:
            if similar is not None:
                similar.append((name, close_row))
            continue
        names.append(name)
    return names


def describe_cipher_problems(problems: list) -> list[str]:
    """Рядки журналу про те, чого бракує в таблиці для закритого змісту."""
    lines = []
    for problem in problems or []:
        if problem[0] == "no_cipher":
            lines.append(
                f"УВАГА: у таблиці порожній шифр (стовпець B) для «{problem[1]}» — "
                "у повідомленні назва лишається відкритою."
            )
        elif problem[0] == "corps_missing":
            lines.append(
                f"УВАГА: для «{problem[1]}» у стовпці D стоїть «{problem[2]}», але рядка "
                "цього корпусу з шифром у таблиці немає — ланку корпусу не додано."
            )
        elif problem[0] == "merged_cell":
            lines.append(
                f"УВАГА: у стовпці A «{problem[1]}» записано дві частини, а шифр у рядку той самий, "
                f"що й у «{problem[2]}» — першу частину в повідомленні не зашифровано. "
                "Запишіть її окремим рядком зі своїм шифром."
            )
    return lines


def _table_open_name_column(sheet) -> int:
    """Номер стовпця відкритої назви (A) — так само, як його визначає читання таблиці."""
    for row in sheet.iter_rows(min_row=1, max_row=10):
        for cell in row:
            key = str(cell.value or "").strip().casefold()
            if "відкрит" in key or ("назва" in key and "закрит" not in key and "скороч" not in key):
                return cell.column
    return 1


def _append_rows_with_excel(source, names: list[str], column: int) -> bool:
    """Дописує рядки-заготовки самим Excel і повертає, чи вдалося.

    openpyxl зберігає формули, але НЕ їхні обчислені значення, а словник
    читається саме за значеннями (розд. 2.2) — після такого запису шифри
    з формул читались би порожніми. Excel під час збереження перераховує
    книгу, тож результати формул лишаються на місці.
    """
    try:
        import win32com.client
    except Exception:
        return False

    excel = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        book = excel.Workbooks.Open(str(Path(source).resolve()))
        try:
            sheet = book.ActiveSheet
            used = sheet.UsedRange
            last_row = used.Row + used.Rows.Count - 1
            for offset, name in enumerate(names, start=1):
                row = last_row + offset
                sheet.Cells(row, column).Value = name
                for shift in (0, 1):
                    sheet.Cells(row, column + shift).Interior.Color = 0x00FFFF  # жовтий (BGR)
            book.Save()
        finally:
            book.Close(SaveChanges=False)
        return True
    except Exception:
        return False
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass


def append_unit_stubs_to_table(table_path: str, names: list[str]) -> dict:
    """Дописує в словник рядки-заготовки для нових частин (розд. 9.5.7).

    Стовпець A — назва так, як її знайдено в наказі (відмінок і повну назву
    виправляє користувач), решта порожня. Такі рядки `read_recipient_mapping`
    пропускає, доки не заповнено шифр, тож маршрутизації вони не зачіпають.

    `.csv` і `.xlsx` без формул дописує openpyxl. Таблицю З ФОРМУЛАМИ дописує
    сам Excel (`_append_rows_with_excel`): openpyxl не зберігає обчислених
    значень, а словник читається саме за ними (розд. 2.2). Якщо Excel
    недоступний, заготовки йдуть в окремий файл поруч. Перед будь-яким записом
    у таблицю робиться резервна копія.

    Повертає `{"added": [...], "path": ..., "backup": ..., "separate": bool}`.
    """
    source = Path(table_path)
    result = {"added": [], "path": str(source), "backup": "", "separate": False}
    wanted: list[str] = []
    for name in names or []:
        clean = re.sub(r"\s+", " ", str(name or "")).strip()
        if clean and clean.casefold() not in {item.casefold() for item in wanted}:
            wanted.append(clean)
    if not wanted or not source.is_file():
        return result

    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")

    def make_backup() -> None:
        backup_path = source.with_name(f"{source.stem}.backup-{stamp}{source.suffix}")
        shutil.copy2(source, backup_path)
        result["backup"] = str(backup_path)

    def busy(path: Path) -> PermissionError:
        return PermissionError(
            f"Файл «{path.name}» відкритий в іншій програмі (найімовірніше в Excel). "
            "Закрийте його, щоб дописати нові частини."
        )

    suffix = source.suffix.casefold()
    if suffix == ".csv":
        import csv

        raw = source.read_text(encoding="utf-8-sig", errors="replace")
        delimiter = ";" if raw.count(";") >= raw.count(",") else ","
        existing = {
            re.sub(r"\s+", " ", row[0]).strip().casefold()
            for row in csv.reader(raw.splitlines(), delimiter=delimiter)
            if row
        }
        added = [name for name in wanted if name.casefold() not in existing]
        if added:
            make_backup()
            try:
                with source.open("a", encoding="utf-8", newline="") as handle:
                    if raw and not raw.endswith(("\n", "\r")):
                        handle.write("\r\n")
                    writer = csv.writer(handle, delimiter=delimiter)
                    for name in added:
                        writer.writerow([name, "", "", "", "", ""])
            except PermissionError:
                raise busy(source)
        result["added"] = added
        return result

    if suffix == ".xlsx":
        from openpyxl.styles import PatternFill

        workbook = openpyxl.load_workbook(source)
        has_formulas = any(
            cell.data_type == "f"
            for sheet in workbook.worksheets
            for row in sheet.iter_rows()
            for cell in row
        )
        sheet = workbook.active
        column = _table_open_name_column(sheet)
        # Наявні назви беремо ЗА ЗНАЧЕННЯМИ: назва може бути результатом формули.
        values_sheet = openpyxl.load_workbook(source, data_only=True).active
        existing = {
            re.sub(r"\s+", " ", str(cell.value)).strip().casefold()
            for (cell,) in values_sheet.iter_rows(min_col=column, max_col=column)
            if cell.value is not None
        }
        added = [name for name in wanted if name.casefold() not in existing]
        if not added:
            result["added"] = []
            return result

        make_backup()
        if has_formulas:
            # Таблицю з формулами дописує сам Excel — інакше зникнуть обчислені
            # значення. Якщо Excel недоступний, нижче спрацює окремий файл.
            if _append_rows_with_excel(source, added, column):
                result["added"] = added
                return result
        else:
            last_row = max(
                (cell.row for row in sheet.iter_rows() for cell in row if cell.value not in (None, "")),
                default=0,
            )
            fill = PatternFill(fill_type="solid", start_color="FFFF00", end_color="FFFF00")
            for offset, name in enumerate(added, start=1):
                sheet.cell(row=last_row + offset, column=column, value=name).fill = fill
                sheet.cell(row=last_row + offset, column=column + 1).fill = fill
            try:
                workbook.save(source)
            except PermissionError:
                raise busy(source)
            result["added"] = added
            return result

    # Excel недоступний або формат не той — окремий файл поруч.
    separate = source.with_name(f"{source.stem} — нові частини.xlsx")
    if separate.is_file():
        stub_book = openpyxl.load_workbook(separate)
        stub_sheet = stub_book.active
    else:
        stub_book = openpyxl.Workbook()
        stub_sheet = stub_book.active
        stub_sheet.append(
            ["Відкрита назва (A)", "Шифр (B)", "Скорочення (C)", "Корпус (D)", "Кому (E)", "Куди (F)"]
        )
    existing = {
        re.sub(r"\s+", " ", str(row[0] or "")).strip().casefold()
        for row in stub_sheet.iter_rows(min_row=2, values_only=True)
        if row
    }
    added = [name for name in wanted if name.casefold() not in existing]
    for name in added:
        stub_sheet.append([name, "", "", "", "", ""])
    if added:
        try:
            stub_book.save(separate)
        except PermissionError:
            raise busy(separate)
    result.update(added=added, path=str(separate), separate=True)
    return result


def build_message_recipient_groups(mapping: dict, routes: dict) -> dict[str, list[str]]:
    """Адресати повідомлення, згруповані за типом: `corps`, `units`, `tck`.

    Корпус і підпорядкована частина завжди є різними адресатами. Корпуси
    розміщуються першими, частини — після них, ТЦК/ОТЦК — наприкінці. Текст
    адресата береться з «Кому» Excel і за потреби доповнюється шифром.
    """
    def is_tck_entry(entry: dict) -> bool:
        entry_names = " ".join(
            str(entry.get(key, "")) for key in ("open_name", "cipher", "abbreviation")
        ).upper()
        return "ТЦК" in entry_names or "КОМПЛЕКТУВАН" in entry_names

    def standalone_cipher(entry: dict) -> str:
        # Не передаємо корпус у цей форматер: він має бути окремим рядком.
        standalone_entry = dict(entry)
        standalone_entry["corps"] = ""
        return _format_full_closed_unit_text(standalone_entry, mapping)

    def recipient_text(entry: dict) -> str:
        recipient_to = str(entry.get("recipient_to") or "").strip()
        # Шифр — лише зі стовпця B (розд. 9.5.7). Без нього в адресат іде тільки
        # «Кому» з таблиці: відкрита назва в закритий супровід не потрапляє.
        cipher = _table_cipher(entry)
        if not cipher:
            return recipient_to
        cipher_text = standalone_cipher(entry)
        if not recipient_to:
            return cipher_text
        if cipher.casefold() in recipient_to.casefold():
            return recipient_to
        # Якщо в рядку «Кому» вже є назва військової частини, додаємо лише
        # шифр, а не повторюємо «військової частини» вдруге.
        if re.search(r"військов\w*\s+частин\w*|в\s*/?\s*ч", recipient_to, re.IGNORECASE):
            return f"{recipient_to} {cipher}".strip()
        return f"{recipient_to} {cipher_text}".strip()

    def find_corps_entry(corps_name: str) -> dict | None:
        direct_entry = mapping.get(corps_name)
        if isinstance(direct_entry, dict):
            return direct_entry
        key = corps_name.casefold()
        for entry in mapping.values():
            if not isinstance(entry, dict):
                continue
            entry_keys = (entry.get("open_name"), entry.get("abbreviation"), entry.get("abbr"))
            if any(str(value or "").strip().casefold() == key for value in entry_keys):
                return entry
        return None

    def find_mapping_entry(identifier: str) -> dict | None:
        key = identifier.strip().casefold()
        for entry in mapping.values():
            if not isinstance(entry, dict):
                continue
            entry_keys = (entry.get("open_name"), entry.get("cipher"), entry.get("abbreviation"), entry.get("abbr"))
            if any(str(value or "").strip().casefold() == key for value in entry_keys):
                return entry
        return None

    def is_oblast_tck_entry(entry: dict | None) -> bool:
        if not entry:
            return False
        identity = " ".join(
            str(entry.get(key, "")) for key in ("open_name", "cipher", "abbreviation")
        )
        return "ОБЛАСН" in identity.upper() or bool(re.search(r"\bОТЦК\b", identity, re.IGNORECASE))

    def is_corps_entry(entry: dict) -> bool:
        values = " ".join(str(entry.get(key, "")) for key in ("open_name", "abbreviation"))
        return "КОРПУС" in values.upper() or bool(re.search(r"\b\d{1,3}\s*АК\b", values, re.IGNORECASE))

    matched_names = set()
    match_report = routes.get("match_report")
    for row in getattr(match_report, "rows", []) or []:
        if row:
            matched_names.add(str(row[0]).strip())

    # Також додаємо всі частини з unit_paragraphs (витяги / маршрутизація)
    for u_key, u_data in routes.get("unit_paragraphs", {}).items():
        if isinstance(u_data, dict):
            open_n = str(u_data.get("open_name") or u_data.get("unit_name") or "").strip()
            if open_n:
                matched_names.add(open_n)
            abbr_n = str(u_data.get("abbreviation") or "").strip()
            if abbr_n:
                matched_names.add(abbr_n)
            code_n = str(u_data.get("unit_code") or "").strip()
            if code_n:
                matched_names.add(code_n)
        if u_key:
            matched_names.add(str(u_key).strip())

    matched_entries = []
    matched_entry_keys = set()
    for open_name, entry in mapping.items():
        if not isinstance(entry, dict):
            continue
        entry_open = str(entry.get("open_name") or open_name).strip()
        entry_cipher = str(entry.get("cipher") or "").strip()
        entry_abbr = str(entry.get("abbreviation") or "").strip()

        is_matched = (
            str(open_name).strip() in matched_names
            or entry_open in matched_names
            or (entry_cipher and entry_cipher in matched_names)
            or (entry_abbr and entry_abbr in matched_names)
        )
        if not is_matched:
            continue
        entry_key = (
            entry_cipher.casefold(),
            entry_open.casefold(),
        )
        if is_tck_entry(entry) or entry_key in matched_entry_keys:
            continue
        matched_entries.append(entry)
        matched_entry_keys.add(entry_key)

    # Додатковий прохід: гарантуємо, що частини з unit_paragraphs додані, навіть якщо не знайшлися напряму за назвою
    for u_key, u_data in routes.get("unit_paragraphs", {}).items():
        if not isinstance(u_data, dict) or is_tck_entry(u_data):
            continue
        u_cipher = str(u_data.get("unit_code") or u_key).strip()
        u_open = str(u_data.get("open_name") or u_data.get("unit_name") or "").strip()
        u_key_tuple = (u_cipher.casefold(), u_open.casefold())
        if u_key_tuple not in matched_entry_keys:
            entry = find_mapping_entry(u_cipher) or find_mapping_entry(u_open) or find_mapping_entry(str(u_key))
            if entry and isinstance(entry, dict):
                e_key = (str(entry.get("cipher") or "").strip().casefold(), str(entry.get("open_name") or "").strip().casefold())
                if e_key not in matched_entry_keys:
                    matched_entries.append(entry)
                    matched_entry_keys.add(e_key)
            elif u_data.get("recipient_to"):
                matched_entries.append(u_data)
                matched_entry_keys.add(u_key_tuple)

    corps_entries = []
    corps_entry_keys = set()
    for entry in matched_entries:
        corps_name = str(entry.get("corps") or "").strip()
        corps_entry = find_corps_entry(corps_name) if corps_name else None
        if corps_entry is not None:
            c_key = (
                str(corps_entry.get("cipher") or "").strip().casefold(),
                str(corps_entry.get("open_name") or "").strip().casefold(),
            )
            if c_key not in corps_entry_keys:
                corps_entries.append(corps_entry)
                corps_entry_keys.add(c_key)
        if is_corps_entry(entry):
            e_key = (
                str(entry.get("cipher") or "").strip().casefold(),
                str(entry.get("open_name") or "").strip().casefold(),
            )
            if e_key not in corps_entry_keys:
                corps_entries.append(entry)
                corps_entry_keys.add(e_key)

    corps_recipients = []
    unit_recipients = []
    for entry in corps_entries:
        recipient = recipient_text(entry)
        if recipient and recipient not in corps_recipients:
            corps_recipients.append(recipient)
    for entry in matched_entries:
        e_key = (
            str(entry.get("cipher") or "").strip().casefold(),
            str(entry.get("open_name") or "").strip().casefold(),
        )
        if e_key in corps_entry_keys:
            continue
        recipient = recipient_text(entry)
        if recipient and recipient not in unit_recipients:
            unit_recipients.append(recipient)

    # Внизу списку лишаються тільки обласні ТЦК, підтверджені рядками Excel.
    tck_recipients = []
    for data in routes.get("unit_paragraphs", {}).values():
        unit_code = str(data.get("unit_code") or "")
        if "ТЦК" not in unit_code.upper():
            continue
        if not is_oblast_tck_entry(find_mapping_entry(unit_code)):
            continue
        recipient = str(data.get("recipient_to") or unit_code).strip()
        if recipient and recipient not in tck_recipients:
            tck_recipients.append(recipient)
    return {"corps": corps_recipients, "units": unit_recipients, "tck": tck_recipients}


def build_message_recipient_list(mapping: dict, routes: dict) -> list[str]:
    """Плоский список для {{кому_список}}: корпуси → частини → ТЦК."""
    groups = build_message_recipient_groups(mapping, routes)
    return groups["corps"] + groups["units"] + groups["tck"]


def build_addressee_kind_text(groups: dict[str, list[str]]) -> str:
    """Текст для тегу {{тцк чі вч}} — кому саме адресоване повідомлення.

    Правило:
    - лише військові частини (враховуючи корпуси) → «командирам військових частин»;
    - лише ТЦК → «начальникам ОТЦК та СП»;
    - і частини, і ТЦК → «командирам військових частин та начальникам ОТЦК та СП».
    """
    has_units = bool(groups.get("corps") or groups.get("units"))
    has_tck = bool(groups.get("tck"))
    if has_units and has_tck:
        return "Командирам військових частин та Начальникам ОТЦК та СП"
    if has_tck:
        return "Начальникам ОТЦК та СП"
    if has_units:
        return "Командирам військових частин"
    return ""


# Написання тегу типу адресата, які розпізнаються в шаблонах повідомлень.
# Пошук у Word не враховує регістр, тому достатньо варіантів написання «чі/чи».
_MESSAGE_ADDRESSEE_KIND_TAGS = ("{{тцк чі вч}}", "{{тцк чи вч}}")


#: Перелік звань окремо: він потрібен і прив'язаним до початку рядка (звичайний
#: підписний блок), і всередині рядка (коли звання стоїть поруч із посадою).
_RANK_ALTERNATIVES = (
    r"генерал(?:[-\s](?:майор|лейтенант|полковник))?|адмірал(?:[-\s]\w+)?|"
    r"полковник|підполковник|майор|капітан(?:[-\s](?:лейтенант|[1-3]\s+рангу))?|"
    r"старший\s+лейтенант|молодший\s+лейтенант|лейтенант|головний\s+сержант|"
    r"штаб[-\s]сержант|майстер[-\s]сержант|старший\s+сержант|молодший\s+сержант|"
    r"сержант|старшина|солдат|матрос"
)
_ORDER_SIGNER_RANK_RE = re.compile(
    r"^\s*(" + _RANK_ALTERNATIVES + r")\b",
    re.IGNORECASE | re.UNICODE,
)
_ORDER_SIGNER_RANK_INLINE_RE = re.compile(
    r"\b(" + _RANK_ALTERNATIVES + r")\b",
    re.IGNORECASE | re.UNICODE,
)
# Пронумерований пункт наказу: «1.», «2.3.», «10)». Та сама форма, що вже
# використовується в recipient_mapping.py.
_ITEM_START_RE = re.compile(r"^\d{1,3}(?:\.\d{1,3})*[\.\)]\s+")


def _last_item_line(lines: list[str]) -> int:
    """Номер рядка ОСТАННЬОГО пронумерованого пункту наказу.

    Підписант завжди йде після пунктів, тому цей рядок — природна нижня межа
    пошуку підписанта. Вона не залежить від довжини службового хвоста, на
    відміну від лічильника рядків, який доводилося збільшувати щоразу, коли
    в наказі траплялася більша таблиця розсилки.

    Подавати сюди треба ЛИШЕ тіло наказу, до маркера розсилки: рядки самої
    таблиці розсилки теж бувають пронумеровані («1. в/ч А0000 — 1 прим.»), і
    тоді межа заїхала б за підписанта, а його пошук не дав би нічого.
    """
    for index in range(len(lines) - 1, -1, -1):
        if _ITEM_START_RE.match(lines[index]):
            return index
    return 0


def plan_2up_page_layout(extract_pages: list[int]) -> list[dict]:
    """
    Розраховує розміщення витягів та вставку порожніх сторінок для друку «2 сторінки на 1 аркуш».
    Кожен фізичний аркуш містить 2 логічні сторінки (ліва - непарна, права - парна).
    Багатосторінкові витяги завжди починаються з нового аркуша і, якщо їхня довжина непарна,
    доповнюються порожньою сторінкою для вирівнювання наступного аркуша.
    """
    if not extract_pages:
        return []

    events = []
    current_doc_pages = 0

    for i, p_count in enumerate(extract_pages):
        # Якщо витяг багатосторінковий і позиція на непарній сторінці -> вставляємо порожню перед
        if p_count > 1 and (current_doc_pages % 2 != 0):
            events.append({"action": "insert_blank_before", "extract_idx": i})
            current_doc_pages += 1

        events.append({
            "action": "insert_extract",
            "extract_idx": i,
            "pages": p_count,
            "start_page": current_doc_pages + 1,
            "end_page": current_doc_pages + p_count,
        })
        current_doc_pages += p_count

        # Якщо витяг багатосторінковий і непарної довжини -> вставляємо порожню після
        if p_count > 1 and (p_count % 2 != 0):
            events.append({"action": "insert_blank_after", "extract_idx": i})
            current_doc_pages += 1

    return events


_DISTRIBUTION_CUTOFF_MARKERS = (
    "розрахунок розсилки",
    "таблиця розсилки",
    "список розсилки",
    "розсилка:",
    "відмітки служби діловодства",
    "служба діловодства",
    "згідно з оригіналом",
    "розіслано:",
    "відмітка про виконання",
)


def find_distribution_cutoff_line(text: str) -> int:
    """Номер рядка, з якого починається службова таблиця розсилки/відміток.

    Усе до цього рядка — власне наказ разом із підписантом. Пошук іде з кінця
    документа, щоб випадкове входження маркера в тілі пункту не обрізало наказ
    передчасно.
    """
    raw_lines = str(text or "").replace("\x07", "").splitlines()
    lines = [re.sub(r"\s+", " ", line).strip() for line in raw_lines]

    # Документ сканується ПОВНІСТЮ: цикл іде з кінця й повертає ОСТАННІЙ
    # маркер, тож обмежувати глибину не потрібно. Раніше тут стояло вікно на
    # кілька сотень рядків, і воно ламалося щоразу, коли службовий хвіст
    # виростав: відколи текст збирається з абзаців (`read_document_text`),
    # кожна комірка таблиці розсилки стала окремим рядком. Маркер лишався поза
    # вікном, підписант визначався неправильно, а останній пункт «затягував»
    # службовий хвіст у витяг.
    for idx in range(len(lines) - 1, -1, -1):
        line_lower = lines[idx].casefold().strip()
        if any(line_lower.startswith(marker) or marker == line_lower for marker in _DISTRIBUTION_CUTOFF_MARKERS):
            return idx
    return len(lines)


def find_service_block_line(text: str, start_line: int = 0) -> int:
    """Номер рядка ПЕРШОГО службового блоку, починаючи з `start_line`.

    На відміну від `find_distribution_cutoff_line`, яка сканує з кінця й
    повертає ОСТАННІЙ маркер, ця функція йде вперед. Для примірника потрібен
    саме перший блок після підписанта: інакше проміжні таблиці на кшталт
    «Розрахунок розсилки витягів із наказу» лишалися всередині документа.
    """
    raw_lines = str(text or "").replace("\x07", "").splitlines()
    lines = [re.sub(r"\s+", " ", line).strip().casefold() for line in raw_lines]
    for index in range(max(0, start_line), len(lines)):
        clean = lines[index]
        if not clean:
            continue
        if any(clean.startswith(marker) or clean == marker for marker in _DISTRIBUTION_CUTOFF_MARKERS):
            return index
    return len(lines)


def _find_order_signer(text: str) -> dict[str, str] | None:
    """Повертає реквізити й номер рядка початку підписанта (Командувача) в наказі, відсікаючи таблицю розсилки."""
    raw_lines = str(text or "").replace("\x07", "").splitlines()
    lines = [re.sub(r"\s+", " ", line).strip() for line in raw_lines]

    reference_index = find_distribution_cutoff_line(text)
    # Пошук іде з кінця й зупиняється на останньому пронумерованому пункті:
    # підписант стоїть після пунктів, а не всередині них. Це водночас знімає
    # залежність від довжини службового хвоста і не дає прийняти за підписанта
    # рядок тіла наказу, що починається з «Начальник…»/«Командир…».
    #
    # Межу пунктів рахуємо до ПЕРШОГО службового блоку, а не до останнього
    # маркера: у зразку звороту останнього аркуша (додаток 43) службових
    # блоків два — «Розрахунок розсилки витягів із наказу:» і «Розрахунок
    # розсилки електронних повідомлень:», — а їхні рядки теж пронумеровані
    # («1. Військова частина А0000 п. 1.»). Рахуючи їх пунктами наказу, межа
    # заїжджала за підписанта, той не знаходився зовсім, і в витяг протікав
    # увесь службовий хвіст.
    body_limit = min(find_service_block_line(text, 0), max(0, reference_index))
    search_start = _last_item_line(lines[:body_limit])

    # Після останнього пункту може бути кілька підписоподібних службових
    # блоків. Для межі тіла наказу потрібен ПЕРШИЙ справжній підписант.
    # Зворотний пошук вибирав останній блок і затягував усе між ними в текст
    # останнього пункту витягу.
    for start_index in range(search_start + 1, reference_index):
        if not _ORDER_SIGNER_START_RE.match(lines[start_index]):
            continue

        position_lines = []
        # Межа блоку — початок службової частини, а не лічильник рядків.
        # Вікно у 8 рядків рвалося, щойно між посадою та званням
        # ставало більше восьми абзаців — а порожніми абзацами підписний
        # блок у наказі часто розсувають до низу сторінки. Підписант
        # тоді не знаходився ЗОВСІМ, тіло наказу не обрізалося, і блок
        # протікав у {{зміст}}.
        for line_index in range(start_index, reference_index):
            line = lines[line_index]
            if not line:
                continue
            rank_match = _ORDER_SIGNER_RANK_RE.match(line)
            if rank_match:
                rank = rank_match.group(1)
                remainder = line[rank_match.end():].strip(" .\t–—-:")
                columns = [
                    part.strip(" .\t–—-:")
                    for part in re.split(r"\t+|\.{3,}", remainder)
                    if part.strip(" .\t–—-:")
                ]
                name = columns[-1] if columns else remainder
                if not name:
                    for next_index in range(line_index + 1, min(reference_index, line_index + 5)):
                        candidate = lines[next_index]
                        if any(marker in candidate.casefold() for marker in _DISTRIBUTION_CUTOFF_MARKERS) or _ORDER_SIGNER_START_RE.match(candidate):
                            break
                        candidate_words = re.findall(r"[A-Za-zА-Яа-яІіЇїЄєҐґ'’`-]+", candidate)
                        if len(candidate_words) >= 2 and candidate.casefold() not in {"підпис", "м. п."}:
                            name = " ".join(candidate_words)
                            break
                return {
                    "start_line": start_index,
                    "position": "\n".join(position_lines),
                    "rank": rank,
                    "name": name,
                }
            if any(marker in line.casefold() for marker in _DISTRIBUTION_CUTOFF_MARKERS):
                break
            if any(char.isalpha() for char in line) and line.casefold() not in {"підпис", "м. п."}:
                position_lines.append(line)

        # Фолбек: звання стоїть не з початку рядка, а в тому ж рядку, що й
        # посада («Командувач військ … генерал-лейтенант Іван ПРІЗВИЩЕНКО»).
        # Без цього блок не впізнавався ЗОВСІМ: тіло наказу не обрізалося, а
        # підписант не переносився у витяг.
        for line_index in range(start_index, reference_index):
            line = lines[line_index]
            if not line:
                continue
            if any(marker in line.casefold() for marker in _DISTRIBUTION_CUTOFF_MARKERS):
                break
            inline_rank = _ORDER_SIGNER_RANK_INLINE_RE.search(line)
            if not inline_rank or inline_rank.start() == 0:
                continue
            tail = line[inline_rank.end():].strip(" .\t–—-:")
            name_words = re.findall(r"[A-Za-zА-Яа-яІіЇїЄєҐґ'’`-]+", tail)
            if len(name_words) < 2:
                continue
            return {
                "start_line": start_index,
                "position": line[: inline_rank.start()].strip(" .\t–—-:"),
                "rank": inline_rank.group(1),
                "name": " ".join(name_words),
            }
    return None


def extract_order_signer(text: str) -> dict[str, str]:
    """Виділяє посаду, звання й ПІБ багаторядкового підписанта наказу."""
    signer = _find_order_signer(text)
    if not signer:
        return {"position": "", "rank": "", "name": ""}
    return {key: signer.get(key, "") for key in ("position", "rank", "name")}


def _first_signer_like_line_after_last_item(text: str) -> int | None:
    """Перша підписоподібна межа після останнього пункту.

    Тимчасове жорстке правило: такий блок ніколи не є продовженням останнього
    пункту, навіть якщо в ньому не вдалося розібрати звання та ПІБ.
    """
    raw_lines = str(text or "").replace("\x07", "").splitlines()
    lines = [re.sub(r"\s+", " ", line).strip() for line in raw_lines]
    body_limit = find_service_block_line(text, 0)
    last_item = _last_item_line(lines[:body_limit])
    for index in range(last_item + 1, body_limit):
        if _ORDER_SIGNER_START_RE.match(lines[index]):
            return index
    return None


def text_before_order_signer(text: str) -> tuple[str, dict[str, str]]:
    """Відсікає підписанта й увесь службовий текст, який іде після нього."""
    signer = _find_order_signer(text)
    signer_like_start = _first_signer_like_line_after_last_item(text)
    if signer_like_start is None and not signer:
        return text, {"position": "", "rank": "", "name": ""}
    raw_lines = str(text or "").splitlines()
    start_line = signer_like_start if signer_like_start is not None else signer["start_line"]
    # Не підтягуємо реквізити з пізнішого підписоподібного блока. Вони
    # прийнятні лише тоді, коли належать саме першій відсіченій межі.
    clean_signer = (
        {key: signer.get(key, "") for key in ("position", "rank", "name")}
        if signer and signer.get("start_line") == start_line
        else {"position": "", "rank": "", "name": ""}
    )
    return "\n".join(raw_lines[:start_line]).rstrip(), clean_signer


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


class UserError(Exception):
    """Помилка, яку показуємо користувачу простими словами.

    `what` — що саме сталося, `todo` — що з цим робити. Текст складається
    так, щоб його зрозуміла людина, яка ніколи не відкривала журнал програми:
    без кодів, без англійських слів, з назвою конкретного файлу.
    """

    def __init__(self, what: str, todo: str = "", detail: str = ""):
        self.what = what.strip()
        self.todo = todo.strip()
        self.detail = detail.strip()
        super().__init__(self.what)

    def __str__(self) -> str:
        parts = [f"Що сталося: {self.what}"]
        if self.todo:
            parts.append(f"Що зробити: {self.todo}")
        if self.detail:
            parts.append(f"(технічні подробиці: {self.detail})")
        return "\n".join(parts)


# Коди Windows, які трапляються в цій програмі найчастіше. Windows часто
# повідомляє їх БЕЗ назви файлу («[WinError 2] The system cannot find the
# file specified»), тому саме тут вони перетворюються на зрозумілий текст.
_WINDOWS_DRIVE_ERRORS = (21, 53, 67, 1231, 1265)  # диск/мережева тека недоступні
_WINDOWS_BUSY_ERRORS = (32, 33)                   # файл тримає інша програма


def _error_file_name(error: BaseException) -> str:
    """Назва файлу з помилки, якщо ОС її повідомила."""
    for attribute in ("filename", "filename2"):
        value = getattr(error, attribute, None)
        if value:
            return str(value)
    return ""


def explain_error(error: BaseException) -> str:
    """Перекладає будь-який збій на просту мову: що сталося і що робити.

    Показується і у вікні, і в журналі. Технічний текст не викидається —
    він лишається в кінці, щоб було з чим прийти по допомогу.
    """
    if isinstance(error, UserError):
        return str(error)

    technical = f"{type(error).__name__}: {error}".strip()
    winerror = getattr(error, "winerror", None)
    name = _error_file_name(error)
    named = f"«{os.path.basename(name)}»" if name else "потрібний файл"
    where = f"\n{name}" if name else ""

    if isinstance(error, FileNotFoundError) or winerror in (2, 3):
        return str(UserError(
            f"комп'ютер не знайшов {named} — такого файлу на диску немає." + where,
            "Файл перейменували, перемістили або видалили. Відкрийте «Зразки та "
            "реквізити» й виберіть його заново кнопкою «Вибрати», а наказ — "
            "кнопкою вибору наказу в головному вікні.",
            technical,
        ))

    if isinstance(error, PermissionError) or winerror in _WINDOWS_BUSY_ERRORS:
        return str(UserError(
            f"файл {named} зараз зайнятий іншою програмою або закритий для запису." + where,
            "Закрийте цей файл у Word чи Excel (і перевірте, чи не відкритий він "
            "у вікні попереднього перегляду) і натисніть кнопку ще раз.",
            technical,
        ))

    if winerror in _WINDOWS_DRIVE_ERRORS:
        return str(UserError(
            "диск або мережева тека, де лежить потрібний файл, зараз недоступні." + where,
            "Перевірте, чи підключений диск (наприклад, E:) і чи є мережа, "
            "потім повторіть.",
            technical,
        ))

    if type(error).__name__ == "com_error":
        return str(UserError(
            "Word не виконав дію: він або зайнятий, або показує власне вікно "
            "(наприклад, запит на збереження), або його закрили під час роботи.",
            "Закрийте всі вікна Word, переконайтесь, що жоден документ не чекає "
            "відповіді, і повторіть. Якщо не допомогло — перезавантажте комп'ютер.",
            technical,
        ))

    return str(UserError(
        "сталася несподівана помилка, і програма зупинила саме цей файл.",
        "Спробуйте ще раз. Якщо повториться — збережіть текст із журналу "
        "(нижнє поле вікна) і покажіть його розробнику.",
        technical,
    ))


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


#: Слова, які у журналі стоять поруч із числом, але частинами НЕ є.
#: Без цього переліку знеособлення зʼїдало б корисну статистику
#: («2 стор.», «13 пунктів», «+39 ентерів») разом із назвами частин.
_LOG_MEASURE_WORDS = (
    "стор", "сторінк", "сторінка", "сторінки", "сторінок", "односторінк", "багатосторінк",
    "арк", "аркуш", "аркуша", "аркушів", "пункт", "пункту", "пунктів", "пунктами",
    "витяг", "витягів", "витягах", "витяги", "абзац", "абзаців", "ентер", "ентерів",
    "шт", "пт", "с", "сек", "секунд", "байт", "записів", "скорочень", "прим",
    "рядк", "рядків", "разів", "файл", "файлів", "з", "із", "та", "і", "на", "від",
)

#: Одне правило знеособлення на всі випадки: перший збіг виграє, і вже
#: підставлена мітка під наступні правила не потрапляє.
_LOG_NAME_WORD = r"[А-ЯІЇЄҐ][а-яіїєґ]+(?:['’\-][А-ЯІЇЄҐа-яіїєґ]+)*"
_LOG_SURNAME_UPPER = r"[А-ЯІЇЄҐ]{2,}(?:['’\-][А-ЯІЇЄҐ]+)*"
_REDACT_RE = re.compile(
    # Шлях закінчується перед « · »: далі в журналі стоять версія модуля
    # маршрутизації, дата зміни словника й розмір. Вони нічого не розкривають,
    # а без них зі знеособленого журналу не видно, чи програму перезапущено.
    r"(?P<path>[A-Za-zА-ЯІЇЄҐ]:[\\/][^\n·]*[^\n·\s])"
    r"|(?P<otck>(?i:[А-ЯІЇЄҐ][а-яіїєґ'’\-]+(?:ськ|цьк|зьк)[а-яіїєґ]*\s+"
    r"(?:обласн\w+\s+|районн\w+\s+|міськ\w+\s+)?(?:О?М?Р?ТЦК|територіальн\w+)"
    r"(?:[^\n,;.]{0,80}?підтримки)?(?:\s+та\s+СП)?))"
    rf"|(?P<name>\b(?:{_LOG_SURNAME_UPPER}[ \t]+{_LOG_NAME_WORD}(?:[ \t]+{_LOG_NAME_WORD})?"
    rf"|{_LOG_NAME_WORD}(?:[ \t]+{_LOG_NAME_WORD})?[ \t]+{_LOG_SURNAME_UPPER}"
    rf"|{_LOG_NAME_WORD}[ \t]+{_LOG_NAME_WORD}(?:[ \t]+{_LOG_NAME_WORD})?)\b)"
    r"|(?P<cipher>\b[АA]\s?\d{3,4}\b)"
    r"|(?P<unit>\b\d{1,3}\s?(?P<unit_word>[А-Яа-яІЇЄҐіїєґ]{1,12})\b)"
    r"|(?P<upper>\b[А-ЯІЇЄҐ]{3,}\b)",
    re.UNICODE,
)

#: ВЕЛИКІ слова, які в журналі є службовими, а не прізвищами.
_LOG_UPPER_KEEP = {
    "УВАГА", "ГЕНЕРАЦІЯ", "ВИТЯГІВ", "ВИТЯГ", "РОЗРАХУНОК", "РОЗСИЛКИ", "СТАТИСТИКА",
    "ТА", "ДРУКУ", "ОТЦК", "ТЦК", "РТЦК", "МТЦК", "ОМТЦК", "СП", "АК", "ВЧ", "ЗС",
    "ПОМИЛКА", "ЗБІЙ", "ГОТОВО", "DOCX", "XLSX", "COM", "ПІБ", "КОМУ", "КУДИ",
}


def redact_sensitive_text(text: str) -> str:
    """Знеособлює журнал: назви частин, шифри, ПІБ і шляхи — на стабільні мітки.

    Потрібно, щоб журнал збою можна було комусь показати. Структура (скільки
    витягів, скільки сторінок, які попередження, скільки часу) лишається
    незмінною — саме вона й потрібна для розбору. Мітки стабільні в межах
    одного копіювання: та сама частина всюди буде тим самим «Ч-1».
    """
    counters: dict[str, dict[str, str]] = {}

    def token(prefix: str, value: str) -> str:
        bucket = counters.setdefault(prefix, {})
        key = value.casefold().strip()
        if key not in bucket:
            bucket[key] = f"{prefix}-{len(bucket) + 1}"
        return bucket[key]

    def replace(match: re.Match) -> str:
        if match.lastgroup == "path":
            return "<шлях>"
        if match.lastgroup == "otck":
            return token("ОТЦК", match.group(0))
        if match.lastgroup == "cipher":
            return token("Ш", match.group(0))
        if match.lastgroup == "name":
            return token("ПІБ", match.group(0))
        if match.lastgroup == "unit":
            word = match.group("unit_word").casefold().rstrip(".")
            if word in _LOG_MEASURE_WORDS:
                return match.group(0)
            return token("Ч", match.group(0))
        word = match.group(0)
        return word if word in _LOG_UPPER_KEEP else token("ПІБ", word)

    # ОДИН прохід: інакше мітка, яка сама містить число («ОТЦК-1»), потрапляла
    # під наступне правило й перетворювалась на «ОТЦК-Ч-2».
    text = _REDACT_RE.sub(replace, text)
    # Ім'я з великої літери поруч із міткою ПІБ («полковник Петро ПІБ-1»).
    text = re.sub(r"\b[А-ЯІЇЄҐ][а-яіїєґ'’\-]+(?=\s+ПІБ-\d)", "<імʼя>", text)
    return text


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
    source_font = source_paragraph.Range.Font
    source_format = source_paragraph.Range.ParagraphFormat

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

    destination = doc.Range(start, end)
    for prop, value in font_values.items():
        try:
            setattr(destination.Font, prop, value)
        except Exception:
            pass
    for index in range(1, destination.Paragraphs.Count + 1):
        dest_format = destination.Paragraphs(index).Range.ParagraphFormat
        try:
            dest_format.PageBreakBefore = False
        except Exception:
            pass
        for prop, value in geometry.items():
            try:
                setattr(dest_format, prop, value)
            except Exception:
                pass


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Node Automation Toolkit — Генератор витягів та примірників")
        self.root.geometry("1180x880")
        self.root.minsize(920, 700)

        # Конфігураційні змінні — Вкладка 1: Розрахунок та Витяги
        self.excel_path = tk.StringVar()
        self.doc_path = tk.StringVar()
        self.template_path = tk.StringVar()
        self.out_folder = tk.StringVar()
        self.executor = tk.StringVar()
        self.group_corps_var = tk.BooleanVar(value=True)
        self.duplex_2up_layout = tk.BooleanVar(value=True)

        # Підписант оригіналу наказу (Командувач/Командир — зчитується автоматично з наказу)
        self.order_signer_position = tk.StringVar()
        self.order_signer_rank = tk.StringVar()
        self.order_signer_name = tk.StringVar()

        # Особа, яка засвідчує витяг («Згідно з оригіналом» / Засвідчувач)
        self.certifier_position = tk.StringVar(value="Т.в.о. начальника штабу – першого заступника командувача військ")
        self.certifier_rank = tk.StringVar(value="полковник")
        self.certifier_name = tk.StringVar()
        self.p2_certifier_position = tk.StringVar()
        self.p2_certifier_rank = tk.StringVar()
        self.p2_certifier_name = tk.StringVar()
        self.manual_order_paths: list[str] = []
        self.p2_manual_order_paths: list[str] = []
        self.source_summary = tk.StringVar(value="Джерело не обрано — оберіть наказ або кілька наказів")
        self.message_source_summary = tk.StringVar(value="Наказ для повідомлень не обрано")
        self.p2_source_summary = tk.StringVar(value="Накази не обрано")
        # Вікно зразків одне на програму: друге таке ж лише плодило б
        # розбіжні значення в тих самих полях.
        self._samples_window = None
        self._batch_running = False

        # Конфігураційні змінні — Вкладка 2: Примірники 2/3
        self.p2_source_mode = tk.StringVar(value="folder")
        self.p2_orders_folder = tk.StringVar()
        self.p2_single_file = tk.StringVar()
        self.p2_back_page_path = tk.StringVar()
        self.p2_out_folder = tk.StringVar()
        # Чи обрано папку результату вручну. Якщо ні — вона щоразу
        # переобчислюється від поточного джерела, інакше примірники
        # продовжували б писатись у папку попереднього наказу.
        self.p2_out_folder_manual = tk.BooleanVar(value=False)
        self.p2_copy_title = tk.StringVar(value="Примірник № 2")
        # Виконавець примірників. Порожнє поле = береться виконавець витягів.
        self.p2_executor = tk.StringVar()
        # Режим превʼю: Word видимий, після кожного кроку пауза, у журналі —
        # що саме робиться. Пауза рядком, бо це поле вводить користувач.
        self.p2_preview = tk.BooleanVar(value=False)
        self.p2_preview_delay = tk.StringVar(value=str(PREVIEW_DEFAULT_DELAY))
        self.last_copy_two_paths: list[str] = []

        # Конфігурація вкладки «Накази» — генератор наказів по особовому складу.
        # «new_order_» означає наказ, який ми складаємо; «order_signer_» вище —
        # підписант чужого наказу, з якого робимо витяги.
        self.new_order_archive_folder = tk.StringVar()
        self.new_order_index_folder = tk.StringVar()
        self.new_order_index_force = tk.BooleanVar(value=False)
        self.new_order_plan_path = tk.StringVar()
        self.new_order_document_paths: list[str] = []
        self.new_order_source_summary = tk.StringVar(
            value="Джерела не обрано — план переміщення або документи"
        )
        self.new_order_template_path = tk.StringVar()
        self.new_order_out_folder = tk.StringVar()
        self.new_order_executor = tk.StringVar()
        self.new_order_number = tk.StringVar()
        self.new_order_date = tk.StringVar()
        self.new_order_section = tk.StringVar()
        self.new_order_points = tk.StringVar(value="пункту ___")
        self.new_order_kind = tk.StringVar(value="осіб офіцерського складу")
        self.new_order_unit = tk.StringVar()
        self.new_order_target_unit = tk.StringVar()
        self.new_order_bases = tk.StringVar()
        self.new_order_footer_bases = tk.StringVar()
        self.new_order_signer_position = tk.StringVar()
        self.new_order_signer_rank = tk.StringVar()
        self.new_order_signer_name = tk.StringVar()

        # Конфігурація вкладки «Повідомлення».
        self.message_cover_template_path = tk.StringVar()
        self.message_content_template_path = tk.StringVar()
        self.message_out_folder = tk.StringVar()
        self.message_executor = tk.StringVar()

        # Смуга прогресу генерації (спільна для всіх вкладок)
        self.progress_value = tk.DoubleVar(value=0.0)
        self.progress_text = tk.StringVar(value="")
        self._progress_total = 0
        self._progress_done = 0
        self._progress_title = ""
        self._progress_started_at = None

        self.config_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")

        self.load_config()
        self.create_widgets()

    def create_widgets(self):
        """Інтерфейс будує Qt-оболонка (`QtShellMixin.create_widgets`)."""
        raise NotImplementedError("Запускайте генератор через generate_extracts_qt.py")

    def handle_drag_and_drop(self, raw_files):
        """Обробка перетягування файлів мишею (Drag-and-Drop) у вікно програми."""
        if not raw_files:
            return

        files = []
        for item in raw_files:
            if isinstance(item, bytes):
                try:
                    path_str = item.decode("utf-8")
                except UnicodeDecodeError:
                    path_str = item.decode("mbcs", errors="ignore")
            else:
                path_str = str(item)
            if os.path.exists(path_str):
                files.append(os.path.abspath(path_str))

        if not files:
            return

        # Визначаємо активну вкладку
        current_tab = self.notebook.index(self.notebook.select())
        handled_count = 0

        TAB_COPIES = 0
        TAB_EXTRACTS = 1
        TAB_MESSAGES = 2

        # Накази збираються в СПИСОК: перетягнути одразу кілька — звичайна
        # справа, а старий обробник лишав тільки останній файл.
        dropped_orders: list[str] = []

        for fpath in files:
            fname = os.path.basename(fpath).lower()
            if os.path.isdir(fpath):
                if current_tab == TAB_COPIES:
                    self.p2_orders_folder.set(fpath)
                    self.p2_source_mode.set("folder")
                    self._on_p2_source_mode_changed()
                    self.log_p2(f"📥 [Drag-and-Drop] Папку наказів встановлено: {fpath}")
                elif current_tab == TAB_MESSAGES:
                    orders_in_folder = self._orders_in_folder(fpath)
                    if orders_in_folder:
                        self._set_orders([orders_in_folder[0]])
                        self.log(
                            f"📥 [Drag-and-Drop] Наказ для повідомлень: {os.path.basename(orders_in_folder[0])}. "
                            "Повідомлення одинарні; з теки взято лише перший наказ."
                        )
                    else:
                        self.log("УВАГА: у перетягнутій теці немає наказів DOCX.")
                else:
                    orders_in_folder = self._orders_in_folder(fpath)
                    if orders_in_folder:
                        self._set_orders(orders_in_folder)
                        self.log(f"📥 [Drag-and-Drop] З теки взято наказів: {len(orders_in_folder)}")
                    else:
                        self.out_folder.set(fpath)
                        self.log(f"📥 [Drag-and-Drop] Папку результату встановлено: {fpath}")
                handled_count += 1
            elif fname.endswith((".xlsx", ".xls")):
                self.excel_path.set(fpath)
                self.log(f"📥 [Drag-and-Drop] Словник Excel встановлено: {os.path.basename(fpath)}")
                handled_count += 1
            elif fname.endswith(".docx") and not fname.startswith("~$"):
                if current_tab == TAB_COPIES and ("задн" in fname or "back" in fname or "шаблон" in fname):
                    self.p2_back_page_path.set(fpath)
                    self.log_p2(f"📥 [Drag-and-Drop] Заготовку примірника встановлено: {os.path.basename(fpath)}")
                elif current_tab == TAB_EXTRACTS and ("шаблон" in fname or "template" in fname or "зразок" in fname):
                    self.template_path.set(fpath)
                    self.log(f"📥 [Drag-and-Drop] Зразок витягу встановлено: {os.path.basename(fpath)}")
                elif current_tab == TAB_MESSAGES and ("зміст" in fname or "content" in fname):
                    self.message_content_template_path.set(fpath)
                    self.log(f"📥 [Drag-and-Drop] Зразок змісту повідомлення встановлено: {os.path.basename(fpath)}")
                elif current_tab == TAB_MESSAGES and ("титул" in fname or "cover" in fname or "шаблон" in fname):
                    self.message_cover_template_path.set(fpath)
                    self.log(f"📥 [Drag-and-Drop] Зразок супроводу встановлено: {os.path.basename(fpath)}")
                else:
                    dropped_orders.append(fpath)
                handled_count += 1

        if dropped_orders:
            if current_tab == TAB_COPIES:
                self._set_p2_orders(dropped_orders)
                self.log_p2(f"📥 [Drag-and-Drop] Наказів для примірників: {len(dropped_orders)}")
            elif current_tab == TAB_MESSAGES:
                self._set_orders([dropped_orders[0]])
                self.log(
                    f"📥 [Drag-and-Drop] Наказ для повідомлень: {os.path.basename(dropped_orders[0])}"
                    + (f". Решту файлів ({len(dropped_orders) - 1}) не взято: повідомлення одинарні."
                       if len(dropped_orders) > 1 else "")
                )
            else:
                self._set_orders(dropped_orders)
                self.log(f"📥 [Drag-and-Drop] Наказів до обробки: {len(dropped_orders)}")

        if handled_count > 0:
            self.save_config()

    def load_config(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.excel_path.set(data.get("excel_path", ""))
                    self.doc_path.set(data.get("doc_path", ""))
                    self.template_path.set(data.get("template_path", ""))
                    self.out_folder.set(data.get("out_folder", ""))
                    self.executor.set(data.get("executor", ""))
                    self.certifier_position.set(
                        data.get("certifier_position", "Т.в.о. начальника штабу – першого заступника командувача військ")
                    )
                    self.certifier_rank.set(data.get("certifier_rank", "полковник"))
                    self.certifier_name.set(data.get("certifier_name", ""))
                    # Разова міграція старих спільних полів; далі незалежні.
                    for field in ("position", "rank", "name"):
                        getattr(self, f"p2_certifier_{field}").set(
                            data.get(f"p2_certifier_{field}", getattr(self, f"certifier_{field}").get())
                        )
                    if "group_corps" in data:
                        self.group_corps_var.set(data["group_corps"])
                    if "duplex_2up_layout" in data:
                        self.duplex_2up_layout.set(data["duplex_2up_layout"])

                    self.p2_orders_folder.set(data.get("p2_orders_folder", ""))
                    self.p2_single_file.set(data.get("p2_single_file", ""))
                    self.manual_order_paths = [
                        path for path in data.get("manual_order_paths", []) if os.path.isfile(path)
                    ]
                    self.p2_manual_order_paths = [
                        path for path in data.get("p2_manual_order_paths", []) if os.path.isfile(path)
                    ]
                    # Режим джерела примірників раніше не зберігався: після
                    # перезапуску програма поверталась до «теки» й мовчки
                    # ігнорувала обрані файли. Для старого конфігу режим
                    # виводимо з того, що в ньому є.
                    saved_mode = data.get("p2_source_mode")
                    if saved_mode in ("folder", "file"):
                        self.p2_source_mode.set(saved_mode)
                    elif self.p2_manual_order_paths or os.path.isfile(self.p2_single_file.get()):
                        self.p2_source_mode.set("file")
                    self.p2_back_page_path.set(data.get("p2_back_page_path", ""))
                    self.p2_out_folder.set(data.get("p2_out_folder", ""))
                    self.p2_out_folder_manual.set(bool(data.get("p2_out_folder_manual", False)))
                    # Правила для примірника № 3 ще не погоджені.
                    self.p2_copy_title.set("Примірник № 2")
                    self.p2_executor.set(data.get("p2_executor", ""))
                    self.p2_preview.set(bool(data.get("p2_preview", False)))
                    self.p2_preview_delay.set(
                        str(data.get("p2_preview_delay", PREVIEW_DEFAULT_DELAY))
                    )

                    for field in (
                        "archive_folder", "index_folder", "plan_path", "template_path",
                        "out_folder", "executor", "number", "date", "section", "unit",
                        "target_unit", "bases", "footer_bases",
                        "signer_position", "signer_rank", "signer_name",
                    ):
                        value = data.get(f"new_order_{field}")
                        if value:
                            getattr(self, f"new_order_{field}").set(value)
                    for field, default in (("points", "пункту ___"), ("kind", "осіб офіцерського складу")):
                        getattr(self, f"new_order_{field}").set(data.get(f"new_order_{field}") or default)
                    self.new_order_document_paths = [
                        path for path in data.get("new_order_document_paths", []) if os.path.isfile(path)
                    ]

                    self.message_cover_template_path.set(data.get("message_cover_template_path", ""))
                    self.message_content_template_path.set(data.get("message_content_template_path", ""))
                    self.message_out_folder.set(data.get("message_out_folder", ""))
                    self.message_executor.set(data.get("message_executor", ""))

            except Exception:
                pass

    def save_config(self):
        data = {
            "excel_path": self.excel_path.get(),
            "doc_path": self.doc_path.get(),
            "template_path": self.template_path.get(),
            "out_folder": self.out_folder.get(),
            "executor": self.executor.get(),
            "certifier_position": self.certifier_position.get(),
            "certifier_rank": self.certifier_rank.get(),
            "certifier_name": self.certifier_name.get(),
            "p2_certifier_position": self.p2_certifier_position.get(),
            "p2_certifier_rank": self.p2_certifier_rank.get(),
            "p2_certifier_name": self.p2_certifier_name.get(),
            "group_corps": self.group_corps_var.get(),
            "duplex_2up_layout": self.duplex_2up_layout.get(),
            "p2_orders_folder": self.p2_orders_folder.get(),
            "p2_single_file": self.p2_single_file.get(),
            "p2_source_mode": self.p2_source_mode.get(),
            "manual_order_paths": self.manual_order_paths,
            "p2_manual_order_paths": self.p2_manual_order_paths,
            "p2_back_page_path": self.p2_back_page_path.get(),
            "p2_out_folder": self.p2_out_folder.get(),
            "p2_out_folder_manual": self.p2_out_folder_manual.get(),
            "p2_copy_title": self.p2_copy_title.get(),
            "p2_executor": self.p2_executor.get(),
            "p2_preview": self.p2_preview.get(),
            "p2_preview_delay": self.p2_preview_delay.get(),
            **{
                f"new_order_{field}": getattr(self, f"new_order_{field}").get()
                for field in (
                    "archive_folder", "index_folder", "plan_path", "template_path",
                    "out_folder", "executor", "number", "date", "section", "points",
                    "kind", "unit", "target_unit", "bases", "footer_bases",
                    "signer_position", "signer_rank", "signer_name",
                )
            },
            "new_order_document_paths": list(self.new_order_document_paths),
            "message_cover_template_path": self.message_cover_template_path.get(),
            "message_content_template_path": self.message_content_template_path.get(),
            "message_out_folder": self.message_out_folder.get(),
            "message_executor": self.message_executor.get(),
        }
        try:
            with open(self.config_file, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
        except Exception:
            pass

    # -------------------------------------------------------------------------
    # СМУГА ПРОГРЕСУ ГЕНЕРАЦІЇ
    # -------------------------------------------------------------------------
    def progress_begin(self, total: int, title: str) -> None:
        """Починає показ прогресу: `total` кроків під назвою `title`."""
        self._progress_total = max(1, int(total))
        self._progress_done = 0
        self._progress_title = title
        self._progress_started_at = time.monotonic()
        self._progress_render(f"{title}: 0 / {self._progress_total}", 0.0)

    def progress_step(self, done: int, detail: str = "") -> None:
        """Показує, що виконано `done` кроків із загальної кількості."""
        total = getattr(self, "_progress_total", 0)
        if not total:
            return
        self._progress_done = max(0, min(int(done), total))
        title = getattr(self, "_progress_title", "Обробка")
        caption = f"{title}: {self._progress_done} / {total}"
        if detail:
            caption += f" — {detail}"
        self._progress_render(caption, 100.0 * self._progress_done / total)

    def progress_end(self, detail: str = "") -> None:
        """Завершує показ прогресу.

        З підсумком (`detail`) смуга лишається заповненою — це успішне
        завершення. Без підсумку виклик означає, що робота урвалась (виняток,
        закритий діалог), тож смуга застигає там, де була, з поміткою.
        """
        total = getattr(self, "_progress_total", 0)
        if not total:
            return  # прогрес уже завершено успішним викликом
        started_at = getattr(self, "_progress_started_at", None)
        done = getattr(self, "_progress_done", 0)
        caption = detail or "Перервано"
        if started_at is not None:
            caption += f" ({time.monotonic() - started_at:.0f} с)"
        percent = 100.0 if detail else (100.0 * done / total)
        self._progress_total = 0
        self._progress_render(caption, percent)

    def _progress_render(self, caption: str, percent: float) -> None:
        """Малює стан смуги. Мовчить, якщо інтерфейс не побудовано (тести, e2e)."""
        try:
            self.progress_text.set(caption)
            self.progress_value.set(percent)
            self.root.update_idletasks()
        except Exception:
            pass

    def _on_p2_source_mode_changed(self):
        """Сумісність із Drag-and-Drop: після зміни джерела оновлюємо список."""
        self._refresh_p2_orders_view()

    def show_template_tags(self):
        messagebox.showinfo(
            "Теги шаблону витягу",
            "Додайте ці теги до DOCX-шаблону витягу:\n\n"
            "• {{кому}} — значення «Кому» з Excel; якщо порожнє — шифр.\n"
            "• {{куди}} — значення «Куди» з Excel.\n"
            "• {{номер_наказу}} або {{номер}} — номер лише з назви файла наказу.\n"
            "• {{дата_наказу}} або {{дата}} — дата лише з назви файла наказу.\n"
            "• {{пункти}} — перелік пунктів цього витягу.\n"
            "• {{зміст}} — повний текст пунктів з оригіналу наказу.\n\n"
            "Підписант оригіналу наказу (зчитується автоматично після пунктів):\n"
            "• {{підписант_посада}} (або {{посада_підписанта}})\n"
            "• {{підписант_звання}} (або {{звання_підписанта}})\n"
            "• {{підписант_піб}} (або {{підписант_імя}}, {{прізвище_підписанта}})\n\n"
            "• {{підписант}} — увесь блок: посада, а нижче звання та ім’я.\n\n"
            "Засвідчення («Згідно з оригіналом»):\n"
            "• {{згідно_з_оригіналом}} або {{засвідчення}} → «Згідно з оригіналом»\n"
            "• {{засвідчувач_посада}} (або {{затверджувач_посада}}, {{згідно_з_оригіналом_посада}})\n"
            "• {{засвідчувач_звання}} (або {{затверджувач_звання}}, {{згідно_з_оригіналом_звання}})\n"
            "• {{засвідчувач_піб}} (або {{затверджувач_піб}}, {{згідно_з_оригіналом_піб}}, {{засвідчувач}})\n\n"
            "Назви в дужках рівноцінні: підставляється те саме значення.\n"
            "Засвідчувач витягів і засвідчувач примірників — окремі поля у вікні «Зразки».\n\n"
            "Витяги до управління:\n"
            "• той самий зразок, але {{кому}} та {{куди}} лишаються порожніми;\n"
            "• складаються в окремий файл і не входять у розрахунок розсилки.\n\n"
            "Службові реквізити:\n"
            "• {{виконавець}} — виконавець з форми програми.",
        )

    def show_p2_info(self):
        messagebox.showinfo(
            "Правила формування примірника № 2",
            "Примірники № 2/3 збираються з DOCX-заготовки. У {{зміст}} переноситься "
            "тіло наказу, а службова остання сторінка лишається із заготовки.\n\n"
            "Основні теги заготовки:\n"
            "• {{зміст}} — тіло наказу (обов’язковий тег)\n"
            "• {{підписант}} або окремі теги {{підписант_посада}}, "
            "{{підписант_звання}}, {{підписант_імя}}\n"
            "• {{згідно_з_оригіналом}} або {{засвідчення}}\n"
            "• {{номер_наказу}} або {{номер}} — номер із назви файла\n"
            "• {{дата_наказу}} або {{дата}} — дата з назви файла\n"
            "• {{примірник}} або {{примірник_номер}} — номер примірника\n\n"
            "Якщо номер або дату не знайдено в назві файла, відповідний тег лишається "
            "для ручного заповнення. Заготовка може бути багатосторінковою.",
        )

    def _populate_result_tab(self, key: str, columns: list[str], rows: list[tuple]):
        tree = self.result_views[key]
        for item in tree.get_children():
            tree.delete(item)
        tree["columns"] = columns
        for position, title in enumerate(columns):
            values = [str(row[position]) if position < len(row) else "" for row in rows]
            width = max(110, min(420, max([len(title), *(len(value) for value in values)] or [len(title)]) * 8))
            tree.heading(title, text=title)
            tree.column(title, width=width, minwidth=90, stretch=True, anchor=W)
        if key == "calculation" and hasattr(tree, "tag_configure"):
            tree.tag_configure(
                "management",
                background="#d1fae5",
                foreground="#065f46",
                font=("Segoe UI", 9, "bold"),
            )
        for row in rows:
            values = tuple("-" if value in (None, "") else str(value) for value in row)
            tags = (
                ("management",)
                if key == "calculation" and values and values[0] == MANAGEMENT_RESULT_LABEL
                else ()
            )
            tree.insert("", tk.END, values=values, tags=tags)

    def show_analysis_results(self, map_result: dict):
        units_table = map_result.get("units_table")
        calculation_rows = []
        if units_table and hasattr(units_table, "rows"):
            calculation_rows = [(row[0], row[2], row[3]) for row in units_table.rows]
        calculation_rows.extend(management_calculation_rows(map_result))
        self._populate_result_tab(
            "calculation",
            ["Військова частина / Відправник", "Пункти витягу", "Кількість"],
            calculation_rows,
        )

        unmatched_rows = [
            (item.get("label", "-"), item.get("text", ""), item.get("reason", ""))
            for item in map_result.get("unmatched_items", [])
        ]
        self._populate_result_tab("unmatched", ["Пункт", "Текст пункту", "Причина"], unmatched_rows)

        routing_rows = [
            (
                item.get("label", "-"),
                item.get("matched_entries", "-"),
                item.get("applied_rules", "-"),
                item.get("final_recipients", "-"),
            )
            for item in map_result.get("routing_audit", [])
        ]
        self._populate_result_tab(
            "routing",
            ["Пункт", "Збіги з таблиці", "Застосовані правила", "Підсумкові адресати"],
            routing_rows,
        )
        self._populate_result_tab("layout", ["Стан макета"], [])

    def show_layout_warnings(self, warnings: list[str]):
        rows = [(warning,) for warning in warnings]
        if not rows:
            rows = [("Макет сформовано без попереджень.",)]
        self._populate_result_tab("layout", ["Стан макета"], rows)
        self.results_notebook.select(self.result_tabs["layout" if warnings else "calculation"])

    # ── Вибір файлів ─────────────────────────────────────────────────────────
    def select_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.excel_path.set(path)
            self.save_config()

    def select_template(self):
        path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx *.doc")])
        if path:
            self.template_path.set(path)
            self.save_config()

    def edit_template(self):
        path = self.template_path.get()
        if not path or not os.path.exists(path):
            messagebox.showwarning("Помилка", "Спочатку виберіть існуючий файл шаблону!")
            return
        try:
            self.log("Відкриваємо шаблон витягу у Word для редагування...")
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = True
            doc = word.Documents.Open(os.path.abspath(path))
            doc.Activate()
        except Exception as e:
            self.log(f"Помилка відкриття шаблону: {str(e)}")

    def select_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.out_folder.set(path)
            self.save_config()

    def select_message_cover_template(self):
        path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx *.doc")])
        if path:
            self.message_cover_template_path.set(path)
            self.save_config()

    def select_message_content_template(self):
        path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx *.doc")])
        if path:
            self.message_content_template_path.set(path)
            self.save_config()

    def select_message_output_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.message_out_folder.set(path)
            self.save_config()

    def show_message_tags(self):
        messagebox.showinfo(
            "Теги повідомлень",
            "Спільні для обох шаблонів:\n"
            "{{номер_наказу}} / {{номер}}, {{дата_наказу}} / {{дата}}, "
            "{{кому_список}}, {{куди}}, {{виконавець}}.\n"
            "Також діють спільні теги підписанта й засвідчувача/затверджувача, "
            "як у витягах і примірниках.\n\n"
            "Розмістіть {{кому_список}} у кожному рядку таблиці окремо: один тег "
            "отримує одного унікального адресата. Якщо рядків замало, програма "
            "додає перед наступним блоком копії останнього рядка таблиці — штамп "
            "зсувається вниз і не перекривається. Порядок адресатів: корпуси, "
            "частини, лише обласні ТЦК.\n\n"
            "{{тцк чі вч}} — тип адресата, визначається автоматично:\n"
            "• лише військові частини → «командирам військових частин»;\n"
            "• лише ТЦК → «начальникам ОТЦК та СП»;\n"
            "• і те, і те → «командирам військових частин та начальникам ОТЦК та СП».\n\n"
            "У шаблоні зі змістом:\n"
            "{{зміст_шифр}} — пункти наказу із заміною відкритих назв частин "
            "на шифри з Excel. Назви, які лишилися без збігу, виділяються жовтим.\n\n"
            "{{виконавець}} підставляється рівно так, як введено у формі, і "
            "розміщується за 3 см від низу сторінки, якщо тег поза таблицею. "
            "Службові рядки («ВІДКРИТА ІНФОРМАЦІЯ» тощо) програма НЕ додає — "
            "їх треба прописати у самому зразку; текст шаблону після "
            "{{виконавець}} зберігається без змін.\n\n"
            "Номер наказу береться з назви файла, дата — теж, у вигляді "
            "20.05.2025 року (без розкриття словами)."
        )

    @staticmethod
    def _replace_message_tags(document, replacements: dict[str, str]) -> list[tuple[int, int]]:
        """Замінює теги й повертає діапазони завершальних блоків виконавця."""
        executor_blocks = []
        for tag, value in expand_common_tags(replacements).items():
            find_obj = document.Content.Find
            find_obj.Text = tag
            iterations = 0
            while find_obj.Execute() and iterations < 100:
                iterations += 1
                replacement_start = find_obj.Parent.Start
                # Виконавець підставляється рівно так, як його ввели.
                # Рядки «ВІДКРИТА ІНФОРМАЦІЯ» та «(Обмежено в розповсюдженні…)»
                # програма НЕ додає: вони прописані у самому зразку окремо.
                replacement = value
                find_obj.Parent.Text = replacement
                if tag == "{{виконавець}}":
                    # Текст шаблону після виконавця більше не видаляється:
                    # там лежать службові рядки зразка, які мають зберігатися.
                    executor_blocks.append((replacement_start, replacement_start + len(replacement)))
                find_obj = document.Content.Find
                find_obj.Text = tag
        return executor_blocks

    @staticmethod
    def _replace_executor_in_headers(document, value: str) -> int:
        """Підставляє `{{виконавець}}` у КОЛОНТИТУЛАХ і повертає кількість замін.

        Колонтитул — окрема «історія» Word: `document.Content` бачить лише
        основну частину документа, тому тег, поставлений у колонтитул зразка,
        лишався незаповненим і друкувався як є.

        Позицію тут не підбираємо: колонтитул і так унизу сторінки, а блок
        виконавця в тілі документа працює як раніше — його закладку й
        вирівнювання ця функція не чіпає.
        """
        tag = "{{виконавець}}"
        replaced = 0
        try:
            sections = list(document.Sections)
        except TypeError:
            sections = [document.Sections(index) for index in range(1, document.Sections.Count + 1)]
        for section in sections:
            for collection in (section.Headers, section.Footers):
                try:
                    stories = list(collection)
                except TypeError:
                    # Первинний, першої сторінки та парних сторінок.
                    stories = [collection(index) for index in (1, 2, 3)]
                for header_footer in stories:
                    for _ in range(20):  # запобіжник від вічного циклу
                        try:
                            find_obj = header_footer.Range.Find
                            find_obj.Text = tag
                            if not find_obj.Execute():
                                break
                            find_obj.Parent.Text = value
                        except Exception:
                            break
                        replaced += 1
        return replaced

    @staticmethod
    def _position_message_executor_at_page_bottom(document, bookmark_name: str):
        """Тримає завершальний блок повідомлення за 3 см віднизу сторінки."""
        if not document.Bookmarks.Exists(bookmark_name):
            return
        block_range = document.Bookmarks(bookmark_name).Range
        executor_start = block_range.Start
        executor_index = block_last_index = None
        for index in range(1, document.Paragraphs.Count + 1):
            paragraph_range = document.Paragraphs(index).Range
            if paragraph_range.Start <= executor_start < paragraph_range.End:
                executor_index = index
            block_end = max(block_range.Start, block_range.End - 1)
            if paragraph_range.Start <= block_end < paragraph_range.End:
                block_last_index = index
            if executor_index is not None and block_last_index is not None:
                break
        if executor_index is None or block_last_index is None:
            return

        try:
            # У таблиці завершальний блок не рухаємо: там відступи контролює шаблон.
            if document.Paragraphs(executor_index).Range.Information(12):  # wdWithInTable
                return
            for index in range(executor_index, block_last_index + 1):
                paragraph_format = document.Paragraphs(index).Range.ParagraphFormat
                paragraph_format.KeepTogether = True
                paragraph_format.KeepWithNext = index < block_last_index
            while executor_index > 1:
                previous = document.Paragraphs(executor_index - 1).Range
                if previous.Text.replace("\r", "").replace("\x07", "").strip():
                    break
                previous.Delete()
                executor_index -= 1
                block_last_index -= 1

            document.Repaginate()
            page_number = document.Paragraphs(executor_index).Range.Information(3)  # wdActiveEndPageNumber
            target_vertical = document.PageSetup.PageHeight - document.Application.CentimetersToPoints(3)
            if document.Paragraphs(block_last_index).Range.Information(3) != page_number:
                return
            # Оптимізоване наближення замість 120 покрокових Repaginate
            block_vertical = document.Paragraphs(block_last_index).Range.Information(6)  # wdVerticalPositionRelativeToPage
            vertical_diff = target_vertical - block_vertical
            if vertical_diff > 25:
                batch_enters = max(1, int(vertical_diff / 18))
                pos = document.Paragraphs(executor_index).Range.Start
                document.Range(pos, pos).InsertBefore("\r" * batch_enters)
                executor_index += batch_enters
                block_last_index += batch_enters
                document.Repaginate()
                # Перевірка: якщо пакетна вставка виштовхнула блок на нову сторінку — відкат
                batch_page = document.Paragraphs(block_last_index).Range.Information(3)
                if batch_page != page_number:
                    for _ in range(batch_enters):
                        document.Paragraphs(executor_index - 1).Range.Delete()
                        executor_index -= 1
                        block_last_index -= 1
                    document.Repaginate()

            for _ in range(15):
                block_vertical = document.Paragraphs(block_last_index).Range.Information(6)  # wdVerticalPositionRelativeToPage
                if block_vertical >= target_vertical:
                    break
                position = document.Paragraphs(executor_index).Range.Start
                document.Range(position, position).InsertBefore("\r")
                executor_index += 1
                block_last_index += 1
                document.Repaginate()
                block_page = document.Paragraphs(block_last_index).Range.Information(3)
                block_vertical = document.Paragraphs(block_last_index).Range.Information(6)
                if block_page != page_number or block_vertical > target_vertical:
                    document.Paragraphs(executor_index - 1).Range.Delete()
                    document.Repaginate()
                    break
        except Exception:
            # Верстка виконавця не має блокувати створення повідомлення.
            return

    @staticmethod
    def _find_recipient_table_cell(document, position: int):
        """Повертає таблицю, рядок і комірку, що містять задану позицію Word.

        Використовує table.Range.Cells для безпечного обходу таблиць
        з вертикально об'єднаними комірками без збоїв COM.
        """
        try:
            for table_index in range(1, document.Tables.Count + 1):
                table = document.Tables(table_index)
                for cell_index in range(1, table.Range.Cells.Count + 1):
                    cell = table.Range.Cells(cell_index)
                    cell_range = cell.Range
                    if cell_range.Start <= position < cell_range.End:
                        return table, cell.RowIndex, cell.ColumnIndex
        except Exception:
            pass
        return None

    @staticmethod
    def _fill_recipient_slots(document, recipients: list[str]) -> tuple[int, int]:
        """Заповнює кожен {{кому_список}} одним наступним адресатом.

        Теги зазвичай лежать у рядках таблиці поруч зі штампом. Тому список
        не вставляється в одну комірку: жоден адресат не повторюється. Якщо
        готових рядків бракує, нижче останнього рядка-адресата копіюється його
        форматований рядок. Нові рядки вставляються до наступного блока, тож
        штамп зсувається вниз, а не перекривається текстом.
        """
        slot_count = 0
        last_row_context = None
        find_obj = document.Content.Find
        find_obj.Text = "{{кому_список}}"
        iterations = 0
        while find_obj.Execute() and iterations < 500:
            iterations += 1
            found_range = find_obj.Parent.Duplicate
            last_row_context = App._find_recipient_table_cell(document, found_range.Start) or last_row_context
            value = recipients[slot_count] if slot_count < len(recipients) else ""
            find_obj.Parent.Text = value
            slot_count += 1
            find_obj = document.Content.Find
            find_obj.Text = "{{кому_список}}"

        if slot_count >= len(recipients) or last_row_context is None:
            return slot_count, max(0, len(recipients) - slot_count)

        table, row_index, column_index = last_row_context
        remaining = recipients[slot_count:]
        try:
            source_row = table.Rows(row_index)
            after_address_rows = table.Rows.Count > row_index
            anchor_row = table.Rows(row_index + 1) if after_address_rows else None
            source_value = recipients[slot_count - 1] if slot_count else ""
            for recipient in remaining:
                new_row = table.Rows.Add(anchor_row) if anchor_row else table.Rows.Add()
                source_cells = source_row.Cells.Count
                for cell_index in range(1, min(source_cells, new_row.Cells.Count) + 1):
                    source_range = source_row.Cells(cell_index).Range.Duplicate
                    source_range.End -= 1  # Службовий знак кінця комірки Word.
                    target_range = new_row.Cells(cell_index).Range.Duplicate
                    target_range.End -= 1
                    target_range.FormattedText = source_range.FormattedText

                target_cell = new_row.Cells(column_index)
                target_find = target_cell.Range.Find
                target_find.Text = source_value
                if target_find.Execute():
                    target_find.Parent.Text = recipient
                else:
                    # Резервний варіант для порожнього або нестандартно
                    # відформатованого рядка таблиці.
                    cell_range = target_cell.Range.Duplicate
                    cell_range.End -= 1
                    cell_range.Text = recipient
                source_row = new_row
                source_value = recipient
                slot_count += 1
        except Exception:
            # Якщо Word не дає продублювати конкретний рядок, повертаємо
            # невміщені адресати як попередження, не вставляючи їх поверх штампа.
            return slot_count, len(recipients) - slot_count
        return slot_count, 0

    def _copy_order_content(self, doc, tag_range, content_source: dict) -> int:
        """Переносить зміст наказу у повідомлення РАЗОМ ІЗ ФОРМАТУВАННЯМ.

        Копіюється `FormattedText` кожного абзацу — точно так само, як у
        витягах, тому стилі, вирівнювання та відступи лишаються такими ж, як
        в оригіналі наказу. Шифрування виконується поабзацно вже ПІСЛЯ
        копіювання, завдяки чому форматування абзацу зберігається (вставка
        простим текстом його втрачала й «кривила» верстку).

        Повертає кількість підсвічених невпізнаних назв частин.
        """
        source_doc = content_source["doc"]
        first_para = int(content_source["first_para"])
        last_para = int(content_source["last_para"])
        mapping = content_source.get("mapping") or {}

        # Валідація ДО будь-яких змін документа: якщо діапазон некоректний,
        # викликач ще може безпечно перейти на запасний спосіб вставки.
        total = source_doc.Paragraphs.Count
        if not 1 <= first_para <= last_para <= total:
            raise ValueError(
                f"некоректний діапазон абзаців наказу {first_para}-{last_para} із {total}"
            )

        insert_point = tag_range.Paragraphs(1).Range.Start
        tag_range.Paragraphs(1).Range.Delete()
        content_start = insert_point
        try:
            for p_index in range(first_para, last_para + 1):
                source_paragraph = source_doc.Paragraphs(p_index)
                source_range = source_paragraph.Range.Duplicate
                if "\x0c" in (source_range.Text or ""):
                    continue  # ручні розриви сторінок з наказу не переносимо
                destination = doc.Range(insert_point, insert_point)
                destination.FormattedText = source_range.FormattedText
                paragraph_start, insert_point = destination.Start, destination.End
                _carry_source_formatting(
                    doc, paragraph_start, insert_point, source_paragraph
                )
        except Exception as error:
            # Частина абзаців уже вставлена — повертатись до простого тексту
            # не можна, інакше зміст задвоївся б.
            raise RuntimeError(f"збій копіювання абзацу наказу: {error}") from error

        # Шифрування змінює довжину абзаців, тому межі змісту тримаємо
        # закладкою — після правок вона вкаже актуальний діапазон.
        # Порожні білі зображення наказу в повідомлення не переносимо.
        pasted_range = doc.Range(content_start, insert_point)
        removed_images = remove_blank_images(doc, pasted_range)
        if removed_images:
            insert_point = pasted_range.End
            self.log(f"Прибрано порожніх білих зображень: {removed_images}.")

        bookmark_name = "nat_message_content"
        try:
            doc.Bookmarks.Add(bookmark_name, doc.Range(content_start, insert_point))
        except Exception:
            bookmark_name = ""

        # У примірниках текст наказу лишається 1-в-1 (без шифрів), тому
        # шифрування вмикається лише там, де воно справді потрібне.
        highlighted = 0
        if content_source.get("cipher", True):
            highlighted = self._cipher_inserted_content(doc, content_start, insert_point, mapping)

        if bookmark_name and doc.Bookmarks.Exists(bookmark_name):
            content_range = doc.Bookmarks(bookmark_name).Range
            try:
                self._apply_message_layout_rules(doc, content_range.Start, content_range.End)
            except Exception as layout_error:
                self.log(f"УВАГА: не вдалося застосувати правила верстки змісту: {layout_error}")
            try:
                doc.Bookmarks(bookmark_name).Delete()
            except Exception:
                pass
        return highlighted

    def _apply_message_layout_rules(self, doc, content_start: int, content_end: int) -> None:
        """Нерозривність блоків і заповнення сторінки — як у витягах.

        1. Пункт разом зі своїм біографічним блоком (р.н., освіта, ІПН/РНОКПП,
           ВОС, «Підлягає направленню…») лишається неподільним: він не може
           розриватись між сторінками.
        2. Шапки (§, «Відповідно до …:», «У ЗАПАС ЗА ПІДПУНКТОМ …:») зчеплені
           з наступним пунктом.
        3. Міжрядковий інтервал добирається в діапазоні 16 → 14 пт так, щоб
           зміст займав якнайменше сторінок і не лишав напівпорожніх.
        """
        content_range = doc.Range(content_start, content_end)
        # Скидаємо KeepWithNext, який міг прийти разом із FormattedText.
        # Далі він виставляється лише за нашою структурою блоків.
        content_range.ParagraphFormat.KeepWithNext = False
        spans = []
        list_numbers = []
        for paragraph in content_range.Paragraphs:
            paragraph_range = paragraph.Range
            spans.append((paragraph_range.Start, paragraph_range.End))
            try:
                list_numbers.append(str(paragraph_range.ListFormat.ListString or ""))
            except Exception:
                list_numbers.append("")

        def paragraph_kind(text: str, list_number: str = "") -> str:
            clean = (text or "").strip()
            if not clean:
                return "blank"
            if clean.startswith("§"):
                return "heading"
            # У конвертованих DOCX номер часто є автонумерацією Word і не
            # входить у Range.Text. Без ListString усі такі пункти ставали
            # «продовженням» і злипались в один завеликий блок.
            if (
                re.match(r"^\d{1,3}(?:\.\d{1,3})*[\.\)]\s", clean)
                or re.match(r"^\d{1,3}(?:\.\d{1,3})*[\.\)]?$", list_number.strip())
            ):
                return "item"
            if clean.endswith(":"):
                return "heading"
            return "continuation"

        kinds = []
        for (start, end), list_number in zip(spans, list_numbers):
            kinds.append(paragraph_kind(doc.Range(start, end).Text, list_number))

        # Останній абзац кожної групи «пункт + біографія» не тягне наступний.
        for index, (start, end) in enumerate(spans):
            if kinds[index] == "blank":
                continue
            paragraph_format = doc.Range(start, end).ParagraphFormat
            paragraph_format.KeepTogether = True

            next_meaningful = next(
                (j for j in range(index + 1, len(kinds)) if kinds[j] != "blank"), None
            )
            if next_meaningful is None:
                keep_with_next = False
            elif kinds[index] == "heading":
                keep_with_next = True
            else:
                # Пункт тримає свій біографічний блок; новий пункт або шапка
                # починають окрему групу.
                keep_with_next = kinds[next_meaningful] == "continuation"
            paragraph_format.KeepWithNext = keep_with_next

            # Порожній абзац після шапки не повинен розривати її зв'язок із
            # першим пунктом. Порожній абзац після завершеного пункту, навпаки,
            # лишається дозволеною точкою розриву між двома пунктами.
            gap_end = next_meaningful if next_meaningful is not None else len(kinds)
            for blank_index in range(index + 1, gap_end):
                doc.Range(*spans[blank_index]).ParagraphFormat.KeepWithNext = keep_with_next

        # Заповнення сторінки: найбільший інтервал із діапазону, що дає
        # найменшу кількість сторінок.
        def set_spacing(points: float) -> None:
            for start, end in spans:
                fmt = doc.Range(start, end).ParagraphFormat
                fmt.LineSpacingRule = 4  # wdLineSpaceExactly
                fmt.LineSpacing = points

        best_points = 16.0
        best_pages = None
        points = 16.0
        while points >= 14.0 - 1e-6:
            set_spacing(points)
            doc.Repaginate()
            pages = doc.ComputeStatistics(2)  # wdStatisticPages
            if best_pages is None or pages < best_pages:
                best_pages, best_points = pages, points
            points -= 0.5
        set_spacing(best_points)
        doc.Repaginate()

    @staticmethod
    def _cipher_inserted_content(doc, content_start: int, content_end: int, mapping: dict) -> int:
        """Шифрує вставлені абзаци поабзацно, зберігаючи їх форматування."""
        content_range = doc.Range(content_start, content_end)
        spans = [
            (content_range.Paragraphs(i).Range.Start, content_range.Paragraphs(i).Range.End)
            for i in range(1, content_range.Paragraphs.Count + 1)
        ]

        highlighted = 0
        # Йдемо з кінця: правки не зсувають позиції попередніх абзаців.
        for start, end in reversed(spans):
            raw = doc.Range(start, end).Text or ""
            core = raw.rstrip("\r\x07")
            if not core.strip():
                continue

            ciphered, _, _ = cipher_unit_names(core, mapping)
            ciphered = apply_ukrainian_typography(clean_duplicated_units(ciphered))
            # Мʼякий перенос посеред фрази зшивається САМЕ ТУТ — у повідомленні
            # довга відкрита назва стає трьома словами, і рядок, розірваний у
            # наказі, лишався б напівпорожнім. Переноси після закінчення
            # (кома, крапка) — тобто біографічний блок — не чіпаються.
            # Витягів і примірників це не стосується: там текст 1-в-1.
            ciphered = reflow_soft_breaks(ciphered)
            if ciphered != core:
                doc.Range(start, start + len(core)).Text = ciphered

            # Відкриті назви, які лишилися без шифру, підсвічуємо жовтим.
            for span_start, span_end in unmatched_open_unit_spans(ciphered, mapping):
                doc.Range(start + span_start, start + span_end).HighlightColorIndex = 7  # wdYellow
                highlighted += 1
        return highlighted

    def _order_body_context(self, source_doc, signer_as_tag: bool = False) -> dict:
        """Межі тіла наказу + реквізити його підписанта для тегів заготовки.

        `signer_as_tag=True` — у заготовці є окремий тег `{{підписант}}`, тож
        блок підписанта підставляється в нього, а `{{зміст}}` завершується
        перед підписантом, щоб не дублювати його.
        """
        parts = self._analyze_order(source_doc)
        body_start = parts["body_start"]
        last_paragraph = parts["last_paragraph"]
        signer_start = parts["signer_start"]
        texts = parts["texts"]

        signature_line = ""
        values: dict[str, str] = {}
        signer = _find_order_signer(source_doc.Content.Text) or {}
        if signer.get("position"):
            values["{{підписант_посада}}"] = _slash_to_lines(signer["position"])
        if signer.get("rank"):
            values["{{підписант_звання}}"] = signer["rank"]
            # Табличний варіант заготовки: звання та прізвище стоять в
            # ОКРЕМИХ комірках, бо вирівняні до різних країв. Пробілами
            # цього не зробити, тому й теги окремі.
            values["{{звання_підписанта}}"] = signer["rank"]
        if signer.get("name"):
            values["{{підписант_піб}}"] = signer["name"]
            values["{{прізвище_підписанта}}"] = signer["name"]

        if signer_as_tag and signer_start and signer_start <= last_paragraph:
            # Блок підписанта беремо з наказу ЯК Є: звання та прізвище в ньому
            # вирівняні пробілами, тому текст переносимо без змін.
            signer_lines = [
                texts[index - 1] for index in range(signer_start, last_paragraph + 1)
            ]
            filled = [line for line in signer_lines if line]
            values["{{підписант}}"] = "\r".join(filled)
            # Останній рядок блоку (звання та прізвище) підкреслюється й
            # притискається праворуч — його треба знайти вже в документі.
            signature_line = filled[-1] if filled else ""

            # Зміст завершується перед підписантом.
            body_end = body_start
            for index in range(signer_start - 1, body_start - 1, -1):
                if texts[index - 1]:
                    body_end = index
                    break
            last_paragraph = body_end

        return {
            "span": (body_start, last_paragraph),
            "values": values,
            "signature_line": signature_line,
            # Чи лишився підписант усередині перенесеного змісту. Від цього
            # залежить нерозривність: інакше останній ПУНКТ вважався б
            # підписантом і втрачав зчеплення з власною шапкою.
            "signer_in_tag": bool(signer_as_tag and signer_start),
        }

    @staticmethod
    def _analyze_order(source_doc) -> dict:
        """Розбирає наказ на межі тіла, підписанта та службової частини.

        Працюємо з АБЗАЦАМИ напряму, а не через зіставлення рядків тексту:
        `Content.Text` не розбиває комірки таблиці на окремі рядки, тоді як
        `doc.Paragraphs` рахує кожну комірку окремо.
        """
        total = source_doc.Paragraphs.Count
        if total < 1:
            raise ValueError("наказ порожній")

        # chr(7) — службовий знак кінця комірки таблиці у Word. Його наявність
        # у тексті абзацу означає, що абзац лежить у таблиці, тож визначити це
        # можна без додаткових звернень до Word.
        raw_texts = [
            (source_doc.Paragraphs(index).Range.Text or "") for index in range(1, total + 1)
        ]
        texts = [raw.replace(chr(7), "").strip() for raw in raw_texts]
        in_table = [chr(7) in raw for raw in raw_texts]

        body_start = 1
        for index, clean in enumerate(texts, start=1):
            if not clean:
                continue
            upper = clean.upper()
            if (
                clean.startswith("§")
                or re.match(r"^\d+[\.\)]", clean)
                or any(keyword in upper for keyword in _ORDER_BODY_KEYWORDS)
            ):
                body_start = index
                break

        def is_service_marker(value: str) -> bool:
            clean = value.casefold()
            return bool(clean) and any(
                clean.startswith(marker) or clean == marker
                for marker in _DISTRIBUTION_CUTOFF_MARKERS
            )

        # Порядок пошуку той самий, що у витягах: спершу з кінця знаходимо
        # службову частину, потім перед нею — підписанта.
        last_marker = total + 1
        for index in range(total, body_start - 1, -1):
            if is_service_marker(texts[index - 1]):
                last_marker = index
                break

        signer_start = None
        for index in range(last_marker - 1, body_start - 1, -1):
            clean = texts[index - 1]
            if clean and _ORDER_SIGNER_START_RE.match(clean):
                signer_start = index
                break

        # Межа копіювання — те, що трапиться раніше ПІСЛЯ підписанта:
        #   • службовий блок («Розрахунок розсилки…» тощо), або
        #   • початок таблиці — це зворот останньої сторінки, його ігноруємо.
        # Таблицю відсікаємо саме за структурою, а не за текстом: у ній може
        # не бути жодного знайомого маркера.
        boundary = total + 1
        for index in range(signer_start or body_start, total + 1):
            if in_table[index - 1] or is_service_marker(texts[index - 1]):
                boundary = index
                break

        # Останній змістовний рядок перед межею — це рядок звання та прізвища
        # підписанта. Абзаци всередині таблиці сюди потрапити не можуть.
        last_paragraph = body_start
        for index in range(boundary - 1, body_start - 1, -1):
            if texts[index - 1] and not in_table[index - 1]:
                last_paragraph = index
                break
        return {
            "body_start": body_start,
            "last_paragraph": last_paragraph,
            "signer_start": signer_start,
            "texts": texts,
        }

    @staticmethod
    def _order_body_span(source_doc) -> tuple[int, int]:
        """Межі тіла наказу разом із підписантом, у номерах абзаців Word."""
        parts = App._analyze_order(source_doc)
        return parts["body_start"], parts["last_paragraph"]

    def _insert_plain_content(self, doc, tag_range, encrypted_content: str, mapping=None) -> int:
        """Запасний спосіб вставки змісту — простим текстом.

        Використовується ЛИШЕ тоді, коли перенести форматування з наказу не
        вдалося. Форматування тут доводиться вгадувати за вмістом рядка, тому
        воно менш точне, ніж копіювання з оригіналу.
        """
        if not encrypted_content:
            return 0

        content_start = tag_range.Start
        cleaned_content = clean_duplicated_units(encrypted_content)
        separated_content = ensure_blank_line_before_items(cleaned_content)
        formatted_content = apply_ukrainian_typography(separated_content)
        tag_range.Text = formatted_content
        content_range = doc.Range(content_start, content_start + len(formatted_content))

        for pi in range(1, content_range.Paragraphs.Count + 1):
            paragraph = content_range.Paragraphs(pi)
            p_text = paragraph.Range.Text.strip()
            if not p_text:
                continue
            p_format = paragraph.Range.ParagraphFormat
            if is_biographical_paragraph(p_text):
                p_format.Alignment = 1  # wdAlignParagraphCenter
                p_format.LeftIndent = 0
                p_format.RightIndent = 0
                p_format.FirstLineIndent = 0
                p_format.SpaceBefore = 0
                p_format.SpaceAfter = 0
                p_format.LineSpacingRule = 0  # wdLineSpaceSingle
            elif re.match(r"^\d{1,3}[\.\)]", p_text):
                p_format.Alignment = 3  # wdAlignParagraphJustify
                p_format.LeftIndent = 0
                p_format.RightIndent = 0
                p_format.FirstLineIndent = 35.45  # 1.25 см
                p_format.SpaceBefore = 6
                p_format.SpaceAfter = 0
            elif (
                p_text.startswith("Відповідно до")
                or "ЗВІЛЬНИТИ" in p_text.upper()
                or "ПРИЗНАЧИТИ:" in p_text.upper()
            ):
                p_format.Alignment = 3
                p_format.LeftIndent = 0
                p_format.RightIndent = 0
                p_format.FirstLineIndent = 35.45
                p_format.SpaceBefore = 6
                p_format.SpaceAfter = 6
            elif "Призначається на" in p_text or "шпк" in p_text.lower():
                p_format.Alignment = 3
                p_format.LeftIndent = 0
                p_format.RightIndent = 0
                p_format.FirstLineIndent = 35.45
                p_format.SpaceBefore = 0
                p_format.SpaceAfter = 6

        highlighted = 0
        for span_start, span_end in unmatched_open_unit_spans(formatted_content, mapping):
            doc.Range(content_start + span_start, content_start + span_end).HighlightColorIndex = 7
            highlighted += 1
        return highlighted

    def _create_message_file(
        self,
        word,
        template_path: str,
        output_path: str,
        replacements: dict[str, str],
        encrypted_content: str = "",
        recipients: list[str] | None = None,
        content_source: dict | None = None,
    ) -> tuple[int, int, int]:
        # Якщо результат уже відкритий у Word, зберегти його неможливо.
        # Перевіряємо це ДО всієї роботи й повідомляємо зрозуміло.
        if not is_path_writable(output_path):
            raise UserError(
                f"файл «{os.path.basename(output_path)}» уже відкритий — "
                "найімовірніше у Word, з минулого разу.",
                "Закрийте це вікно Word і натисніть кнопку ще раз. "
                "Поки файл відкритий, програма не може його перезаписати.",
            )

        # Робоча копія шаблону створюється у тимчасовій папці, а не за
        # кінцевим шляхом: так оригінал шаблону лишається недоторканим, а
        # заблокований результат не зриває обробку. Розширення копії
        # відповідає реальному вмісту шаблону (він може бути у форматі
        # Word 97-2003), інакше Word відмовиться її відкрити.
        temp_dir = os.path.join(os.path.dirname(os.path.abspath(output_path)), "_nat_temp")
        os.makedirs(temp_dir, exist_ok=True)
        working_path = copy_template_for_editing(
            template_path,
            os.path.join(temp_dir, os.path.basename(output_path)),
            label="шаблон повідомлення",
        )
        doc = word.Documents.Open(os.path.abspath(working_path), ReadOnly=False)
        try:
            executor_blocks = self._replace_message_tags(doc, replacements)
            if "{{виконавець}}" in replacements:
                header_hits = self._replace_executor_in_headers(
                    doc, replacements["{{виконавець}}"]
                )
                if header_hits:
                    self.log(
                        f"  Виконавця підставлено в колонтитул: {header_hits} раз(и) "
                        f"({os.path.basename(output_path)})."
                    )
            executor_bookmarks = []
            for index, (start, end) in enumerate(executor_blocks):
                bookmark_name = f"nat_message_executor_{index}"
                doc.Bookmarks.Add(bookmark_name, doc.Range(start, end))
                executor_bookmarks.append(bookmark_name)
            recipient_slots = recipient_overflow = 0
            if recipients is not None:
                recipient_slots, recipient_overflow = self._fill_recipient_slots(doc, recipients)
            highlighted_count = 0
            if encrypted_content or content_source:
                find_obj = doc.Content.Find
                find_obj.Text = "{{зміст_шифр}}"
                if not find_obj.Execute():
                    self.log("УВАГА: у шаблоні зі змістом не знайдено тег {{зміст_шифр}}.")
                else:
                    # Основний шлях — перенос змісту РАЗОМ ІЗ ФОРМАТУВАННЯМ
                    # наказу, точно як у витягах. Запасний шлях (простий текст)
                    # лишається на випадок, коли діапазон абзаців визначити не вдалося.
                    copied_with_formatting = False
                    if content_source is not None:
                        try:
                            highlighted_count = self._copy_order_content(
                                doc, find_obj.Parent, content_source
                            )
                            copied_with_formatting = True
                        except ValueError as prepare_error:
                            self.log(
                                "УВАГА: зміст не вдалося перенести з форматуванням "
                                f"({prepare_error}); вставляємо простим текстом."
                            )

                    if not copied_with_formatting:
                        highlighted_count = self._insert_plain_content(
                            doc,
                            find_obj.Parent,
                            encrypted_content,
                            (content_source or {}).get("mapping"),
                        )
            for bookmark_name in executor_bookmarks:
                self._position_message_executor_at_page_bottom(doc, bookmark_name)
                if doc.Bookmarks.Exists(bookmark_name):
                    doc.Bookmarks(bookmark_name).Delete()
            doc.SaveAs2(os.path.abspath(output_path), 16)
            return highlighted_count, recipient_slots, recipient_overflow
        finally:
            doc.Close(False)
            shutil.rmtree(temp_dir, ignore_errors=True)

    # Дописувати нові частини в таблицю заготовками (розд. 9.5.7). Увімкнено
    # лише в Qt-оболонці (`QtShellMixin`); Tk-версія тільки попереджає.
    ADD_NEW_UNITS_TO_TABLE = False

    def _report_table_gaps(self, order_text: str, mapping: dict) -> dict:
        """Показує, чого бракує в таблиці для закритого змісту повідомлення.

        Збій запису заготовок генерацію не зупиняє — лише попереджає.
        """
        body = "\n".join(order_text.splitlines()[find_content_start_line(order_text):])
        problems: list = []
        cipher_unit_names(body, mapping, problems=problems)
        for line in describe_cipher_problems(problems):
            self.log(line)

        similar: list = []
        new_units = collect_new_unit_names(body, mapping, similar=similar)
        for name, row in similar:
            self.log(
                f"УВАГА: «{name}» не зашифровано: у таблиці є схожий рядок «{row}», але назва "
                "в наказі збігається з ним не повністю — перевірте написання в наказі або в стовпці A."
            )
        if new_units:
            self.log(
                f"УВАГА: частин немає в таблиці ({len(new_units)}): {'; '.join(new_units)}. "
                "У повідомленні вони лишаються відкритими й виділені жовтим."
            )
            if self.ADD_NEW_UNITS_TO_TABLE:
                try:
                    stubs = append_unit_stubs_to_table(self.excel_path.get(), new_units)
                except Exception as error:
                    self.log(f"УВАГА: не вдалося дописати нові частини в таблицю: {error}")
                else:
                    if stubs["added"] and stubs["separate"]:
                        self.log(
                            f"УВАГА: дописати в таблицю не вдалося (потрібен Excel), тому нові частини "
                            f"({len(stubs['added'])}) записано окремо: {stubs['path']}. "
                            "Перенесіть рядки в таблицю й заповніть шифри."
                        )
                    elif stubs["added"]:
                        self.log(
                            f"Дописано в таблицю заготовок: {len(stubs['added'])} (виділені жовтим, "
                            f"шифр порожній). Резервна копія: {stubs['backup']}. "
                            "Заповніть шифр і за потреби «Кому»/«Куди», потім створіть повідомлення ще раз."
                        )
                    elif stubs["separate"]:
                        # Таблиця з формулами: заготовки лежать в окремому файлі,
                        # а не в таблиці — інакше користувач шукав би їх не там.
                        self.log(
                            f"Заготовки для цих частин уже є в окремому файлі: {stubs['path']}. "
                            "Перенесіть рядки в таблицю й заповніть шифри."
                        )
                    else:
                        self.log("Заготовки для цих частин уже є в таблиці — заповніть у них шифри.")
        return {"problems": len(problems) + len(similar), "new_units": len(new_units)}

    def run_generate_messages(self):
        self.save_config()
        required = (
            self.excel_path.get(),
            self.doc_path.get(),
            self.message_cover_template_path.get(),
            self.message_content_template_path.get(),
        )
        if not all(required):
            messagebox.showwarning(
                "Помилка",
                "Виберіть словник Excel, наказ і два DOCX-шаблони повідомлень.",
            )
            return

        order_path = os.path.abspath(self.doc_path.get())
        if not os.path.isfile(order_path):
            messagebox.showwarning("Помилка", "Файл наказу не знайдено.")
            return

        out_folder = self.message_out_folder.get() or os.path.join(os.path.dirname(order_path), "Messages_Output")
        try:
            os.makedirs(out_folder, exist_ok=True)
        except OSError as error:
            messagebox.showwarning("Помилка папки", f"Не вдалося створити папку результату:\n{error}")
            return
        self.message_out_folder.set(out_folder)

        self.btn_generate_messages.config(state=DISABLED)
        word = None
        source_doc = None
        try:
            mapping = read_recipient_mapping(path=self.excel_path.get()).get("mapping", {})

            # Наказ відкриваємо в тому самому екземплярі Word, з якого потім
            # копіюємо зміст із форматуванням, — інакше номери абзаців не
            # відповідали б прочитаному тексту.
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0  # wdAlertsNone: не показувати блокуючі діалоги
            source_doc = word.Documents.Open(order_path, ReadOnly=True)

            source_text = source_doc.Content.Text
            order_text, order_signer = text_before_order_signer(source_text)
            order_num, order_date = extract_metadata_from_filename(os.path.basename(order_path))
            # У повідомленнях дата не розкривається словами (на відміну від витягів).
            order_date_formatted = format_message_date(order_date)

            table_gaps = self._report_table_gaps(order_text, mapping)
            routes = map_military_units(text=order_text, mapping=mapping)
            recipient_groups = build_message_recipient_groups(mapping, routes)
            recipients = (
                recipient_groups["corps"] + recipient_groups["units"] + recipient_groups["tck"]
            )
            addressee_kind = build_addressee_kind_text(recipient_groups)
            destinations = []
            for data in routes.get("unit_paragraphs", {}).values():
                destination = str(data.get("destination_where") or "").strip()
                if destination and destination not in destinations:
                    destinations.append(destination)

            replacements = {}
            if order_num:
                replacements["{{номер_наказу}}"] = f"№{order_num}"
            if order_date_formatted:
                replacements["{{дата_наказу}}"] = order_date_formatted
            if destinations:
                replacements["{{куди}}"] = "\r".join(destinations)
            # Тег типу адресата підставляється ДО {{виконавець}}: обробка
            # виконавця видаляє весь службовий хвіст шаблону після себе.
            if addressee_kind:
                for kind_tag in _MESSAGE_ADDRESSEE_KIND_TAGS:
                    replacements[kind_tag] = addressee_kind
                self.log(f"Тип адресата ({{{{тцк чі вч}}}}): {addressee_kind}.")
            # Підписант наказу — ті самі теги, що у витягах. Реквізити беруться
            # з наказу, а якщо користувач заповнив поля вручну — з полів.
            gui_signer_position = self.order_signer_position.get().strip()
            gui_signer_rank = self.order_signer_rank.get().strip()
            gui_signer_name = self.order_signer_name.get().strip()
            if gui_signer_position or gui_signer_rank or gui_signer_name:
                order_signer = {
                    "position": gui_signer_position.replace(" / ", chr(10)) or order_signer.get("position", ""),
                    "rank": gui_signer_rank or order_signer.get("rank", ""),
                    "name": gui_signer_name or order_signer.get("name", ""),
                }
            signer_position = _slash_to_lines(str(order_signer.get("position", "")).strip())
            signer_rank = str(order_signer.get("rank", "")).strip()
            signer_name = str(order_signer.get("name", "")).strip()
            replacements.update(signer_tags(signer_position, signer_rank, signer_name))

            # Засвідчувач — СПІЛЬНІ поля з витягами (AGENT.md 11.2).
            replacements["{{згідно_з_оригіналом}}"] = "Згідно з оригіналом"
            replacements["{{засвідчення}}"] = "Згідно з оригіналом"
            if self.certifier_position.get().strip():
                certifier_position = _slash_to_lines(self.certifier_position.get().strip())
                replacements["{{засвідчувач_посада}}"] = certifier_position
                replacements["{{згідно_з_оригіналом_посада}}"] = certifier_position
            if self.certifier_rank.get().strip():
                certifier_rank = self.certifier_rank.get().strip()
                replacements["{{засвідчувач_звання}}"] = certifier_rank
                replacements["{{згідно_з_оригіналом_звання}}"] = certifier_rank
            if self.certifier_name.get().strip():
                certifier_name = self.certifier_name.get().strip()
                replacements["{{засвідчувач_піб}}"] = certifier_name
                replacements["{{згідно_з_оригіналом_піб}}"] = certifier_name
                replacements["{{засвідчувач}}"] = certifier_name

            if self.message_executor.get().strip():
                replacements["{{виконавець}}"] = self.message_executor.get().strip()

            decision = generate_decision_order(text=order_text, mapping=mapping, new_header="")
            encrypted_content = decision.get("decision_text", "")

            # Відповідність рядків тексту абзацам Word — так само, як у витягах.
            content_source = None
            try:
                source_line_to_para = []
                for paragraph_index in range(1, source_doc.Paragraphs.Count + 1):
                    raw = source_doc.Paragraphs(paragraph_index).Range.Text
                    logical_lines = raw.rstrip("\r\x07").splitlines() or [""]
                    source_line_to_para.extend([paragraph_index] * len(logical_lines))

                order_lines = order_text.splitlines()
                line_map = source_line_to_para[: len(order_lines)]
                body_start_line = find_content_start_line(order_text)
                if line_map and 0 <= body_start_line < len(line_map):
                    content_source = {
                        "doc": source_doc,
                        "first_para": line_map[body_start_line],
                        "last_para": line_map[-1],
                        "mapping": mapping,
                    }
            except Exception as map_error:
                self.log(
                    f"УВАГА: не вдалося зіставити абзаци наказу ({map_error}); "
                    "зміст буде вставлено простим текстом."
                )

            safe_number = re.sub(r'[\\/:*?"<>|]', "_", order_num) if order_num else ""
            suffix = f"_№{safe_number}" if safe_number else ""
            cover_output = os.path.join(out_folder, f"Повідомлення_супровід{suffix}.docx")
            content_output = os.path.join(out_folder, f"Повідомлення_шифрований_зміст{suffix}.docx")

            _, recipient_slots, recipient_overflow = self._create_message_file(
                word,
                self.message_cover_template_path.get(),
                cover_output,
                replacements,
                recipients=recipients,
            )
            highlights, _, _ = self._create_message_file(
                word,
                self.message_content_template_path.get(),
                content_output,
                replacements,
                encrypted_content,
                content_source=content_source,
            )
            if recipient_overflow:
                self.log(
                    f"УВАГА: у шаблоні супроводу {recipient_slots} комірок {{кому_список}}, "
                    f"але адресатів {len(recipients)}. Не вмістилося: {recipient_overflow}."
                )
            self.save_config()
            self.log(
                f"Повідомлення створено: {cover_output}; {content_output}. "
                f"Адресатів: {len(recipients)}; жовтих позначок: {highlights}."
            )
            messagebox.showinfo(
                "Успіх",
                "Створено 2 повідомлення.\n\n"
                f"Заповнено комірок {{кому_список}}: {min(recipient_slots, len(recipients))} з {len(recipients)}\n"
                f"Невпізнаних назв, виділених жовтим: {highlights}\n"
                f"Частин, яких немає в таблиці: {table_gaps['new_units']}\n"
                f"Без шифру або корпусу в таблиці: {table_gaps['problems']}"
                + (f"\nУВАГА: не вмістилося адресатів: {recipient_overflow}" if recipient_overflow else ""),
            )
        except Exception as error:
            explanation = explain_error(error)
            self.log(f"ПОМИЛКА повідомлень:\n{explanation}")
            messagebox.showerror("Повідомлення не створені", explanation)
        finally:
            if source_doc is not None:
                try:
                    source_doc.Close(False)
                except Exception:
                    pass
            if word:
                force_quit_word(word)
            self.btn_generate_messages.config(state=NORMAL)

    def select_p2_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.p2_orders_folder.set(path)
            self.p2_source_mode.set("folder")
            # Папку результату переобчислюємо під нове джерело, доки
            # користувач не вибрав її вручну.
            if not self.p2_out_folder_manual.get():
                self.p2_out_folder.set(os.path.join(path, "Примірники_2"))
            self._refresh_p2_orders_view()
            self.save_config()

    def select_p2_file(self):
        paths = filedialog.askopenfilenames(filetypes=[("Word files", "*.docx *.doc")])
        if paths:
            self._set_p2_orders(list(paths))
            self.log_p2(f"Обрано наказів для примірників: {len(self.p2_manual_order_paths)}")

    def select_p2_back_page(self):
        path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx *.doc")])
        if path:
            self.p2_back_page_path.set(path)
            self.save_config()

    def edit_p2_back_page(self):
        path = self.p2_back_page_path.get()
        if not path or not os.path.exists(path):
            messagebox.showwarning("Помилка", "Спочатку виберіть файл шаблону задньої сторінки!")
            return
        try:
            self.log_p2("Відкриваємо задню сторінку у Word для редагування...")
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = True
            doc = word.Documents.Open(os.path.abspath(path))
            doc.Activate()
        except Exception as e:
            self.log_p2(f"Помилка відкриття задньої сторінки: {str(e)}")

    def select_p2_out_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.p2_out_folder.set(path)
            # Явний вибір користувача — більше не переобчислюємо автоматично.
            self.p2_out_folder_manual.set(True)
            self.save_config()

    def _read_word_text(self, doc_path: str) -> str:
        """Текст наказу через Word, з кешем за файлом і часом його зміни.

        Запуск окремого екземпляра Word коштує близько секунди, а той самий наказ
        читається кілька разів за сеанс: розрахунок розсилки, оновлення підписанта,
        потім генерація витягів. Поки файл не змінився, вдруге Word не запускаємо.
        """
        cache = getattr(self, "_word_text_cache", None)
        if cache is None:
            cache = self._word_text_cache = {}
        try:
            stat = os.stat(doc_path)
            cache_key = (os.path.abspath(doc_path), stat.st_mtime_ns, stat.st_size)
        except OSError:
            cache_key = None
        if cache_key is not None and cache_key in cache:
            return cache[cache_key]

        word = None
        doc = None
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0  # wdAlertsNone: не показувати блокуючі діалоги (напр. "Зберегти зміни?")
            doc = word.Documents.Open(os.path.abspath(doc_path), ReadOnly=True)
            text = read_document_text(doc)
            if cache_key is not None:
                # Кеш свідомо маленький: тексти наказів важкі, а користь дає
                # лише останній оброблюваний файл (і сусідні в пакеті).
                if len(cache) >= 4:
                    cache.clear()
                cache[cache_key] = text
            return text
        finally:
            try:
                if doc:
                    doc.Close(False)
            except Exception:
                pass
            if word:
                force_quit_word(word)

    def _refresh_order_signer(self, text: str | None = None) -> tuple[str, dict[str, str]]:
        if text is None:
            try:
                text = self._read_word_text(self.doc_path.get())
            except Exception as error:
                self.log(f"УВАГА: не вдалося зчитати підписанта наказу: {error}")
                text = ""
        usable_text, signer = text_before_order_signer(text)
        self.order_signer_position.set(signer["position"].replace("\n", " / "))
        self.order_signer_rank.set(signer["rank"])
        self.order_signer_name.set(signer["name"])
        if signer["position"] or signer["rank"] or signer["name"]:
            self.log("Підписанта наказу зчитано; текст після нього буде проігноровано.")
        else:
            self.log("УВАГА: підписанта наказу не розпізнано; текст наказу не обрізано.")
        return usable_text, signer

    def _set_copy_two_sources(self, paths: list[str], selected: bool = True):
        """Показує примірники № 2 останнього проходу як позначуваний список."""
        self.last_copy_two_paths = [os.path.abspath(path) for path in paths if os.path.isfile(path)]
        for item in self.copy_two_tree.get_children():
            self.copy_two_tree.delete(item)
        if not self.last_copy_two_paths:
            self.copy_two_tree.insert(
                "", tk.END,
                values=("—", "Спершу сформуйте примірники № 2 у вкладці «Примірники 2/3»."),
            )
            self._refresh_source_summary()
            return
        mark = "☑" if selected else "☐"
        for index, path in enumerate(self.last_copy_two_paths):
            self.copy_two_tree.insert("", tk.END, iid=f"copy2_{index}", values=(mark, path))
        self._select_order_path(self.copy_two_tree, self.doc_path.get())
        self._refresh_source_summary()

    @staticmethod
    def _marked_paths(tree) -> list[str]:
        """Позначені галочкою й досі наявні файли зі списку джерел."""
        return [
            values[1]
            for item in tree.get_children()
            if (values := tree.item(item, "values"))
            and len(values) == 2 and values[0] == "☑" and os.path.isfile(values[1])
        ]

    @staticmethod
    def _orders_in_folder(folder: str) -> list[str]:
        """Накази DOCX у теці, без службових і вже згенерованих файлів."""
        if not folder or not os.path.isdir(folder):
            return []
        return [
            os.path.join(folder, name)
            for name in sorted(os.listdir(folder))
            if name.lower().endswith(".docx")
            and not name.startswith("~$")
            and not is_generated_copy_filename(name)
        ]

    def _fill_orders_tree(self, tree, paths: list[str], empty_hint: str):
        for item in tree.get_children():
            tree.delete(item)
        if not paths:
            tree.insert("", tk.END, values=("—", empty_hint))
            return
        for index, path in enumerate(paths):
            tree.insert("", tk.END, iid=f"order_{index}", values=("☑", path))

    @staticmethod
    def _select_order_path(tree, preferred_path: str = ""):
        """Позначає активний рядок окремо від галочки пакетної обробки."""
        children = tree.get_children()
        if not children:
            return
        preferred = os.path.normcase(os.path.abspath(preferred_path)) if preferred_path else ""
        chosen = ""
        for item in children:
            values = tree.item(item, "values")
            if len(values) < 2 or not os.path.isfile(values[1]):
                continue
            if not chosen:
                chosen = item
            if preferred and os.path.normcase(os.path.abspath(values[1])) == preferred:
                chosen = item
                break
        if chosen:
            tree.selection_set(chosen)
            tree.see(chosen)

    def _activate_order_from_tree(self, tree, alongside: bool = False):
        """Перемикає активний наказ і показує його збережені контрольні XLSX."""
        if self._batch_running or getattr(self, "_busy_depth", 0):
            return
        selected = tree.selection()
        if not selected:
            return
        values = tree.item(selected[0], "values")
        if len(values) < 2 or not os.path.isfile(values[1]):
            return
        order_path = os.path.abspath(values[1])
        output_folder = (
            os.path.dirname(order_path)
            if alongside
            else os.path.join(os.path.dirname(order_path), "Extracts_Output")
        )
        self.doc_path.set(order_path)
        self.out_folder.set(output_folder)
        self._refresh_source_summary()
        self._load_saved_analysis_into_ui(order_path, output_folder)

    def _load_saved_analysis_into_ui(self, order_path: str, output_folder: str):
        """Заповнює три вкладки з диска; Word і маршрутизатор не запускаються."""
        if not hasattr(self, "result_views"):
            return
        saved = load_saved_analysis_reports(order_path, output_folder)
        placeholders = {
            "calculation": ("— Для цього наказу ще немає збереженого розрахунку —", "—", "—"),
            "unmatched": ("—", "— Збережений контроль пропущених ще не створено —", "—"),
            "routing": ("—", "—", "— Збережений контроль маршрутизації ще не створено —", "—"),
        }
        columns = {
            "calculation": ["Військова частина / Відправник", "Пункти витягу", "Кількість"],
            "unmatched": ["Пункт", "Текст пункту", "Причина"],
            "routing": ["Пункт", "Збіги з таблиці", "Застосовані правила", "Підсумкові адресати"],
        }
        for key in ("calculation", "unmatched", "routing"):
            rows = saved[key]
            if not rows:
                if saved["found"][key] and key == "unmatched":
                    rows = [("—", "Пропущених пунктів немає.", "—")]
                else:
                    rows = [placeholders[key]]
            self._populate_result_tab(key, columns[key], rows)

        found_count = sum(saved["found"].values())
        name = os.path.basename(order_path)
        if found_count:
            self.log(
                f"Підвантажено збережені результати для {name}: "
                f"{found_count}/3 контрольних файлів; повторний розрахунок не запускався."
            )
        else:
            self.log(f"Для {name} збережених результатів ще немає.")
        for error in saved["errors"]:
            self.log(f"УВАГА: не вдалося відкрити збережений результат: {error}")
        self.results_notebook.select(self.result_tabs["calculation"])

    def _refresh_orders_view(self):
        """Оновлює список наказів вкладки витягів і підсумковий рядок."""
        tree = getattr(self, "orders_tree", None)
        if tree is None:
            return
        self._fill_orders_tree(
            tree, self.manual_order_paths,
            "Накази не обрано — натисніть «📂 Обрати накази» або перетягніть файли сюди.",
        )
        self._select_order_path(tree, self.doc_path.get())
        self._refresh_source_summary()

    def _refresh_source_summary(self):
        """Один рядок, який чесно каже, що саме піде в обробку."""
        if not hasattr(self, "source_summary"):
            return
        paths, from_copies = self._selected_order_paths()
        if not paths:
            self.source_summary.set("Джерело не обрано — оберіть наказ або кілька наказів")
            self.message_source_summary.set("Наказ для повідомлень не обрано")
            return
        names = ", ".join(os.path.basename(path) for path in paths[:3])
        if len(paths) > 3:
            names += f" … ще {len(paths) - 3}"
        prefix = "Примірників № 2 в обробці" if from_copies else "Наказів в обробці"
        self.source_summary.set(f"{prefix}: {len(paths)} — {names}")
        active = self.doc_path.get()
        self.message_source_summary.set(
            "Один наказ: " + os.path.basename(active) if active and os.path.isfile(active)
            else "Наказ для повідомлень не обрано"
        )

    def _set_orders(self, paths: list[str]):
        """Приймає накази для вкладки витягів і готує папку результату."""
        self.manual_order_paths = [
            os.path.abspath(path) for path in paths if os.path.isfile(path)
        ]
        if self.manual_order_paths:
            # Поточним лишається ПЕРШИЙ наказ: його реквізити показуються в
            # інтерфейсі, з нього ж працюють повідомлення й порівняння.
            self.doc_path.set(self.manual_order_paths[0])
            self.out_folder.set(
                os.path.join(os.path.dirname(self.manual_order_paths[0]), "Extracts_Output")
            )
            self._refresh_order_signer()
        self._refresh_orders_view()
        self.save_config()

    def select_orders(self):
        """Один або кілька наказів — головний (і єдиний) вибір у вікні витягів."""
        paths = filedialog.askopenfilenames(filetypes=[("Word files", "*.docx *.doc")])
        if not paths:
            return
        self._set_orders(list(paths))
        self.log(f"Обрано наказів: {len(self.manual_order_paths)}")

    def select_message_order(self):
        """Повідомлення навмисно одинарні: один наказ — два вихідні файли."""
        path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx *.doc")])
        if path:
            self._set_orders([path])
            self.log(f"Обрано наказ для повідомлень: {os.path.basename(path)}")

    def select_orders_folder(self):
        """Уся тека наказів одним пакетом."""
        folder = filedialog.askdirectory()
        if not folder:
            return
        orders = self._orders_in_folder(folder)
        if not orders:
            messagebox.showwarning("Немає файлів", "У цій теці немає наказів DOCX для обробки.")
            return
        self._set_orders(orders)
        self.log(f"З теки {folder} обрано наказів: {len(orders)}")

    def clear_orders(self):
        self.manual_order_paths = []
        self._refresh_orders_view()
        self.save_config()

    def _refresh_p2_orders_view(self):
        """Список наказів для примірників: окремі файли або вміст теки."""
        tree = getattr(self, "p2_orders_tree", None)
        if tree is None:
            return
        if self.p2_source_mode.get() == "folder":
            paths = self._orders_in_folder(self.p2_orders_folder.get())
        else:
            paths = self._p2_manual_paths()
        self._fill_orders_tree(
            tree, paths,
            "Накази не обрано — натисніть «📂 Обрати накази» або перетягніть файли сюди.",
        )
        self._refresh_p2_source_summary()

    def _refresh_p2_source_summary(self):
        if not hasattr(self, "p2_source_summary"):
            return
        paths = self._selected_p2_order_paths()
        if not paths:
            self.p2_source_summary.set("Накази не обрано")
            return
        source = (
            f"тека {os.path.basename(self.p2_orders_folder.get().rstrip(os.sep))}"
            if self.p2_source_mode.get() == "folder"
            else "обрані файли"
        )
        self.p2_source_summary.set(f"Наказів для примірників: {len(paths)} ({source})")

    def _p2_manual_paths(self) -> list[str]:
        """Окремо обрані накази для примірників, що досі існують на диску."""
        manual = [path for path in self.p2_manual_order_paths if os.path.isfile(path)]
        single = self.p2_single_file.get()
        if not manual and single and os.path.isfile(single):
            # Старий «окремий файл» із конфігу, записаного до появи списку.
            manual = [os.path.abspath(single)]
        return manual

    def _selected_p2_order_paths(self) -> list[str]:
        """Накази, позначені галочкою на вкладці примірників."""
        tree = getattr(self, "p2_orders_tree", None)
        if tree is not None:
            marked = self._marked_paths(tree)
            if marked:
                return marked
        if self.p2_source_mode.get() == "folder":
            return self._orders_in_folder(self.p2_orders_folder.get())
        return self._p2_manual_paths()

    def _set_p2_orders(self, paths: list[str]):
        self.p2_manual_order_paths = [
            os.path.abspath(path) for path in paths if os.path.isfile(path)
        ]
        self.p2_source_mode.set("file")
        if self.p2_manual_order_paths:
            self.p2_single_file.set(self.p2_manual_order_paths[0])
            if not self.p2_out_folder_manual.get():
                self.p2_out_folder.set(
                    os.path.join(os.path.dirname(self.p2_manual_order_paths[0]), "Примірники_2")
                )
        self._refresh_p2_orders_view()
        self.save_config()

    def clear_p2_orders(self):
        self.p2_manual_order_paths = []
        self.p2_single_file.set("")
        self._refresh_p2_orders_view()
        self.save_config()

    def copy_certifier_from_extracts(self):
        """Переносить засвідчувача з витягів у примірники — коли він той самий."""
        self.p2_certifier_position.set(self.certifier_position.get())
        self.p2_certifier_rank.set(self.certifier_rank.get())
        self.p2_certifier_name.set(self.certifier_name.get())
        self.save_config()
        self.log_p2("Засвідчувача примірників скопійовано з витягів.")

    def _selected_order_paths(self) -> tuple[list[str], bool]:
        """Повертає `(шляхи, чи це примірники № 2)`.

        Прапорець потрібен для вибору папки результату: для примірників,
        створених пакетом, витяги кладемо поряд із ними у папку наказу.
        """
        if self.last_copy_two_paths:
            paths = self._marked_paths(self.copy_two_tree)
            if paths:
                return paths, True
            # Галочки зняті з усіх примірників — працюємо з наказами,
            # обраними вручну, замість того щоб відмовляти користувачу.
        tree = getattr(self, "orders_tree", None)
        if tree is not None:
            marked = self._marked_paths(tree)
            if marked:
                return marked, False
        manual_orders = [path for path in self.manual_order_paths if os.path.isfile(path)]
        if manual_orders:
            return manual_orders, False
        manual_order = self.doc_path.get()
        if manual_order and os.path.isfile(manual_order):
            return [manual_order], False
        return [], False

    def _set_processing_order(self, doc_path: str, alongside: bool = False):
        """Готує наказ до обробки й визначає папку результату.

        `alongside=True` — це примірник № 2 із пакетної обробки: витяги та
        звіти розрахунку кладемо ПОРЯД із ним, у ту саму папку наказу.
        Для вручну обраного наказу лишається окрема `Extracts_Output`,
        щоб не засмічувати папку, де лежать самі накази.
        """
        self.doc_path.set(doc_path)
        order_dir = os.path.dirname(os.path.abspath(doc_path))
        self.out_folder.set(order_dir if alongside else os.path.join(order_dir, "Extracts_Output"))
        os.makedirs(self.out_folder.get(), exist_ok=True)

    # =========================================================================
    # ДІЇ: РОЗРАХУНОК ТА ВИТЯГИ
    # =========================================================================
    def _missing_selected_files(self, *entries: tuple[str, str]) -> list[str]:
        """Перелік «підпис: шлях» для обраних файлів, яких немає на диску.

        Шляхи зберігаються в config.json, тож між запусками файл можуть
        перейменувати, перемістити або лишити на відключеному диску. Без цієї
        перевірки генерація доходить аж до копіювання зразка й падає з
        «[WinError 2] The system cannot find the file specified» — уже після
        розбору наказу й запису контрольних таблиць.
        """
        return [
            f"{label}: {path}"
            for label, path in entries
            if path and not os.path.isfile(path)
        ]

    def _report_missing_files(self, missing: list[str]) -> None:
        """Називає конкретні відсутні файли — і в журналі, і у вікні."""
        details = "\n".join(f"• {item}" for item in missing)
        self.log("ПОМИЛКА: не знайдено обраних файлів:\n" + details)
        messagebox.showwarning(
            "Немає потрібного файлу",
            "Що сталося: файл(и), вибрані для роботи, не лежать там, де вказано:\n\n"
            f"{details}\n\n"
            "Найчастіше файл перейменували, перемістили або він на диску, "
            "який зараз не підключений.\n\n"
            "Що зробити:\n"
            "1. Відкрийте «Зразки та реквізити».\n"
            "2. Натисніть «Вибрати» біля рядка, названого вище.\n"
            "3. Вкажіть потрібний файл і збережіть вікно.\n"
            "4. Натисніть кнопку генерації ще раз.",
        )

    def _set_extract_action_buttons_state(self, state):
        """Разом блокує дії вкладки, щоб два пакети не стартували паралельно."""
        for name in ("btn_calc", "btn_extracts", "btn_management_extracts"):
            button = getattr(self, name, None)
            if button is not None:
                button.config(state=state)

    def run_rozrahunok_action(self):
        self.save_config()
        order_paths, from_copies = self._selected_order_paths()
        if not self.excel_path.get() or not order_paths:
            messagebox.showwarning(
                "Помилка",
                "Виберіть словник Excel і хоча б один наказ або примірник № 2.",
            )
            return

        missing = self._missing_selected_files(("Словник Excel", self.excel_path.get()))
        if missing:
            self._report_missing_files(missing)
            return

        self._set_extract_action_buttons_state(DISABLED)
        try:
            # Кожен наказ обробляється окремо: збій одного не має зривати
            # решту пакета.
            failures = self._run_batch(
                order_paths, self.run_rozrahunok, "Розрахунок для", alongside=from_copies
            )
            self.save_config()
            self._report_batch_result(
                len(order_paths), failures, "Розрахунок розсилки", "розрахунок розсилки"
            )
        finally:
            self._set_extract_action_buttons_state(NORMAL)

    def _run_extracts_scope_action(
        self,
        scope: str,
        stage_label: str,
        result_title: str,
        result_action: str,
    ):
        self.save_config()
        order_paths, from_copies = self._selected_order_paths()
        if not self.excel_path.get() or not self.template_path.get() or not order_paths:
            messagebox.showwarning(
                "Помилка",
                "Виберіть словник Excel, шаблон витягу і хоча б один наказ або примірник № 2.",
            )
            return

        missing = self._missing_selected_files(
            ("Словник Excel", self.excel_path.get()),
            ("Зразок витягу", self.template_path.get()),
        )
        if missing:
            self._report_missing_files(missing)
            return

        self._set_extract_action_buttons_state(DISABLED)
        try:
            failures = self._run_batch(
                order_paths,
                lambda: self.run_extracts(scope=scope),
                stage_label,
                alongside=from_copies,
            )
            self.save_config()
            self._report_batch_result(
                len(order_paths), failures, result_title, result_action
            )
        finally:
            self._set_extract_action_buttons_state(NORMAL)

    def run_extracts_action(self, scope: str = "all"):
        """Створює адресний файл і окремий файл витягів до управління."""
        return self._run_extracts_scope_action(
            "all", "Усі витяги для", "Генерація всіх витягів", "генерацію всіх витягів"
        )

    def run_management_extracts_action(self):
        """Створює лише окремі витяги до управління без адресування й 2-на-1."""
        return self._run_extracts_scope_action(
            "management",
            "Витяги до управління для",
            "Витяги до управління",
            "генерацію витягів до управління",
        )

    def run_full_cycle(self):
        """Повний прохід по обраних наказах: примірники № 2 → розрахунок → витяги.

        Джерелом є список наказів головного вікна. Примірники формуються
        першими, і далі розрахунок та витяги йдуть уже по них — саме так, як
        це робиться вручну послідовними діями вкладки.
        """
        if self._batch_running:
            self.log("Пакет уже виконується — дочекайтеся завершення.")
            return

        order_paths, from_copies = self._selected_order_paths()
        if from_copies:
            # Примірники вже є: другий раз їх робити нема з чого, тож повний
            # цикл починався б із самого себе. Беремо вихідні накази.
            order_paths = [path for path in self.manual_order_paths if os.path.isfile(path)]
        if not order_paths:
            messagebox.showwarning("Помилка", "Спочатку оберіть наказ або кілька наказів.")
            return
        if not self.p2_back_page_path.get() or not self.excel_path.get() or not self.template_path.get():
            messagebox.showwarning(
                "Не вистачає зразків",
                "Для повного циклу потрібні: заготовка примірника, словник Excel і зразок витягу.\n\n"
                "Заповніть їх у вікні «Зразки та реквізити».",
            )
            return

        missing = self._missing_selected_files(
            ("Словник Excel", self.excel_path.get()),
            ("Зразок витягу", self.template_path.get()),
            ("Заготовка примірника (остання сторінка)", self.p2_back_page_path.get()),
        )
        if missing:
            self._report_missing_files(missing)
            return

        self._batch_running = True
        self.btn_full_cycle.config(state=DISABLED)
        try:
            self.log(
                f"\n⚙️ ПОВНИЙ ЦИКЛ для {len(order_paths)} наказ(ів): "
                "примірники → розрахунок → адресні витяги + витяги до управління"
            )
            # 1. Примірники № 2 з тих самих наказів.
            self._set_p2_orders(order_paths)
            self.run_generate_copies()

            # 2 і 3. Далі працюємо з примірниками, якщо вони створились:
            # _set_copy_two_sources уже наповнив список і позначив їх.
            if not self.last_copy_two_paths:
                self.log("Примірники не створено — розрахунок і витяги виконуємо за самими наказами.")
            self.run_rozrahunok_action()
            self.run_extracts_action(scope="all")
            self.log("⚙️ Повний цикл завершено.")
        finally:
            self._batch_running = False
            self.btn_full_cycle.config(state=NORMAL)

    def _run_batch(
        self, order_paths: list[str], handler, stage_label: str, alongside: bool = False
    ) -> list[tuple[str, str]]:
        """Обробляє накази по черзі, не перериваючи пакет через збій одного.

        Повертає перелік `(назва файлу, текст помилки)` для тих, що не вдалися.
        """
        import traceback

        failures: list[tuple[str, str]] = []
        for index, order_path in enumerate(order_paths, start=1):
            name = os.path.basename(order_path)
            self.log(f"\n[{index}/{len(order_paths)}] {stage_label}: {name}")
            try:
                self._set_processing_order(order_path, alongside=alongside)
                handler()
            except Exception as error:
                traceback.print_exc()
                explanation = explain_error(error)
                self.log(f"  ПОМИЛКА під час обробки «{name}»:\n    "
                         + explanation.replace("\n", "\n    "))
                failures.append((name, explanation))
        return failures

    def _report_batch_result(
        self, total: int, failures: list[tuple[str, str]], title: str, action_name: str
    ) -> None:
        """Показує підсумок пакета: скільки вдалося, а що саме — ні."""
        succeeded = total - len(failures)
        if not failures:
            messagebox.showinfo("Успіх", f"{title}: успішно оброблено {succeeded} з {total} файл(ів).")
            return

        details = "\n\n".join(
            f"Наказ «{name}» — не оброблено.\n{error}" for name, error in failures[:10]
        )
        if len(failures) > 10:
            details += f"\n\n… і ще {len(failures) - 10} файл(ів) з такими самими збоями."
        self.log(f"\nНе вдалося виконати {action_name} для {len(failures)} файл(ів).")
        messagebox.showwarning(
            title,
            f"Готово: {succeeded} з {total} файл(ів). Не вийшло: {len(failures)}.\n\n"
            "Нижче названо наказ, на якому програма зупинилась, і причину. "
            "Назва наказу — це те, що оброблялося, а не обов'язково те, чого бракує.\n\n"
            f"{details}",
        )

    def _log_routing_module(self):
        """Пише в журнал, який саме модуль маршрутизації завантажено."""
        mapping_module = sys.modules.get(map_military_units.__module__)
        module_file = getattr(mapping_module, "__file__", "")
        module_version = getattr(mapping_module, "ROUTING_VERSION", "без версії")
        self.log(
            "Модуль маршрутизації: "
            + (os.path.abspath(str(module_file)) if module_file else "невідомо")
            + f" · версія {module_version}"
        )

    def _log_mapping_source(self, excel_res: dict):
        """Пише в журнал реквізити прочитаного Excel-еталона.

        Правило: вибраний Excel — єдиний еталон маршрутизації, і кожен запуск
        читає його з диска без кешу. Журнал має це підтверджувати, щоб було
        видно, що розрахунок зроблено саме за поточним файлом.
        """
        modified_ns = excel_res.get("source_modified_ns")
        if modified_ns:
            source_modified = datetime.fromtimestamp(
                modified_ns / 1_000_000_000
            ).strftime("%d.%m.%Y %H:%M:%S")
        else:
            source_modified = "невідомо"
        self.log(
            f"Еталон Excel: {excel_res.get('source_path', self.excel_path.get())} · "
            f"змінено {source_modified} · {excel_res.get('source_size', 0)} байт. "
            "Файл перечитано з диска без кешу."
        )

    def run_rozrahunok(self):
        self.log("\n=== РОЗРАХУНОК РОЗСИЛКИ ===")
        self._log_routing_module()
        self.log(f"Читаємо словник: {self.excel_path.get()}")
        excel_res = read_recipient_mapping(path=self.excel_path.get())
        self._log_mapping_source(excel_res)
        mapping_base = excel_res.get("mapping", {})
        unique_entries = {id(value): value for value in mapping_base.values() if isinstance(value, dict)}
        abbreviation_count = sum(bool(str(value.get("abbreviation", "")).strip()) for value in unique_entries.values())

        self.log(f"Читаємо наказ: {self.doc_path.get()}")
        source_text = self._read_word_text(self.doc_path.get())
        text, _ = self._refresh_order_signer(source_text)

        self.log(
            f"Зчитано {len(unique_entries)} записів з Excel. "
            f"Скорочень з колонки C: {abbreviation_count}."
        )
        os.makedirs(self.out_folder.get(), exist_ok=True)
        import copy

        mapping_extracts = copy.deepcopy(mapping_base)

        if not self.group_corps_var.get():
            self.log("Режим: БЕЗ угруповання по Корпусах (всі частини окремо)")
            for k, v in mapping_extracts.items():
                if isinstance(v, dict) and "corps" in v:
                    v["corps"] = ""
        else:
            self.log("Режим: З угрупованням по Корпусах (Варіант 1)")

        self.log("Аналізуємо текст наказу (структура, ТЦК, ВЧ)...")
        map_res_extracts = map_military_units(text=text, mapping=mapping_extracts)
        self.show_analysis_results(map_res_extracts)

        for invalid_link in map_res_extracts.get("invalid_corps_links", []):
            self.log(
                "УВАГА: корпус не знайдено окремим рядком у таблиці; "
                f"частину залишено самостійним адресатом: {invalid_link.get('unit', '')} → {invalid_link.get('corps', '')}"
            )
        for tck_reference in map_res_extracts.get("unresolved_tck_references", []):
            self.log(f"УВАГА: ТЦК не визначено або його ОТЦК відсутній у таблиці; пункт не додано до розсилки: {tck_reference}")

        order_base = sanitize_filename(os.path.splitext(os.path.basename(self.doc_path.get()))[0])
        unmatched_items = map_res_extracts.get("unmatched_items", [])
        unmatched_report = os.path.join(self.out_folder.get(), f"Контроль_пропущених_пунктів_{order_base}.xlsx")
        _save_table_to_excel(
            unmatched_report,
            ["Пункт", "Текст пункту", "Причина"],
            [
                (item.get("label", ""), item.get("text", ""), item.get("reason", ""))
                for item in unmatched_items
            ],
        )
        if unmatched_items:
            self.log(f"УВАГА: {len(unmatched_items)} пункт(ів) без адресата. Контрольний файл: {unmatched_report}")
        else:
            self.log(f"Контрольний файл створено: {unmatched_report} (пропущених пунктів немає)")

        routing_report = os.path.join(self.out_folder.get(), f"Контроль_маршрутизації_{order_base}.xlsx")
        routing_data = [
            (
                item.get("label", ""),
                item.get("matched_entries", ""),
                item.get("applied_rules", ""),
                item.get("item_recipients", ""),
                item.get("context_recipients", ""),
                item.get("final_recipients", ""),
            )
            for item in map_res_extracts.get("routing_audit", [])
        ]
        _save_table_to_excel(
            routing_report,
            ["Пункт", "Збіги з таблиці", "Застосовані правила", "Адресати з пункту", "Адресати з контексту", "Підсумкові адресати"],
            routing_data,
        )
        self.log(f"Контроль маршрутизації: {routing_report}")

        table_extracts = map_res_extracts.get("units_table")
        if table_extracts and hasattr(table_extracts, "rows"):
            filtered_rows_ext = [(row[0], row[2], row[3]) for row in table_extracts.rows]
            out_file_ext = os.path.join(self.out_folder.get(), f"Розрахунок_розсилки_{order_base}.xlsx")
            _save_table_to_excel(
                out_file_ext,
                ["Військова частина / Відправник", "Номери пунктів витягу", "Кількість пунктів"],
                filtered_rows_ext,
            )
            self.log(f"Збережено розрахунок: {out_file_ext}")
        else:
            self.log("Помилка: NAT не повернув таблицю.")

        self.log("Готово! Завершено розрахунок розсилки.")

    def run_extracts(self, scope: str = "all"):
        scope_titles = {
            "general": "ГЕНЕРАЦІЯ АДРЕСНИХ ВИТЯГІВ",
            "management": "ГЕНЕРАЦІЯ ВИТЯГІВ ДО УПРАВЛІННЯ",
            "all": "ГЕНЕРАЦІЯ ВСІХ ВИТЯГІВ",
        }
        if scope not in scope_titles:
            raise ValueError(f"Невідомий режим генерації витягів: {scope}")
        self.log(f"\n=== {scope_titles[scope]} ===")
        self._log_routing_module()
        excel_res = read_recipient_mapping(path=self.excel_path.get())
        self._log_mapping_source(excel_res)
        mapping = excel_res.get("mapping", {})
        unique_entries = {id(value): value for value in mapping.values() if isinstance(value, dict)}
        abbreviation_count = sum(bool(str(value.get("abbreviation", "")).strip()) for value in unique_entries.values())
        self.log(
            f"Зчитано {len(unique_entries)} записів з Excel. "
            f"Скорочень з колонки C: {abbreviation_count}."
        )

        source_text = self._read_word_text(self.doc_path.get())
        text, _detected_order_signer = self._refresh_order_signer(source_text)
        # Реквізити підписанта беремо з ПОЛІВ інтерфейсу, а не з розпізнаного
        # напряму: `_refresh_order_signer` щойно заповнив їх із самого наказу,
        # але користувач має право виправити розпізнане — так само, як для
        # засвідчувача. (Тимчасове правило «підписанта у витягах не переносимо»
        # скасоване користувачем 04.09.2026.)
        order_signer = {
            "position": self.order_signer_position.get().strip(),
            "rank": self.order_signer_rank.get().strip(),
            "name": self.order_signer_name.get().strip(),
        }
        if any(order_signer.values()):
            signer_summary = " ".join(
                part for part in (order_signer["rank"], order_signer["name"]) if part
            )
            self.log(
                "Підписант оригіналу наказу переноситься у витяг"
                + (f": {signer_summary}." if signer_summary else ".")
            )
        else:
            self.log(
                "УВАГА: реквізитів підписанта наказу немає — теги {{підписант_…}} "
                "буде прибрано з витягу разом із їхніми рядками."
            )
        filename = os.path.basename(self.doc_path.get())
        order_num, order_date = extract_metadata_from_filename(filename)
        if not order_num:
            self.log("УВАГА: Номер наказу не знайдено в назві файлу. Тег {{номер_наказу}} залишиться для ручного заповнення.")
        if not order_date:
            self.log("УВАГА: Дату наказу не знайдено в назві файлу. Тег {{дата_наказу}} залишиться для ручного заповнення.")

        order_date_formatted = format_ukr_date(order_date)

        if not self.group_corps_var.get():
            self.log("УВАГА: Вибрано режим без угруповання по корпусах для витягів!")
            for k, v in mapping.items():
                if isinstance(v, dict) and "corps" in v:
                    v["corps"] = ""

        self.log("Аналізуємо структуру наказу (блоки, адресати)...")
        map_res = map_military_units(text=text, mapping=mapping)
        self.show_analysis_results(map_res)
        preamble_recipient = str(map_res.get("preamble_recipient") or "").strip()
        if preamble_recipient:
            self.log(f"Адресат у преамбулі знайдено за колонкою A: {preamble_recipient}")
        else:
            self.log("Адресата у преамбулі за колонкою A не знайдено.")

        for invalid_link in map_res.get("invalid_corps_links", []):
            self.log(
                "УВАГА: корпус не знайдено окремим рядком у таблиці; "
                f"частину залишено самостійним адресатом: {invalid_link.get('unit', '')} → {invalid_link.get('corps', '')}"
            )
        for tck_reference in map_res.get("unresolved_tck_references", []):
            self.log(f"УВАГА: ТЦК не визначено або його ОТЦК відсутній у таблиці; витяг не буде створено: {tck_reference}")

        order_base = sanitize_filename(os.path.splitext(os.path.basename(self.doc_path.get()))[0])
        routing_report = os.path.join(
            self.out_folder.get(), f"Контроль_маршрутизації_{order_base}.xlsx"
        )
        routing_data = [
            (
                item.get("label", ""),
                item.get("matched_entries", ""),
                item.get("applied_rules", ""),
                item.get("item_recipients", ""),
                item.get("context_recipients", ""),
                item.get("final_recipients", ""),
            )
            for item in map_res.get("routing_audit", [])
        ]
        _save_table_to_excel(
            routing_report,
            ["Пункт", "Збіги з таблиці", "Застосовані правила", "Адресати з пункту", "Адресати з контексту", "Підсумкові адресати"],
            routing_data,
        )

        unmatched_items = map_res.get("unmatched_items", [])
        missing_report = os.path.join(
            self.out_folder.get(), f"Контроль_пропущених_пунктів_{order_base}.xlsx"
        )
        _save_table_to_excel(
            missing_report,
            ["Пункт", "Текст пункту", "Причина"],
            [
                (item.get("label", ""), item.get("text", ""), item.get("reason", ""))
                for item in unmatched_items
            ],
        )
        if unmatched_items:
            self.log(
                f"УВАГА: {len(unmatched_items)} пункт(ів) без адресата. "
                f"Контрольний файл: {missing_report}"
            )

        # Рахуємо записи аудиту, а не унікальні мітки: однакова мітка («1.»)
        # трапляється в різних § і тоді множина злила б їх в один пункт.
        audited_items = map_res.get("routing_audit", [])
        routed_count = sum(
            1
            for item in audited_items
            if str(item.get("final_recipients", "")).strip() not in ("", "—")
        )
        skipped_items = map_res.get("skipped_items", [])
        self.log(
            f"Розібрано пунктів: {len(audited_items)}; "
            f"з адресатами: {routed_count}; без адресата: {len(unmatched_items)}; "
            f"виключено із загального переліку (зміна до управління): {len(skipped_items)}."
        )

        units_data, management_data = select_extracts_for_scope(map_res, scope)
        # Зміни по управлінню нікому не розсилаються, тому вони НЕ входять до
        # розрахунку розсилки. Але витяг за таким пунктом усе одно потрібен —
        # окремим файлом, без адресата й без компонування під друк «2 на 1».
        if not units_data and not management_data:
            if scope == "management":
                empty_message = "У вибраному наказі немає пунктів для витягів до управління."
            elif scope == "general":
                empty_message = "У вибраному наказі немає адресних витягів для військових частин."
            else:
                empty_message = "Жодної військової частини або пункту до управління не знайдено."
            self.log(empty_message)
            messagebox.showwarning("Результат", empty_message)
            return
        if management_data:
            self.log(
                f"Витяги до управління (окремий файл, поза розрахунком розсилки): "
                f"{len(management_data)} — " + ", ".join(management_data)
            )

        # Склад кожного витягу з рядками наказу, з яких береться пункт. Саме тут
        # видно причину задвоєного пункту: два записи з ОДНАКОВИМИ рядками або
        # шапка, чиї рядки перекриваються з рядками самого пункту.
        composition_report = os.path.join(
            self.out_folder.get(), f"Контроль_складу_витягів_{order_base}.xlsx"
        )
        composition_rows = []
        duplicate_spans = 0
        for unit_key, unit_data in {**units_data, **management_data}.items():
            seen_spans: dict[tuple, str] = {}
            for unit_item in sorted(
                unit_data.get("items", []),
                key=lambda value: value.get("source_start_line", 10**9),
            ):
                span = (unit_item.get("source_start_line"), unit_item.get("source_end_line"))
                headings = ", ".join(
                    f"{heading[0]}–{heading[1]}"
                    for heading in (unit_item.get("heading_ranges") or [])
                    if isinstance(heading, (list, tuple)) and len(heading) == 2
                )
                note = ""
                if span in seen_spans:
                    note = f"ДУБЛЬ рядків із пунктом «{seen_spans[span]}»"
                    duplicate_spans += 1
                else:
                    seen_spans[span] = unit_item.get("label", "")
                composition_rows.append(
                    (
                        unit_key,
                        unit_item.get("label", ""),
                        f"{span[0]}–{span[1]}",
                        headings or "—",
                        (unit_item.get("text", "") or "").strip()[:120],
                        note,
                    )
                )
        _save_table_to_excel(
            composition_report,
            ["Витяг", "Пункт", "Рядки наказу", "Рядки шапок", "Початок тексту", "Примітка"],
            composition_rows,
        )
        if duplicate_spans:
            self.log(
                f"УВАГА: маршрутизація віддала {duplicate_spans} пункт(ів) із тими самими "
                f"рядками наказу — у витягу вони не задвоюються, але причину видно у файлі "
                f"{os.path.basename(composition_report)}."
            )

        # Пункт, який зник БЕЗ СЛІДУ: його номер є в тексті наказу, але він не
        # потрапив ні у витяги, ні в «Пропущені», ні у виключені. Так буває,
        # коли пункт поглинула шапка розділу або сусідній пункт — тоді він
        # узагалі не існує як окремий блок, і жоден інший контроль його не
        # показує. Рахуємо лише «дірки» всередині наявної нумерації, щоб не
        # чіплятися до нумерованих рядків усередині самих пунктів.
        item_number_re = re.compile(r"^\s*(\d{1,3})\s*[.)]")

        def numbers_from_labels(labels) -> set[int]:
            found = set()
            for label in labels:
                match = re.search(r"(\d{1,3})", str(label))
                if match:
                    found.add(int(match.group(1)))
            return found

        numbers_in_text = {
            int(match.group(1))
            for match in (item_number_re.match(line) for line in text.splitlines())
            if match
        }
        numbers_reached = numbers_from_labels(
            [item.get("label", "") for data in units_data.values() for item in data.get("items", [])]
            + [item.get("label", "") for item in unmatched_items]
            + [item.get("label", "") for item in skipped_items]
        )
        lost_numbers = sorted(
            number
            for number in numbers_in_text - numbers_reached
            if numbers_reached and number < max(numbers_reached)
        )
        if lost_numbers:
            self.log(
                "УВАГА: у тексті наказу є пункт(и) "
                + ", ".join(str(number) for number in lost_numbers)
                + ", але їх немає ні у витягах, ні в контрольних переліках. "
                "Найімовірніше, такий пункт поглинула шапка розділу або сусідній "
                f"пункт — подивіться діапазони рядків у файлі {os.path.basename(composition_report)}."
            )

        template_path = os.path.abspath(self.template_path.get())
        out_file = os.path.join(
            self.out_folder.get(), build_extracts_filename(order_num, order_date)
        )
        os.makedirs(self.out_folder.get(), exist_ok=True)
        # Наприкінці генерації результат відкривається у Word, тому при
        # повторному запуску він може бути ще зайнятий. Перевіряємо одразу,
        # щоб не витрачати час на обробку й показати зрозумілу причину.
        management_out_file = (
            os.path.join(
                self.out_folder.get(),
                build_extracts_filename(order_num, order_date, "Витяги до управління за наказом"),
            )
            if management_data
            else None
        )
        busy_candidates = []
        if units_data:
            busy_candidates.append(out_file)
        if management_out_file:
            busy_candidates.append(management_out_file)
        for busy_candidate in busy_candidates:
            if busy_candidate and not is_path_writable(busy_candidate):
                raise UserError(
                    f"файл витягів «{os.path.basename(busy_candidate)}» уже відкритий — "
                    "найімовірніше у Word, з минулого разу.",
                    "Закрийте це вікно Word і натисніть кнопку ще раз. "
                    "Поки файл відкритий, програма не може записати в нього нові витяги.",
                )
        temp_dir = os.path.join(self.out_folder.get(), "_nat_temp")
        os.makedirs(temp_dir, exist_ok=True)

        temp_files = []
        management_temp_files = []
        layout_warnings = []

        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0  # wdAlertsNone: не показувати блокуючі діалоги (напр. "Зберегти зміни?")

        word_settings = tune_word_for_batch(word)
        batch_started_at = time.monotonic()

        try:
            source_path = os.path.abspath(self.doc_path.get())
            source_doc = word.Documents.Open(source_path, ReadOnly=True)

            self.log("Індексуємо абзаци оригіналу наказу...")
            source_paragraphs = iter_paragraphs(source_doc)
            para_count = len(source_paragraphs)
            source_line_to_para = []
            for pi, source_paragraph in enumerate(source_paragraphs, start=1):
                raw = source_paragraph.Range.Text
                logical_lines = raw.rstrip("\r\x07").splitlines() or [""]
                source_line_to_para.extend([pi] * len(logical_lines))

            usable_line_count = len(text.splitlines())
            source_line_to_para = source_line_to_para[:usable_line_count]
            self.log(f"Проіндексовано {para_count} абзаців.")

            def has_manual_page_break(para_or_range):
                r = getattr(para_or_range, "Range", para_or_range)
                return "\x0c" in getattr(r, "Text", "")

            def is_blank_paragraph(para_or_range):
                r = getattr(para_or_range, "Range", para_or_range)
                return not getattr(r, "Text", "").strip("\r\x07\v\f \t")

            def clean_redundant_blanks():
                """Прибирає зайві «висячі» порожні абзаци в самому кінці документа,
                щоб дотримати правило 5.3 AGENT.md: 0 порожніх абзаців наприкінці."""
                try:
                    while doc.Paragraphs.Count > 1:
                        tail_paragraph = last_paragraph(doc)
                        if not is_blank_paragraph(tail_paragraph):
                            break
                        # Останній знак абзацу документа (і останній у комірці)
                        # Word не видаляє й помилки не кидає. Без перевірки, що
                        # документ справді скоротився, цикл був би вічним.
                        length_before = doc.Content.End
                        tail_paragraph.Range.Delete()
                        if doc.Content.End >= length_before:
                            break
                except Exception:
                    pass

            def normalize_signature_gap(signer_start, content_end_pos):
                """Нормалізує відступ перед підписантом: рівно 2 порожні абзаци.

                Після видалення/додавання абзаців позиції зсуваються, тому
                підписант шукається заново по непорожньому тексту після content_end_pos.
                """
                document_paragraphs = iter_paragraphs(doc)
                signer_index = None
                for paragraph_index, paragraph in enumerate(document_paragraphs, start=1):
                    if paragraph.Range.Start == signer_start:
                        signer_index = paragraph_index
                        break
                if signer_index is None:
                    return signer_start

                blank_indexes = []
                for paragraph_index in range(signer_index - 1, 0, -1):
                    if not is_blank_paragraph(document_paragraphs[paragraph_index - 1]):
                        break
                    blank_indexes.append(paragraph_index)

                # Перед підписантом залишаємо рівно два порожні абзаци
                while len(blank_indexes) < 2:
                    new_p = doc.Paragraphs.Add(doc.Paragraphs(signer_index).Range)
                    new_p.Range.Text = "\r"
                    blank_indexes.append(signer_index)
                    signer_index += 1
                if len(blank_indexes) > 2:
                    for paragraph_index in sorted(blank_indexes[2:], reverse=True):
                        doc.Paragraphs(paragraph_index).Range.Delete()

                # Після видалення/вставки — позиції зсунулись, шукаємо підписанта
                # заново, але лише в хвості документа після вставленого змісту.
                for paragraph in iter_paragraphs(doc.Range(content_end_pos, doc.Content.End)):
                    paragraph_range = paragraph.Range
                    if paragraph_range.Start >= content_end_pos and not is_blank_paragraph(paragraph_range):
                        return paragraph_range.Start
                return signer_start


            def line_span_to_paragraphs(start_line, end_line):
                if not isinstance(start_line, int) or not isinstance(end_line, int):
                    return []
                if start_line < 0 or end_line < start_line:
                    return []
                # Обрізаємо до останнього валідного індексу замість повернення []
                end_line = min(end_line, len(source_line_to_para) - 1)
                if start_line >= len(source_line_to_para):
                    return []
                return list(dict.fromkeys(source_line_to_para[start_line:end_line + 1]))

            def bind_range(range_object):
                """Зчіплює ВСІ абзаци діапазону, включно з порожніми всередині.

                Word застосовує ParagraphFormat діапазону до кожного його абзацу,
                тож це два звернення замість двох на КОЖЕН абзац.
                """
                pf = range_object.ParagraphFormat
                pf.KeepTogether = True
                pf.KeepWithNext = True

            def release_last_paragraph(range_object):
                """Знімає KeepWithNext з ОСТАННЬОГО абзацу діапазону — саме там
                дозволено розрив сторінки (між пунктами, після ланцюга)."""
                try:
                    last_paragraph(range_object).Range.ParagraphFormat.KeepWithNext = False
                except Exception:
                    pass

            def bind_span(start_position, end_position):
                """Зчіплює діапазон, заданий позиціями."""
                if start_position is None or end_position is None:
                    return
                if end_position <= start_position:
                    return
                bind_range(doc.Range(start_position, end_position))

            def apply_keep_rules(item_ranges, heading_ranges, heading_item_pairs, signer_start, executor_start=None):
                # Рідкісний Word-баг: FormattedText переносить KeepWithNext із
                # наказу, а вставлений після пункту порожній абзац іноді
                # успадковує цей прапорець. Тоді два сусідні пункти випадково
                # стають одним великим нерозривним блоком і псують підбір
                # інтервалу. Спочатку скидаємо лише зчеплення, а нижче заново
                # накладаємо всі наші правила шапок, пунктів і підписанта.
                all_content_ranges = [*heading_ranges, *item_ranges]
                if all_content_ranges:
                    keep_start = min(value.Start for value in all_content_ranges)
                    keep_end = signer_start
                    if keep_end is None:
                        keep_end = max(value.End for value in all_content_ranges)
                    if keep_end > keep_start:
                        doc.Range(keep_start, keep_end).ParagraphFormat.KeepWithNext = False

                # 1. Шапки (§, розділ, підрозділ) зчіплюються з ПЕРШИМ пунктом
                # розділу — разом з порожніми абзацами між ними.
                #
                # Раніше KeepWithNext діставався лише абзацам самої шапки. Але
                # після кожної шапки генератор вставляє порожній абзац, і саме
                # він лишався без зчеплення: Word чесно тримав шапку з наступним
                # абзацом (порожнім) і мав повне право рвати сторінку одразу за
                # ним. Через це шапка лишалась унизу сторінки, а пункт їхав на
                # наступну. Зчіплюємо суцільним діапазоном — так само, як це
                # зроблено для фінального ланцюга в пункті 3 нижче.
                #
                # Зчіплюємо кожну шапку з тим, що стоїть ПІД нею — до початку
                # наступного вставленого блока (це або наступний рівень шапки,
                # або сам пункт). Так само тримається й шапка, для якої пари
                # «шапка → пункт» не склалось: інакше вона висіла б сама.
                blocks = [(heading_range.Start, "heading", heading_range) for heading_range in heading_ranges]
                blocks += [(item_range.Start, "item", item_range) for item_range in item_ranges]
                blocks.sort(key=lambda block: block[0])
                for index, (start_position, kind, range_object) in enumerate(blocks):
                    if kind != "heading":
                        continue
                    if index + 1 < len(blocks):
                        bind_span(start_position, blocks[index + 1][0])
                    else:
                        bind_range(range_object)

                # 2. Пункти витягу: кожен пункт цілісний
                for item_index, item_range in enumerate(item_ranges):
                    bind_range(item_range)
                    if item_index != len(item_ranges) - 1:
                        # Після пункту сторінку рвати можна — крім останнього
                        # пункту, який зчеплений з підписантом (ланцюг нижче).
                        release_last_paragraph(item_range)

                # 3. Суцільний нерозривний ланцюг: [Шапка перед останнім пунктом (якщо є)] -> [Останній пункт] -> [Відступ] -> [Підписант] -> [Згідно з оригіналом]
                # Гарантує, що шапка, останній пункт і підписант становлять неподільний блок і вміщуються разом!
                if item_ranges:
                    last_item_range = item_ranges[-1]
                    chain_start = last_item_range.Start
                    for heading_range, first_item_range, _ in heading_item_pairs:
                        if first_item_range.Start == last_item_range.Start:
                            chain_start = heading_range.Start
                            break

                    signature_end = executor_start if executor_start is not None else doc.Content.End
                    chain_range = doc.Range(chain_start, signature_end)
                    bind_range(chain_range)
                    release_last_paragraph(chain_range)

            def layout_issues(item_ranges, item_labels, heading_item_pairs, signer_start=None, executor_start=None):
                issues = []
                for item_range, label in zip(item_ranges, item_labels):
                    start_page, end_page = range_pages(item_range)
                    if start_page is not None and start_page != end_page:
                        issues.append(f"{label} не вміщується на одну сторінку")

                for heading_range, first_item_range, label in heading_item_pairs:
                    heading_start, heading_end = range_pages(heading_range)
                    item_start, _ = range_pages(first_item_range)
                    if (
                        heading_start is not None
                        and (heading_start != heading_end or heading_start != item_start)
                    ):
                        issues.append(f"шапка перед {label} відірвана від першого пункту")

                # Контроль нерозривності останнього пункту (та його шапки) з підписантом та блоком «Згідно з оригіналом»:
                if item_ranges and signer_start is not None:
                    last_item_range = item_ranges[-1]
                    last_item_start_p, _ = range_pages(last_item_range)
                    signer_p = page_of(signer_start)
                    sig_end_pos = max(signer_start, (executor_start - 1) if executor_start is not None else (doc.Content.End - 1))
                    zg_end_p = page_of(sig_end_pos)

                    # Перевірка для шапки перед останнім пунктом:
                    for heading_range, first_item_range, label in heading_item_pairs:
                        if first_item_range.Start == last_item_range.Start:
                            h_start, h_end = range_pages(heading_range)
                            if h_start is not None and signer_p is not None and h_start != signer_p:
                                issues.append("шапка, останній пункт та підписант розділені на різні сторінки")

                    if last_item_start_p is not None and signer_p is not None and last_item_start_p != signer_p:
                        issues.append("підписант та блок «Згідно з оригіналом» відірвані від останнього пункту")
                    elif signer_p is not None and zg_end_p is not None and signer_p != zg_end_p:
                        issues.append("блок «Згідно з оригіналом» відірваний від підписанта наказу")

                return list(dict.fromkeys(issues))

            def find_executor_paragraph_index(bookmark_name=None):
                """Знаходить номер абзацу виконавця за закладкою, текстом або як останній непорожній абзац."""
                # 1. За закладкою — точна позиція підстановки {{виконавець}}
                if bookmark_name and doc.Bookmarks.Exists(bookmark_name):
                    try:
                        b_pos = doc.Bookmarks(bookmark_name).Range.Start
                        for pi, paragraph in enumerate(iter_paragraphs(doc), start=1):
                            pr = paragraph.Range
                            if pr.Start <= b_pos < pr.End:
                                return pi
                    except Exception:
                        pass

                document_paragraphs = iter_paragraphs(doc)

                # 2. За текстом виконавця (фолбек, якщо закладку втрачено)
                exec_val = self.executor.get().strip()
                if exec_val:
                    lines = [ln.strip() for ln in exec_val.replace("/", "\n").splitlines() if ln.strip()]
                    if lines:
                        first_line = lines[0].lower()
                        for pi in range(len(document_paragraphs), 0, -1):
                            pt = document_paragraphs[pi - 1].Range.Text.strip().lower()
                            if first_line in pt or (len(first_line) > 4 and first_line[:4] in pt):
                                return pi

                # 3. Фолбек: останній непорожній абзац у документі
                for pi in range(len(document_paragraphs), 0, -1):
                    if not is_blank_paragraph(document_paragraphs[pi - 1]):
                        return pi
                return None

            def executor_start_from_bookmark(bookmark_name):
                """Повертає позицію (Start) закладки виконавця, або None, якщо її немає."""
                if bookmark_name and doc.Bookmarks.Exists(bookmark_name):
                    try:
                        return doc.Bookmarks(bookmark_name).Range.Start
                    except Exception:
                        return None
                return None

            def executor_paragraph_start(bookmark_name):
                """Початок АБЗАЦУ виконавця без проходу по всіх абзацах документа.

                Закладка віддає позицію одразу, а `Range(pos, pos).Paragraphs(1)`
                — її абзац. Потрібно для меж, у яких дозволено міняти інтервал:
                блок виконавця не чіпаємо ніколи.
                """
                position = executor_start_from_bookmark(bookmark_name)
                if position is None:
                    return None
                try:
                    return doc.Range(position, position).Paragraphs(1).Range.Start
                except Exception:
                    return None

            def executor_paragraph(bookmark_name=None):
                """Абзац виконавця як ОБ'ЄКТ — із закладки, без обходу документа.

                Увага: сама закладка для цього не годиться після вставок —
                вставлений перед нею текст Word включає В НЕЇ. Тому позицію з
                закладки беремо один раз, а далі тримаємось за абзац.
                """
                position = executor_start_from_bookmark(bookmark_name)
                if position is not None:
                    try:
                        return doc.Range(position, position).Paragraphs(1)
                    except Exception:
                        pass
                index = find_executor_paragraph_index(bookmark_name)
                if index is None:
                    return None
                try:
                    return doc.Paragraphs(index)
                except Exception:
                    return None

            def previous_paragraph(paragraph):
                """Попередній абзац або None (на початку документа Word кидає помилку)."""
                try:
                    return paragraph.Previous()
                except Exception:
                    return None

            def paragraph_in_table(paragraph):
                """Чи стоїть абзац усередині таблиці (wdWithInTable)."""
                try:
                    return bool(paragraph.Range.Information(12))
                except Exception:
                    # Не змогли визначити — вважаємо, що в таблиці: чіпати
                    # невідомий абзац небезпечніше, ніж лишити його на місці.
                    return True

            def delete_paragraph_if_possible(paragraph):
                """Видаляє абзац і повертає True, лише якщо він СПРАВДІ зник.

                Word мовчки ігнорує видалення знака абзацу, який прибрати не
                можна: останнього в документі та останнього в комірці таблиці.
                Помилки при цьому не буде, тож цикл, який спирається на
                `Delete()`, без перевірки прогресу крутиться вічно.
                """
                try:
                    before = doc.Content.End
                    paragraph.Range.Delete()
                    return doc.Content.End < before
                except Exception:
                    return False

            def position_executor_at_page_bottom(bookmark_name=None, signer_start=None, needs_manual_review=False):
                """Вирівнює виконавця до низу сторінки, окрім випадків, де це неможливо
                без наставляння штучних порожніх абзаців — тоді лишаємо позицію зі зразка
                (правило 5.5 AGENT.md, виключення №2).

                needs_manual_review навмисно НЕ блокує спробу підштовхування: загальна
                проблема макета деінде у витягу (напр. інший пункт не вміщується) не
                означає, що саме область виконавця не можна коректно розмістити.
                Єдина причина відкату — фізична неможливість (перехід на іншу сторінку),
                що обробляється нижче в самому циклі."""
                # Працюємо з АБЗАЦОМ-ОБ'ЄКТОМ, а не з його номером: `doc.Paragraphs(i)`
                # щоразу проходить документ, а самих звернень тут десятки. Абзац
                # Word тримає прив'язку до свого тексту навіть після вставок перед
                # ним (перевірено), тож об'єкт лишається чинним увесь підбір.
                executor = executor_paragraph(bookmark_name)
                if executor is None:
                    clean_redundant_blanks()
                    return

                # Запам'ятовуємо кількість порожніх абзаців зі зразка перед виконавцем,
                # щоб мати змогу відновити цю позицію, якщо автоматичне вирівнювання
                # донизу не вдасться виконати чисто (без переходу на іншу сторінку).
                # Порожні комірки таблиці порожніми абзацами зразка НЕ рахуємо:
                # у шаблоні з табличним підписантом (див. drop_blanks_before_executor)
                # виконавцю передує саме комірка, а не відступ.
                original_blank_count = 0
                probe = previous_paragraph(executor)
                while (
                    probe is not None
                    and not paragraph_in_table(probe)
                    and is_blank_paragraph(probe)
                ):
                    original_blank_count += 1
                    probe = previous_paragraph(probe)

                def drop_blanks_before_executor():
                    """Знімає порожні абзаци безпосередньо перед виконавцем.

                    Зупиняємось на межі таблиці: якщо блок підписанта/засвідчувача
                    у зразку зверстано таблицею, попередній абзац — це остання
                    (часто порожня) КОМІРКА. Її знак абзацу Word не видаляє й не
                    повідомляє про це помилкою, тож цикл крутився вічно — програма
                    зависала на першому ж витягу. Комірка тут і не є відступом:
                    зсовувати виконавця треба порожніми абзацами ПІСЛЯ таблиці.
                    """
                    while True:
                        previous = previous_paragraph(executor)
                        if (
                            previous is None
                            or paragraph_in_table(previous)
                            or not is_blank_paragraph(previous)
                        ):
                            return previous
                        if not delete_paragraph_if_possible(previous):
                            return previous

                previous_before_executor = drop_blanks_before_executor()
                if previous_before_executor is not None:
                    try:
                        previous_before_executor.Range.ParagraphFormat.KeepWithNext = False
                    except Exception:
                        pass

                clean_redundant_blanks()
                doc.Repaginate()

                def restore_sample_position():
                    """Відновлює рівно original_blank_count порожніх абзаців перед
                    виконавцем (позиція зі зразка) без штучного доштовхування донизу."""
                    drop_blanks_before_executor()
                    if original_blank_count:
                        position = executor.Range.Start
                        doc.Range(position, position).InsertBefore("\r" * original_blank_count)
                    doc.Repaginate()

                # Опускаємо виконавця до самого низу поточної сторінки.
                #
                # ВАЖЛИВО: вертикальну позицію (Information(6),
                # wdVerticalPositionRelativeToPage) тут НЕ використовуємо — у
                # прихованому екземплярі Word (Visible=False) вона стабільно кидає
                # E_FAIL (-2147467259), через що опускання мовчки не спрацьовувало.
                # Орієнтуємось виключно на номер сторінки (Information(3)), який
                # працює надійно: додаємо порожні абзаци, доки виконавець не
                # перескочить на наступну сторінку, і прибираємо останню порцію —
                # це і є найнижча позиція, яка ще вміщується на сторінці.
                # Кількість ентерів підбираємо ДВІЙКОВИМ пошуком: ознака
                # «виконавець ще на своїй сторінці» монотонна (що більше
                # ентерів, то нижче), тож замість до 38 покрокових вставок,
                # кожна зі своїм Repaginate, вистачає ~7 вимірювань.
                current_enters = 0

                def set_enters(count):
                    """Доводить кількість порожніх абзаців перед виконавцем до `count`."""
                    nonlocal current_enters
                    position = executor.Range.Start
                    if count > current_enters:
                        doc.Range(position, position).InsertBefore("\r" * (count - current_enters))
                    elif count < current_enters:
                        # Кожен доданий порожній абзац — рівно один символ,
                        # тож зайві знімаємо одним діапазоном, а не по одному.
                        delta = current_enters - count
                        doc.Range(max(0, position - delta), position).Delete()
                    current_enters = count
                    doc.Repaginate()

                def executor_still_on_page(count, start_page):
                    set_enters(count)
                    return page_of(executor.Range.Start) == start_page

                try:
                    start_page = page_of(executor.Range.Start)
                    if start_page is None:
                        self.log(f"  ⚠️ {cipher}: не вдалося визначити сторінку виконавця — лишено як у зразку.")
                    else:
                        # fits — найбільша кількість ентерів, яка ще тримає
                        # виконавця на його сторінці; overflow — перша, яка вже ні.
                        fits, overflow = 0, 80
                        if executor_still_on_page(overflow, start_page):
                            fits = overflow
                        else:
                            while overflow - fits > 1:
                                middle = (fits + overflow) // 2
                                if executor_still_on_page(middle, start_page):
                                    fits = middle
                                else:
                                    overflow = middle
                            set_enters(fits)
                        self.log(f"  ↧ {cipher}: виконавця опущено вниз сторінки (+{fits} ентер(ів)).")
                except Exception as exec_error:
                    self.log(f"  ⚠️ {cipher}: помилка опускання виконавця: {exec_error}")
                    try:
                        # Повертаємо відступ зі зразка, щоб не лишити виконавця
                        # без жодного порожнього абзацу після невдалої спроби.
                        restore_sample_position()
                    except Exception:
                        pass

                clean_redundant_blanks()

            def page_of(position):
                try:
                    return doc.Range(position, position).Information(3)  # wdActiveEndPageNumber
                except Exception:
                    return None

            def range_pages(item_range):
                start_page = page_of(item_range.Start)
                end_position = max(item_range.Start, item_range.End - 1)
                return start_page, page_of(end_position)

            signer_pattern = re.compile(
                r"^\s*(?:командир|командувач|начальник|заступник|керівник|голова|директор|"
                r"т\.?\s*в\.?\s*о\.?|тимчасово\s+виконуюч(?:ий|а)?)",
                re.IGNORECASE,
            )

            # Витяги управління формуються тим самим проходом і за тим самим
            # зразком: різниця лише в порожніх {{кому}}/{{куди}} та в тому, що
            # вони збираються в окремий файл без компонування під друк.
            generation_plan = [(key, value, False) for key, value in units_data.items()]
            generation_plan += [(key, value, True) for key, value in management_data.items()]
            self.progress_begin(len(generation_plan) + 1, "Витяги")
            for idx, (cipher, data, is_management) in enumerate(generation_plan):
                self.log(f"[{idx+1}/{len(generation_plan)}] Генеруємо витяг для: {cipher}")
                self.progress_step(idx, cipher)
                extract_started_at = time.monotonic()

                extract_needs_manual_review = False
                # Початок вставленого змісту: усе, що ВИЩЕ цієї позиції (герб,
                # {{кому}}/{{куди}}, назва, дата/номер) — недоторкана шапка зразка.
                content_start_pos = None
                signer_start = None
                # Робимо копію шаблону, щоб Word COM ніколи не відкривав і не
                # змінював оригінальний файл шаблону. Розширення копії має
                # відповідати реальному вмісту шаблону (він може бути у форматі
                # Word 97-2003), інакше Word відмовиться її відкрити.
                temp_path = copy_template_for_editing(
                    template_path,
                    os.path.join(temp_dir, f"extract_{idx:04d}.docx"),
                    label="зразок витягу",
                )
                doc = word.Documents.Open(os.path.abspath(temp_path), ReadOnly=False)
                # Текст шаблону читаємо один раз: більшості з тих двох десятків
                # тегів у конкретному зразку немає (різні регістри «кому/КОМУ»,
                # окремі теги засвідчувача), а порожній пошук Word коштує
                # дорожче, ніж це читання.
                try:
                    template_text = str(doc.Content.Text or "").casefold()
                except Exception:
                    template_text = None

                def replace_tag(tag, replacement_text, document=doc, collect_paragraphs=False,
                               highlight_red=False, bold_pattern=None):
                    if template_text is not None and tag.casefold() not in template_text:
                        return []
                    replaced_paragraphs = []
                    find_obj = document.Content.Find
                    find_obj.Text = tag
                    while find_obj.Execute():
                        found_start = find_obj.Parent.Start
                        find_obj.Parent.Text = str(replacement_text)
                        found_end = found_start + len(str(replacement_text))
                        repl_range = document.Range(found_start, found_end)
                        if highlight_red:
                            try:
                                repl_range.Font.Color = 255  # wdColorRed (RGB 255, 0, 0)
                                repl_range.Font.Bold = 1
                            except Exception:
                                pass
                        if bold_pattern:
                            # Жирним виділяється лише ЧАСТИНА підставленого
                            # тексту: цифри дня в даті та номер наказу без «№».
                            match = re.search(bold_pattern, str(replacement_text))
                            if match:
                                try:
                                    document.Range(
                                        found_start + match.start(), found_start + match.end()
                                    ).Font.Bold = 1
                                except Exception:
                                    pass
                        if collect_paragraphs:
                            replaced_paragraphs.append(
                                document.Range(found_start, found_start).Paragraphs(1).Range.Duplicate
                            )
                        find_obj = document.Content.Find
                        find_obj.Text = tag
                    return replaced_paragraphs

                def replace_tag_group(tag, replacement_text, **kwargs):
                    """Підставляє значення в усі синоніми тега зі зразка.

                    У заготовках трапляються і «засвідчувач», і «затверджувач»,
                    і «згідно_з_оригіналом» — це те саме поле, тому шаблон не
                    повинен залежати від того, яку назву обрав автор.
                    """
                    for alias in tag_aliases(tag):
                        replace_tag(alias, replacement_text, **kwargs)

                signer_template_tags = SIGNER_TAGS

                def remove_original_signer_template_block():
                    """Видаляє з шаблону весь блок тегів підписанта оригіналу.

                    Викликається лише тоді, коли реквізитів підписанта немає
                    зовсім. Якщо тег стоїть у таблиці, видаляється відповідний
                    рядок, щоб не лишалися порожні комірки.
                    """
                    # У зразку без цих тегів обходити абзаци нема сенсу.
                    if template_text is not None and not any(
                        tag.casefold() in template_text for tag in signer_template_tags
                    ):
                        return
                    # Щоразу шукаємо ПЕРШИЙ абзац із тегом заново. Зібраний
                    # наперед список тут не годиться: після видалення рядка
                    # таблиці абзаци-об'єкти з наступних рядків «сповзають» на
                    # сусідній рядок, і разом із підписантом зникав рядок
                    # «Згідно з оригіналом». Тегів одиниці, а шаблон на цьому
                    # кроці ще без змісту — перепрохід коштує копійки.
                    for _ in range(len(signer_template_tags) * 4):  # запобіжник від вічного циклу
                        target_range = None
                        for paragraph in iter_paragraphs(doc):
                            paragraph_range = paragraph.Range
                            text_of_paragraph = str(paragraph_range.Text or "").casefold()
                            if any(tag.casefold() in text_of_paragraph for tag in signer_template_tags):
                                target_range = paragraph_range
                                break
                        if target_range is None:
                            return
                        try:
                            if target_range.Information(12):  # wdWithInTable
                                target_range.Cells(1).Row.Delete()
                            else:
                                length_before = doc.Content.End
                                target_range.Delete()
                                if doc.Content.End >= length_before:
                                    # Знак абзацу Word не віддав (останній у
                                    # документі) — прибираємо принаймні тег.
                                    target_range.Text = ""
                        except Exception:
                            try:
                                # Абзац видалити не вдалось — прибираємо хоча б
                                # текст, інакше наступний прохід знайде той самий.
                                target_range.Text = ""
                            except Exception:
                                return

                if is_management:
                    # Такий витяг нікуди не надсилається: адресні теги просто
                    # стираємо, а не підсвічуємо червоним для ручної правки.
                    rec_to_val = ""
                    dest_where_val = ""
                    is_dest_manual = False
                else:
                    rec_to_val = data.get("recipient_to") or cipher
                    dest_where_val = (data.get("destination_where") or "").strip()
                    is_dest_manual = not bool(dest_where_val) or dest_where_val.upper() in ("КУДИ", "[КУДИ]")
                    if is_dest_manual:
                        dest_where_val = "КУДИ"

                for tag_var in ("{{кому}}", "{{Кому}}", "{{КОМУ}}"):
                    replace_tag(tag_var, rec_to_val)
                for tag_var in ("{{куди}}", "{{Куди}}", "{{КУДИ}}"):
                    replace_tag(tag_var, dest_where_val, highlight_red=is_dest_manual)
                if order_date_formatted:
                    # Жирним — лише цифри дня: “29” серпня 2026 року.
                    replace_tag_group("{{дата_наказу}}", order_date_formatted,
                                      bold_pattern=r"\d+")
                if order_num:
                    # Жирним — лише номер, знак «№» лишається звичайним.
                    replace_tag_group("{{номер_наказу}}", f"№{order_num}",
                                      bold_pattern=r"(?<=№).+")

                # Підписант оригіналу наказу
                if any(order_signer.values()):
                    for signer_tag, signer_value in signer_tags(
                        _slash_to_lines(order_signer["position"]),
                        order_signer["rank"],
                        order_signer["name"],
                    ).items():
                        if signer_value:
                            replace_tag(signer_tag, signer_value)
                    # Реквізит, якого в наказі немає (буває звання), просто
                    # стираємо. Видаляти тут рядок таблиці НЕ можна: у ньому
                    # стоять сусідні, вже заповнені теги підписанта.
                    for empty_tag in signer_template_tags:
                        replace_tag(empty_tag, "")
                else:
                    # Реквізитів немає зовсім — прибираємо блок разом з
                    # абзацами та рядками таблиці, щоб не лишити порожнечі.
                    remove_original_signer_template_block()

                # Особа, яка засвідчує витяг («Згідно з оригіналом» / Засвідчувач)
                replace_tag_group("{{засвідчення}}", "Згідно з оригіналом")
                for cert_tag, cert_value in certifier_tags(
                    _slash_to_lines(self.certifier_position.get().strip()),
                    self.certifier_rank.get().strip(),
                    self.certifier_name.get().strip(),
                ).items():
                    if cert_value:
                        replace_tag(cert_tag, cert_value)

                executor_bookmark = None
                if self.executor.get().strip():
                    executor_paragraphs = replace_tag(
                        "{{виконавець}}", _slash_to_lines(self.executor.get().strip()), collect_paragraphs=True
                    )
                    if executor_paragraphs:
                        executor_bookmark = f"nat_executor_{idx}"
                        doc.Bookmarks.Add(executor_bookmark, executor_paragraphs[0])
                        if len(executor_paragraphs) > 1:
                            layout_warnings.append(
                                f"{cipher}: у зразку кілька тегів {{виконавець}}; донизу вирівняно лише перший."
                            )

                raw_labels = [item.get("label", "") for item in data.get("items", [])]
                points_text = _format_item_numbers_range(raw_labels)
                replace_tag("{{пункти}}", points_text)

                # Вставка змісту
                find_zmist = doc.Content.Find
                find_zmist.Text = "{{зміст}}"
                if find_zmist.Execute():
                    zmist_rng = find_zmist.Parent
                    items = sorted(
                        data.get("items", []),
                        key=lambda item: item.get("source_start_line", 10**9),
                    )
                    zmist_para = zmist_rng.Paragraphs(1).Range
                    insert_point = zmist_para.Start
                    content_start_pos = insert_point
                    zmist_para.Delete()
                    inserted_item_ranges = []
                    inserted_item_labels = []
                    inserted_heading_ranges = []
                    heading_item_pairs = []
                    copied_heading_keys = ()
                    extract_needs_manual_review = False

                    def insert_empty_paragraph():
                        """Вставляє один порожній абзац у поточну позицію змісту."""
                        nonlocal insert_point
                        doc.Range(insert_point, insert_point).InsertBefore("\r")
                        insert_point += 1

                    def source_runs(paragraph_indexes):
                        """Ділить абзаци наказу на суцільні прогони для копіювання.

                        Розрив прогону робить лише РУЧНИЙ розрив сторінки в наказі:
                        такий абзац у витяг не переноситься, пагінація будується
                        заново. Наявність розриву перевіряємо одним читанням тексту
                        всього діапазону, а поабзацно — тільки якщо він там справді є.
                        """
                        if not paragraph_indexes:
                            return []
                        span_text = ""
                        try:
                            span_text = str(
                                source_doc.Range(
                                    source_paragraphs[paragraph_indexes[0] - 1].Range.Start,
                                    source_paragraphs[paragraph_indexes[-1] - 1].Range.End,
                                ).Text
                                or ""
                            )
                        except Exception:
                            span_text = "\x0c"  # не вдалося прочитати — перевіряємо поабзацно
                        if "\x0c" not in span_text:
                            return [list(paragraph_indexes)]

                        runs, current = [], []
                        for paragraph_index in paragraph_indexes:
                            if has_manual_page_break(source_paragraphs[paragraph_index - 1]):
                                self.log("Пропущено вихідний розрив сторінки; пагінація витягу буде побудована заново.")
                                if current:
                                    runs.append(current)
                                    current = []
                                continue
                            current.append(paragraph_index)
                        if current:
                            runs.append(current)
                        return runs

                    def carry_geometry(source_range, run, start, end):
                        """Переносить геометрію абзаців наказу на вставлений діапазон.

                        Геометрію переносимо ЯВНО, а не покладаємось на FormattedText:
                        Word не записує властивість, яка дорівнює типовій (Alignment=Left,
                        FirstLineIndent=0), і такий абзац у витягу успадковував стиль
                        `Normal` ШАБЛОНА. Якщо в шаблоні стоїть «за шириною» та відступ
                        1.25 см — біографічний блок наказу (left 8 см, перший рядок 0,
                        уліво) з'їжджав і розтягувався по ширині, і так само «плив» §.
                        `ParagraphFormat` діапазону віддає ДІЮЧІ значення наказу.

                        Для прогону з однаковою геометрією вистачає одного
                        `ParagraphFormat` на весь діапазон. Якщо всередині геометрія
                        різна, Word віддає wdUndefined — тоді переносимо поабзацно.
                        """
                        geometry_props = ("Alignment", "LeftIndent", "RightIndent", "FirstLineIndent")

                        def read_geometry(range_object):
                            source_format = range_object.ParagraphFormat
                            values = {}
                            for prop in geometry_props:
                                try:
                                    values[prop] = getattr(source_format, prop)
                                except Exception:
                                    pass
                            return values

                        def write_geometry(range_object, values):
                            dest_format = range_object.ParagraphFormat
                            try:
                                dest_format.PageBreakBefore = False
                            except Exception:
                                pass
                            for prop, value in values.items():
                                try:
                                    setattr(dest_format, prop, value)
                                except Exception:
                                    pass

                        geometry = read_geometry(source_range)
                        if all(value != WD_UNDEFINED for value in geometry.values()):
                            write_geometry(doc.Range(start, end), geometry)
                            return

                        destination_paragraphs = iter_paragraphs(doc.Range(start, end))
                        if len(destination_paragraphs) != len(run):
                            # Кількість абзаців не збіглася (таблиця, зноска):
                            # безпечніше взяти геометрію першого абзацу прогону.
                            write_geometry(
                                doc.Range(start, end),
                                read_geometry(source_paragraphs[run[0] - 1].Range),
                            )
                            return
                        for paragraph_index, destination_paragraph in zip(run, destination_paragraphs):
                            write_geometry(
                                destination_paragraph.Range,
                                read_geometry(source_paragraphs[paragraph_index - 1].Range),
                            )

                    def insert_source_span(start_line, end_line, fallback_text, kind):
                        nonlocal insert_point
                        paragraph_indexes = line_span_to_paragraphs(start_line, end_line)
                        # Абзаци наказу беремо з готового списку: наказ відкрито
                        # лише для читання й не змінюється, тож індекси стабільні,
                        # а кожне `source_doc.Paragraphs(i)` — це прохід по всьому
                        # документу заново.
                        # Обрізаємо порожні абзаци на початку та в кінці діапазону, але зберігаємо внутрішні ентери
                        while paragraph_indexes and is_blank_paragraph(source_paragraphs[paragraph_indexes[0] - 1]):
                            paragraph_indexes.pop(0)
                        while paragraph_indexes and is_blank_paragraph(source_paragraphs[paragraph_indexes[-1] - 1]):
                            paragraph_indexes.pop()

                        first_start = None
                        last_end = None
                        # Копіюємо не по абзацу, а СУЦІЛЬНИМИ прогонами. Одне
                        # звернення до Word коштує кілька мілісекунд, а на абзац
                        # їх виходило близько п'ятнадцяти; до того ж поабзацне
                        # копіювання не відтворює таблиці. Прогін рветься лише
                        # там, де в наказі стоїть ручний розрив сторінки.
                        for run in source_runs(paragraph_indexes):
                            source_range = source_doc.Range(
                                source_paragraphs[run[0] - 1].Range.Start,
                                source_paragraphs[run[-1] - 1].Range.End,
                            )
                            start = insert_point
                            destination_range = doc.Range(start, start)
                            destination_range.FormattedText = source_range.FormattedText
                            insert_point = destination_range.End
                            carry_geometry(source_range, run, start, insert_point)
                            # Порожні білі зображення наказу у витяг не переносимо.
                            pasted_range = doc.Range(start, insert_point)
                            removed_images = remove_blank_images(doc, pasted_range)
                            if removed_images:
                                insert_point = pasted_range.End
                                self.log(f"  Прибрано порожніх білих зображень: {removed_images}.")
                            first_start = start if first_start is None else first_start
                            last_end = insert_point

                        if first_start is None and fallback_text:
                            self.log(f"УВАГА: не знайдено позицію оригіналу для {kind}; вставлено резервний текст без форматування.")
                            start = insert_point
                            destination_range = doc.Range(start, start)
                            compact_text = fallback_text.strip()
                            destination_range.Text = compact_text + "\r"
                            insert_point = start + len(compact_text) + 1
                            first_start, last_end = start, insert_point
                        return doc.Range(first_start, last_end) if first_start is not None else None

                    # Один і той самий шматок наказу не переноситься у витяг
                    # двічі — ні як пункт, ні як шапка. Маршрутизація може
                    # віддати той самий діапазон рядків двічі (наприклад, коли
                    # пункт потрапив у два різні § або коли шапка обчислилась
                    # по рядках самого пункту) — тоді у витягу з'являвся той
                    # самий пункт двома копіями поспіль.
                    inserted_source_spans = set()

                    # Найбільший рядок наказу, який уже перенесено у витяг.
                    # Пункти йдуть у порядку документа, тож усе, що не нижче
                    # цієї межі, у витягу ВЖЕ Є.
                    max_inserted_line = -1

                    def trim_to_new_lines(start_line, end_line, what):
                        """Обрізає діапазон до рядків наказу, яких у витягу ще немає.

                        Інваріант: кожен рядок наказу потрапляє у витяг
                        щонайбільше один раз. Перевірки на ТОЧНИЙ збіг
                        діапазонів для цього не досить — класичний випадок
                        пункт «ВИКЛАСТИ В ТАКІЙ РЕДАКЦІЇ»: усередині нього
                        цитується цілий § з преамбулою, маршрутизація бачить
                        там і «шапку», і окремий пункт, і їхні діапазони
                        ПЕРЕКРИВАЮТЬСЯ зі вступним абзацом. Через це пункт
                        друкувався у витягу двічі.

                        Повертає (початок, кінець, «усе вже перенесено»).
                        """
                        if start_line is None or end_line is None:
                            return start_line, end_line, False
                        if start_line > max_inserted_line:
                            return start_line, end_line, False
                        new_start = max_inserted_line + 1
                        fully_covered = new_start > end_line
                        message = (
                            f"{what}: рядки {start_line}–"
                            f"{min(end_line, max_inserted_line)} наказу вже перенесено вище — "
                            + ("пропущено повністю" if fully_covered
                               else f"переносимо лише {new_start}–{end_line}")
                            + " (щоб не задвоїти текст)."
                        )
                        self.log(f"  ⚠️ {cipher}: {message}")
                        layout_warnings.append(f"{cipher}: {message}")
                        if fully_covered:
                            return None, None, True
                        return new_start, end_line, False

                    def span_already_inserted(start_line, end_line):
                        # Пункт без відомих рядків наказу (вставляється резервним
                        # текстом) дублем НЕ вважається: у таких пунктів «діапазон»
                        # однаковий — (None, None), — і другий та наступні такі
                        # пункти мовчки випадали б із витягу.
                        if start_line is None or end_line is None:
                            return False
                        return (start_line, end_line) in inserted_source_spans

                    def heading_repeats_item(heading_span, item_span):
                        """Чи є «шапка» насправді тим самим текстом, що й пункт.

                        Дотик рівно одним рядком (шапка закінчується там, де
                        починається пункт) — нормальний випадок: цей рядок майже
                        завжди порожній і при вставці однаково обрізається.
                        Дублювання — це коли шапка ЗАХОДИТЬ у пункт.
                        """
                        if None in heading_span or None in item_span:
                            return False
                        return heading_span[1] > item_span[0] or heading_span[0] >= item_span[0]

                    for item in items:
                        item_span = (item.get("source_start_line"), item.get("source_end_line"))
                        item_label = item.get("label", "пункт")

                        # Перевірку на дубль робимо ДО вставки шапок: інакше
                        # шапка вже стоїть у документі, а пункт під нею
                        # пропускається — і § лишається висіти сам, без нічого,
                        # з чим його можна зчепити.
                        if span_already_inserted(item_span[0], item_span[1]):
                            message = (
                                f"{item_label}: цей самий фрагмент наказу (рядки "
                                f"{item_span[0]}–{item_span[1]}) уже перенесено — пропущено як дубль."
                            )
                            self.log(f"  ⚠️ {cipher}: {message}")
                            layout_warnings.append(f"{cipher}: {message}")
                            continue

                        heading_keys = tuple(
                            (heading_range[0], heading_range[1])
                            for heading_range in item.get("heading_ranges", [])
                            if isinstance(heading_range, (list, tuple))
                            and len(heading_range) == 2
                            and all(value is not None for value in heading_range)
                        )
                        if not heading_keys:
                            fallback_key = (item.get("heading_start_line"), item.get("heading_end_line"))
                            if all(value is not None for value in fallback_key):
                                heading_keys = (fallback_key,)

                        # Шапка, що перекривається з власним пунктом, — це не
                        # шапка: інакше текст пункту вставився б двічі поспіль.
                        kept_heading_keys = tuple(
                            heading_key for heading_key in heading_keys
                            if not heading_repeats_item(heading_key, item_span)
                        )
                        if kept_heading_keys != heading_keys:
                            message = (
                                f"{item_label}: шапка заходить у рядки самого пункту "
                                f"(шапка {heading_keys}, пункт {item_span[0]}–{item_span[1]}) — "
                                "шапку не переносимо, щоб не задвоїти пункт."
                            )
                            self.log(f"  ⚠️ {cipher}: {message}")
                            layout_warnings.append(f"{cipher}: {message}")
                        heading_keys = kept_heading_keys

                        new_heading_ranges = []
                        if heading_keys != copied_heading_keys:
                            # Порівнюємо ієрархію рівень-за-рівнем (§ → шапка →
                            # підшапка → ...) і повторюємо лише ті рівні, що
                            # реально змінилися. Спільні рівні (напр. незмінний
                            # § чи шапка одного розділу) НЕ дублюються для
                            # кожного пункту — лише нова/інша частина ієрархії.
                            common_len = 0
                            while (
                                common_len < len(heading_keys)
                                and common_len < len(copied_heading_keys)
                                and heading_keys[common_len] == copied_heading_keys[common_len]
                            ):
                                common_len += 1
                            for heading_key in heading_keys[common_len:]:
                                if span_already_inserted(heading_key[0], heading_key[1]):
                                    continue
                                heading_start, heading_end, heading_covered = trim_to_new_lines(
                                    heading_key[0], heading_key[1], f"шапка перед {item_label}"
                                )
                                if heading_covered:
                                    continue
                                heading_range = insert_source_span(
                                    heading_start, heading_end, item.get("parent_heading", ""), "заголовка"
                                )
                                if heading_range:
                                    inserted_source_spans.add((heading_key[0], heading_key[1]))
                                    if heading_end is not None:
                                        max_inserted_line = max(max_inserted_line, heading_end)
                                    inserted_heading_ranges.append(heading_range)
                                    new_heading_ranges.append(heading_range)
                                    # §, основна шапка та підшапка — окремі
                                    # елементи; після кожного лишається відступ.
                                    insert_empty_paragraph()
                        copied_heading_keys = heading_keys

                        item_start, item_end, item_covered = trim_to_new_lines(
                            item.get("source_start_line"), item.get("source_end_line"), item_label
                        )
                        if item_covered:
                            continue
                        item_range = insert_source_span(
                            item_start, item_end,
                            item.get("original_text") or item.get("text", ""), "пункту"
                        )
                        if item_range:
                            inserted_source_spans.add(item_span)
                            if item_end is not None:
                                max_inserted_line = max(max_inserted_line, item_end)
                            inserted_item_ranges.append(item_range)
                            inserted_item_labels.append(item_label)
                            heading_item_pairs.extend(
                                (heading_range, item_range, item_label)
                                for heading_range in new_heading_ranges
                            )
                            # Між кожним пунктом лишаємо один порожній абзац.
                            # Наприкінці змісту normalize_signature_gap доведе
                            # відступ перед підписантом до двох таких абзаців.
                            #
                            # Це ЄДИНИЙ порожній абзац, який додає генератор.
                            # Усе, що всередині пункту (зокрема порожній абзац
                            # перед рядком «р. н.»), переноситься 1-в-1 з наказу
                            # — у реальних наказах він там завжди є. Своїх
                            # порожніх абзаців усередину пункту не досипаємо:
                            # офіційний зразок (додаток 44) сам непослідовний,
                            # тож оригінал наказу — єдиний надійний еталон.
                            insert_empty_paragraph()

                    signer_start = None
                    content_end = insert_point
                    for paragraph_index in range(1, doc.Paragraphs.Count + 1):
                        paragraph = doc.Paragraphs(paragraph_index).Range
                        if paragraph.Start >= content_end and not is_blank_paragraph(paragraph):
                            signer_start = paragraph.Start
                            break
                    if signer_start is not None:
                        signer_start = normalize_signature_gap(signer_start, content_end)

                    executor_start = executor_start_from_bookmark(executor_bookmark)

                    if inserted_item_ranges:
                        apply_keep_rules(
                            inserted_item_ranges, inserted_heading_ranges, heading_item_pairs, signer_start, executor_start
                        )
                        doc.Repaginate()
                        # Остаточну перевірку макета (і рішення про needs_manual_review)
                        # відкладаємо до ПІСЛЯ підбору міжрядкового інтервалу нижче —
                        # адже стиснення 16→14 пт часто саме й усуває цю розбіжність.

                # Закладку виконавця НЕ видаляємо тут: вона потрібна нижче для
                # точного пошуку абзацу виконавця. Видаляємо після позиціонування.

                clean_redundant_blanks()
                doc.Repaginate()

                pages_count = doc.ComputeStatistics(2)  # wdStatisticPages

                def spacing_bounds():
                    """Межі, у яких ДОЗВОЛЕНО міняти інтервали: лише вставлений зміст
                    та блок підписанта (правило «змінюємо лише текст і підписанта»).

                    Шапка зразка (герб, {{кому}}/{{куди}}, назва, дата/номер) та блок
                    виконавця не чіпаються НІКОЛИ: точний інтервал (wdLineSpaceExactly)
                    обрізає зображення по висоті рядка — саме через це герб ховався
                    під текстом."""
                    if content_start_pos is None:
                        return None, None
                    # Межу беремо із закладки виконавця — це одне звернення до
                    # Word. Пошук абзацу перебором лишаємо фолбеком: ці межі
                    # рахуються на кожне виставлення інтервалу.
                    upper = executor_paragraph_start(executor_bookmark)
                    if upper is None:
                        upper = doc.Content.End
                        exec_idx = find_executor_paragraph_index(executor_bookmark)
                        if exec_idx:
                            try:
                                upper = doc.Paragraphs(exec_idx).Range.Start
                            except Exception:
                                upper = doc.Content.End
                    return content_start_pos, upper

                def apply_exact_line_spacing(points):
                    """Точний міжрядковий інтервал (wdLineSpaceExactly) на весь
                    вставлений зміст і підписанта — ОДНИМ діапазоном.

                    Раніше тут був обхід усіх абзаців документа з ~10 зверненнями
                    до Word на кожен. Оскільки інтервал перебирається, це давало
                    тисячі викликів COM на кожен витяг. Word сам застосовує
                    ParagraphFormat до всіх абзаців діапазону.

                    Порожні абзаци отримують той самий інтервал, що й текст:
                    правило «максимально заповнена сторінка» стосується всієї
                    сторінки, а різна висота порожніх рядків давала різні
                    відступи між пунктами в сусідніх витягах одного наказу.
                    """
                    lower, upper = spacing_bounds()
                    if lower is None or upper is None or upper <= lower:
                        return
                    target = doc.Range(lower, upper)
                    # Абзац із вбудованим зображенням точного інтервалу не
                    # отримує — інакше Word обріже картинку по висоті рядка.
                    # Запам'ятовуємо їхні інтервали й повертаємо після заміни.
                    protected = []
                    try:
                        for shape in target.InlineShapes:
                            shape_format = shape.Range.Paragraphs(1).Range.ParagraphFormat
                            protected.append(
                                (shape_format, shape_format.LineSpacingRule, shape_format.LineSpacing)
                            )
                    except Exception:
                        protected = []
                    target_format = target.ParagraphFormat
                    target_format.LineSpacingRule = 4  # wdLineSpaceExactly
                    target_format.LineSpacing = points
                    target_format.SpaceBefore = 0
                    target_format.SpaceAfter = 0
                    for shape_format, shape_rule, shape_spacing in protected:
                        try:
                            shape_format.LineSpacingRule = shape_rule
                            shape_format.LineSpacing = shape_spacing
                        except Exception:
                            pass

                # ПІДБІР МІЖРЯДКОВОГО ІНТЕРВАЛУ — один пошук замість трьох.
                #
                # Було три послідовні перебори: 16→14 з кроком 0.5 «щоб влізло на
                # сторінку», потім ще один такий самий для багатосторінкових, потім
                # 16→14 з кроком 0.25 «щоб сторінка була заповнена». Вони міряли
                # ті самі значення й разом давали до 14 репагінацій на витяг.
                #
                # Обидві ознаки монотонні: що менший інтервал, то менше сторінок і
                # то більше пунктів вміщується на першу. Тому мета одна — НАЙБІЛЬШИЙ
                # інтервал, що дає мінімальну кількість сторінок і максимум пунктів
                # на першій сторінці. Для витягу на одну сторінку це і є «максимально
                # заповнена сторінка», для багатосторінкового — колишній «бонус».
                if content_start_pos is not None:
                    spacing_max, spacing_min, spacing_step = 16.0, 14.0, 0.25
                    spacing_grid = [
                        round(spacing_max - index * spacing_step, 2)
                        for index in range(int(round((spacing_max - spacing_min) / spacing_step)) + 1)
                    ]
                    applied_spacing = None
                    spacing_measurements: dict[float, tuple[int, int]] = {}

                    def items_fitting_first_page():
                        """Скільки пунктів ПОВНІСТЮ вміщується на першу сторінку.

                        Пункти йдуть у порядку документа, тож межу шукаємо двійково:
                        інакше на кожне вимірювання інтервалу припадало по два
                        звернення `Information` на КОЖЕН пункт.
                        """
                        if not inserted_item_ranges:
                            return 0
                        low, high = 0, len(inserted_item_ranges)
                        while low < high:
                            middle = (low + high) // 2
                            start_page, end_page = range_pages(inserted_item_ranges[middle])
                            if start_page == 1 and end_page == 1:
                                low = middle + 1
                            else:
                                high = middle
                        return low

                    def measure_spacing(points):
                        """(сторінок, пунктів на першій) для інтервалу; стан документа
                        після виклику завжди відповідає виміряному значенню."""
                        nonlocal applied_spacing
                        if applied_spacing != points:
                            apply_exact_line_spacing(points)
                            doc.Repaginate()
                            applied_spacing = points
                        if points not in spacing_measurements:
                            pages = doc.ComputeStatistics(2)  # wdStatisticPages
                            # На одній сторінці лічити нема чого: там усі пункти.
                            # Кожне звернення до номера сторінки коштує дорого.
                            spacing_measurements[points] = (
                                pages,
                                len(inserted_item_ranges) if pages <= 1 else items_fitting_first_page(),
                            )
                        return spacing_measurements[points]

                    best_possible = measure_spacing(spacing_min)
                    chosen_spacing = spacing_min
                    if measure_spacing(spacing_max) == best_possible:
                        # Стискати нема сенсу: найбільший інтервал дає той самий
                        # результат, що й найменший.
                        chosen_spacing = spacing_max
                    else:
                        low, high = 0, len(spacing_grid) - 1  # high (14.0) свідомо годиться
                        while low < high:
                            middle = (low + high) // 2
                            if measure_spacing(spacing_grid[middle]) == best_possible:
                                high = middle
                            else:
                                low = middle + 1
                        chosen_spacing = spacing_grid[low]
                        if measure_spacing(chosen_spacing) != best_possible:
                            # Підстраховка на випадок, коли верстка Word виявилась
                            # немонотонною: повертаємось до гарантованого значення.
                            chosen_spacing = spacing_min
                            measure_spacing(chosen_spacing)

                    pages_count, items_on_first = spacing_measurements[chosen_spacing]
                    if chosen_spacing != spacing_max:
                        self.log(
                            f"  ℹ️ {cipher}: точний інтервал {chosen_spacing} пт — "
                            f"сторінок: {pages_count}, пунктів на першій: {items_on_first}."
                        )

                # 2. Остаточна перевірка макета — ПІСЛЯ підбору міжрядкового
                # інтервалу (16→14 пт), який часто сам усуває розбіжність
                # (відірвану шапку, перебір на 2 сторінки тощо). Лише якщо
                # проблема лишається й після цього — витяг справді потребує
                # ручної перевірки, і виконавця НЕ підштовхуємо штучно.
                #
                # Витяг на ОДНІЙ сторінці не перевіряємо взагалі: жоден пункт не
                # може її не вміститись, ніщо не може відірватись від шапки, і
                # підписант завжди там само. А кожна така перевірка — це два
                # запити номера сторінки на кожен пункт і кожну шапку.
                if inserted_item_ranges and pages_count > 1:
                    remaining_issues = layout_issues(
                        inserted_item_ranges, inserted_item_labels, heading_item_pairs, signer_start, executor_start
                    )
                    for issue in remaining_issues:
                        layout_warnings.append(f"{cipher}: {issue}.")
                    if remaining_issues:
                        extract_needs_manual_review = True

                # 3. Виконавець ЗАВЖДИ вирівнюється строго до низу поточної сторінки,
                # окрім випадків, де це вимагає ручних змін (див. position_executor_at_page_bottom)
                position_executor_at_page_bottom(executor_bookmark, signer_start, extract_needs_manual_review)
                doc.Repaginate()
                pages_count = doc.ComputeStatistics(2)

                # Службову закладку виконавця прибираємо лише після позиціонування.
                if executor_bookmark and doc.Bookmarks.Exists(executor_bookmark):
                    try:
                        doc.Bookmarks(executor_bookmark).Delete()
                    except Exception:
                        pass

                doc.Save()
                doc.Close(False)
                (management_temp_files if is_management else temp_files).append(
                    (temp_path, pages_count, cipher)
                )
                self.log(
                    f"  ✓ {cipher}: {pages_count} стор., "
                    f"{time.monotonic() - extract_started_at:.1f} с."
                )

            source_doc.Close(False)

            # Витяги до управління збираються окремо: без вирівнювання під
            # друк «2 на 1» — їх не друкують пакетом і нікому не надсилають.
            if management_temp_files:
                self.log(
                    f"\nЗбираємо {len(management_temp_files)} витяг(ів) до управління "
                    "в окремий документ..."
                )
                management_doc = word.Documents.Open(os.path.abspath(management_temp_files[0][0]))
                for next_path, _next_pages, _next_name in management_temp_files[1:]:
                    tail_range = management_doc.Content
                    tail_range.Collapse(0)
                    tail_range.InsertBreak(2)  # wdSectionBreakNextPage
                    tail_range = management_doc.Content
                    tail_range.Collapse(0)
                    tail_range.InsertFile(os.path.abspath(next_path))
                tail_paragraph = last_paragraph(management_doc)
                if not tail_paragraph.Range.Text.strip():
                    tail_paragraph.Range.Delete()
                management_doc.SaveAs2(os.path.abspath(management_out_file), 16)
                management_doc.Close(False)
                self.log(f"Збережено витяги до управління: {management_out_file}")

            if temp_files:
                # Збираємо всі витяги в один фінальний документ
                self.progress_step(len(temp_files), "складання документа")
                self.log(f"\nЗбираємо {len(temp_files)} витягів в один документ...")
                enable_2up = self.duplex_2up_layout.get()

                first_path, first_pages, first_cipher = temp_files[0]
                if len(temp_files) == 1:
                    if enable_2up and first_pages > 1 and (first_pages % 2 != 0):
                        target_doc = word.Documents.Open(os.path.abspath(first_path))
                        rng = target_doc.Content
                        rng.Collapse(0)
                        rng.InsertBreak(2)  # wdSectionBreakNextPage
                        target_doc.SaveAs2(os.path.abspath(out_file), 16)
                        target_doc.Close(False)
                        self.log(f"Додано порожню сторінку для вирівнювання витягу {first_cipher} ({first_pages} стор.) під друк 2 на 1.")
                    else:
                        # Через SaveAs2, а не копіюванням: тимчасовий файл може мати
                        # формат шаблону (Word 97-2003), а результат завжди .docx.
                        target_doc = word.Documents.Open(os.path.abspath(first_path))
                        target_doc.SaveAs2(os.path.abspath(out_file), 16)
                        target_doc.Close(False)
                else:
                    target_doc = word.Documents.Open(os.path.abspath(first_path))

                    def sheet_pages() -> int:
                        """ФАКТИЧНА кількість сторінок зібраного документа.

                        Вимір потрібен лише для друку «2 на 1»: там від парності
                        залежить, на яку половину аркуша сяде наступний витяг. Без
                        цього режиму репагінувати документ, що росте, після кожної
                        вставки — це квадратична робота ні для чого.
                        """
                        if not enable_2up:
                            return 0
                        target_doc.Repaginate()
                        return target_doc.ComputeStatistics(2)  # wdStatisticPages

                    def add_blank_page() -> None:
                        rng_blank = target_doc.Content
                        rng_blank.Collapse(0)
                        rng_blank.InsertBreak(2)  # wdSectionBreakNextPage

                    def strip_trailing_blank_paragraphs() -> None:
                        """Прибирає «висячі» порожні абзаци в кінці документа.

                        Без цього вимірювання рахувало ФАНТОМНУ останню сторінку,
                        якої в друці немає: парність виходила невірна, і наступний
                        витяг сідав на праву половину того самого аркуша, де
                        закінчувався попередній. Викликати ЛИШЕ після вставки
                        витягу — навмисні порожні сторінки чіпати не можна.
                        """
                        try:
                            while target_doc.Paragraphs.Count > 1:
                                paragraph = last_paragraph(target_doc).Range
                                if (paragraph.Text or "").strip(chr(13) + chr(7) + chr(11) + chr(12) + chr(32) + chr(9)):
                                    break
                                # Той самий захист, що й у clean_redundant_blanks:
                                # знак абзацу, який Word видалити не може, мовчки
                                # лишається на місці й зациклює обрізання хвоста.
                                length_before = target_doc.Content.End
                                paragraph.Delete()
                                if target_doc.Content.End >= length_before:
                                    break
                        except Exception:
                            pass

                    # Сторінки ЗАВЖДИ міряємо, а не рахуємо додаванням: розрив
                    # розділу не гарантовано додає рівно одну сторінку, тож
                    # лічильник «повзе». Через це багатосторінковий витяг міг
                    # опинитися на ПРАВІЙ половині аркуша, а кінець попереднього
                    # витягу — ділити аркуш із початком наступного.
                    strip_trailing_blank_paragraphs()
                    current_doc_pages = sheet_pages()

                    if enable_2up and current_doc_pages > 1 and (current_doc_pages % 2 != 0):
                        add_blank_page()
                        current_doc_pages = sheet_pages()
                        self.log(
                            f"Додано порожню сторінку після витягу {first_cipher} "
                            f"({current_doc_pages - 1} стор.) для вирівнювання аркуша."
                        )

                    for i in range(1, len(temp_files)):
                        t_path, t_pages, t_cipher = temp_files[i]

                        # Багатосторінковий витяг починається з ЛІВОЇ половини
                        # аркуша, тобто з НЕПАРНОЇ логічної сторінки.
                        if enable_2up and t_pages > 1 and (current_doc_pages % 2 != 0):
                            add_blank_page()
                            current_doc_pages = sheet_pages()
                            self.log(
                                f"Додано порожню сторінку перед багатосторінковим витягом "
                                f"{t_cipher} ({t_pages} стор.), щоб він почався з нового аркуша."
                            )

                        pages_before = current_doc_pages

                        rng = target_doc.Content
                        rng.Collapse(0)
                        rng.InsertBreak(2)  # wdSectionBreakNextPage

                        rng = target_doc.Content
                        rng.Collapse(0)
                        rng.InsertFile(os.path.abspath(t_path))
                        strip_trailing_blank_paragraphs()
                        current_doc_pages = sheet_pages()

                        # Скільки сторінок витяг займає САМЕ В ЗІБРАНОМУ документі.
                        # Окремим файлом він міг мати іншу кількість, і саме довіра
                        # до тієї, старої, ламала вирівнювання. Без режиму «2 на 1»
                        # зібраний документ не міряємо, тож беремо кількість файлу.
                        actual_pages = max(1, current_doc_pages - pages_before) if enable_2up else t_pages
                        if enable_2up and actual_pages != t_pages:
                            self.log(
                                f"УВАГА: витяг {t_cipher} у зібраному документі займає "
                                f"{actual_pages} стор. замість {t_pages}; вирівнювання "
                                "рахується за фактичною кількістю."
                            )

                        # Після багатосторінкового витягу наступний має починатися з
                        # НОВОГО аркуша: якщо документ закінчився на лівій половині
                        # (непарна сторінка) — доповнюємо порожньою.
                        if enable_2up and actual_pages > 1 and (current_doc_pages % 2 != 0):
                            add_blank_page()
                            current_doc_pages = sheet_pages()
                            self.log(
                                f"Додано порожню сторінку після багатосторінкового витягу "
                                f"{t_cipher} ({actual_pages} стор.) для вирівнювання наступного аркуша."
                            )

                    last_para = last_paragraph(target_doc)
                    if last_para.Range.Text.strip() == "":
                        last_para.Range.Delete()

                    target_doc.SaveAs2(os.path.abspath(out_file), 16)
                    target_doc.Close(False)

                # Точний підрахунок сторінок та фізичних аркушів паперу для друку
                check_doc = word.Documents.Open(os.path.abspath(out_file), ReadOnly=True)
                check_doc.Repaginate()
                final_total_pages = check_doc.ComputeStatistics(2)
                check_doc.Close(False)

                total_extracts = len(temp_files)
                single_page_cnt = sum(1 for tf in temp_files if tf[1] == 1)
                multi_page_cnt = sum(1 for tf in temp_files if tf[1] > 1)

                if enable_2up:
                    sheets_needed = (final_total_pages + 1) // 2
                    stats_msg = (
                        f"📋 Сформовано витягів: {total_extracts} шт. ({single_page_cnt} односторінк. та {multi_page_cnt} багатосторінк.)\n"
                        f"📄 Загальна кількість сторінок у документі: {final_total_pages} стор.\n"
                        f"🖨️ Кількість фізичних аркушів паперу для друку («2 сторінки на 1 аркуш»): {sheets_needed} арк. А4"
                    )
                else:
                    sheets_1side = final_total_pages
                    sheets_2side = (final_total_pages + 1) // 2
                    stats_msg = (
                        f"📋 Сформовано витягів: {total_extracts} шт. ({single_page_cnt} односторінк. та {multi_page_cnt} багатосторінк.)\n"
                        f"📄 Загальна кількість сторінок у документі: {final_total_pages} стор.\n"
                        f"🖨️ Кількість фізичних аркушів паперу для друку:\n"
                        f"   • Односторонній друк (1 на 1): {sheets_1side} арк. А4\n"
                        f"   • Двосторонній друк (Duplex): {sheets_2side} арк. А4"
                    )

                # Кількість пропущених пунктів показується ЗАВЖДИ, зокрема «0».
                # Інакше пункт, який не отримав адресата, лишався помітним лише
                # в журналі, і його легко було не побачити серед статистики друку.
                missed_count = len(unmatched_items)
                excluded_count = len(map_res.get("skipped_items", []))
                missed_lines = [
                    f"{'⚠️' if missed_count else '✅'} Пунктів без адресата (пропущено): {missed_count}"
                ]
                if missed_count:
                    missed_lines.append(f"   Перелік: {os.path.basename(missing_report)}")
                if excluded_count:
                    missed_lines.append(
                        f"ℹ️ Виключено із загального переліку (управління): {excluded_count}"
                    )
                if management_temp_files:
                    missed_lines.append(
                        f"📁 Витяги до управління — окремий файл на {len(management_temp_files)} шт.: "
                        f"{os.path.basename(management_out_file)}"
                    )
                stats_msg = chr(10).join([stats_msg] + missed_lines)

                batch_seconds = time.monotonic() - batch_started_at
                self.progress_end(f"Витягів: {total_extracts}")
                self.log(f"Збережено файл: {out_file}")
                self.log(
                    f"⏱️ Час генерації: {batch_seconds:.1f} с "
                    f"({batch_seconds / max(1, total_extracts):.1f} с на витяг)."
                )
                self.log(f"\n📊 СТАТИСТИКА ТА РОЗРАХУНОК ДРУКУ:\n{stats_msg}\n")
                self.show_layout_warnings(layout_warnings)

                if layout_warnings:
                    self.log("\nУВАГА: потрібна ручна перевірка макета:")
                    for warning in layout_warnings:
                        self.log(f"• {warning}")
                    messagebox.showwarning(
                        "Перевірте макет витягів",
                        f"Сформовано {total_extracts} витягів, але деякі витяги потребують ручного коригування.\n\n"
                        f"{stats_msg}\n\nДеталі у журналі.",
                    )
                else:
                    messagebox.showinfo(
                        "Успіх",
                        f"Успішно сформовано {total_extracts} витягів!\n\n"
                        f"{stats_msg}\n\n"
                        f"Файл збережено:\n{os.path.basename(out_file)}"
                    )
            else:
                # У наказі лише зміни по управлінню: загального файла немає,
                # і це не помилка — розрахунок розсилки для них не ведеться.
                total_extracts = 0
                self.progress_end(f"До управління: {len(management_temp_files)}")
                self.log("Загальний файл адресних витягів у цьому запуску не створювався.")
                messagebox.showinfo(
                    "Успіх",
                    f"Сформовано витягів до управління: {len(management_temp_files)}.\n\n"
                    "Загальний файл адресних витягів у цьому запуску не створювався.\n\n"
                    f"Файл збережено:\n{os.path.basename(management_out_file)}",
                )

            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir, ignore_errors=True)

            # Налаштування повертаємо ДО показу вікна: з вимкненим ScreenUpdating
            # видимий Word не перемальовується й виглядає зависшим.
            restore_word_settings(word, word_settings)
            word_settings = None

            self.log("Відкриваємо згенерований документ...")
            # Якщо загального файла немає (у наказі лише зміни по управлінню),
            # показуємо той єдиний документ, який справді створено.
            final_doc = word.Documents.Open(
                os.path.abspath(out_file if temp_files else management_out_file)
            )
            word.Visible = True
            final_doc.Activate()
            word = None

        finally:
            self.progress_end()
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir, ignore_errors=True)
            if word:
                restore_word_settings(word, word_settings)
                force_quit_word(word)

    # =========================================================================
    # ДІЇ: ГЕНЕРАЦІЯ ПРИМІРНИКІВ 2/3
    # =========================================================================
    def run_generate_copies(self):
        self.save_config()
        back_page = self.p2_back_page_path.get()
        if not back_page or not os.path.exists(back_page):
            messagebox.showwarning("Помилка", "Виберіть дійсний файл шаблону «Задня сторінка»!")
            return

        # Беремо рівно те, що видно в списку на вкладці: і тека, і окремо
        # обрані файли проходять той самий шлях, тож пакет завжди збігається
        # з тим, що бачить користувач.
        order_files = [
            path for path in self._selected_p2_order_paths()
            if os.path.normcase(os.path.abspath(path)) != os.path.normcase(os.path.abspath(back_page))
        ]
        if not order_files:
            messagebox.showwarning("Немає файлів", "Не знайдено файлів наказів DOCX для обробки!")
            return

        copy_title = "Примірник № 2"
        back_page_abs = os.path.abspath(back_page)

        out_root = self.p2_out_folder.get()
        if not out_root or not self.p2_out_folder_manual.get():
            # Поки папку не обрано вручну, вона завжди відповідає ПОТОЧНОМУ
            # джерелу: інакше примірники писались би в папку попереднього наказу.
            base_dir = (
                self.p2_orders_folder.get()
                if self.p2_source_mode.get() == "folder"
                else os.path.dirname(order_files[0])
            )
            out_root = os.path.join(base_dir, "Примірники_2")
            self.p2_out_folder.set(out_root)
        os.makedirs(out_root, exist_ok=True)
        self.save_config()

        self.btn_run_p2.config(state=DISABLED)
        self.p2_log_text.delete(1.0, tk.END)
        for item in self.p2_tree.get_children():
            self.p2_tree.delete(item)

        self.log_p2(f"Початок пакетного формування примірників № 2: {len(order_files)} наказ(ів)...")

        import traceback

        created_records = []
        failed_orders: list[tuple[str, str]] = []
        # У режимі превʼю Word показується — інакше дивитись нема на що.
        preview_on = bool(self.p2_preview.get())
        preview_delay = self.p2_preview_delay.get()
        if preview_on:
            self.log_p2(
                f"🐢 Режим превʼю увімкнено: Word буде видимим, пауза після "
                f"кожного кроку — {preview_delay} сек. Для великого пакета це довго."
            )

        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = preview_on
        word.DisplayAlerts = 0  # wdAlertsNone: не показувати блокуючі діалоги (напр. "Зберегти зміни?")

        try:
            # Заготовка примірника може бути багатосторінковою.
            test_tmpl = word.Documents.Open(back_page_abs, ReadOnly=True)
            try:
                test_tmpl.Repaginate()
                tmpl_pages = test_tmpl.ComputeStatistics(2)
                template_text = test_tmpl.Content.Text
            finally:
                test_tmpl.Close(False)

            # Заготовка більше не є «останньою сторінкою», яку підставляють у
            # наказ: тепер це повноцінний шаблон примірника (шапка + теги +
            # {{зміст}}), тому обмеження в 1 сторінку зняте.
            if tmpl_pages != 1:
                self.log_p2(f"  ℹ️ Заготовка займає {tmpl_pages} стор.")

            if "{{зміст}}" not in template_text:
                self.log_p2("❌ ПОМИЛКА: у заготовці примірника не знайдено тег {{зміст}}!")
                messagebox.showerror(
                    "Помилка шаблону",
                    "У заготовці примірника не знайдено тег {{зміст}}.\n\n"
                    f"Файл: {os.path.basename(back_page_abs)}\n\n"
                    "Саме в це місце вставляються пункти наказу разом із підписантом.",
                )
                return

            # Якщо у заготовці є окремий тег підписанта — блок підписанта йде
            # в нього, а зміст завершується перед ним. Інакше підписант
            # лишається частиною змісту (сумісність зі старими заготовками).
            signer_tag_in_template = any(tag in template_text.casefold() for tag in SIGNER_TAGS)
            self.log_p2(
                "  Підписант: окремий тег {{підписант}}."
                if signer_tag_in_template
                else "  Підписант: у складі {{зміст}} (тегу {{підписант}} у заготовці немає)."
            )

            required_tags = tuple(back_page_tag_values("номер", "01.01.2000"))
            missing_tags = [tag for tag in required_tags if tag not in template_text]
            if missing_tags:
                self.log_p2(
                    "УВАГА: у шаблоні не знайдено тегів: " + ", ".join(missing_tags) + ". "
                    "Їхні значення не буде підставлено автоматично."
                )

            for idx, order_path in enumerate(order_files, start=1):
                fname = os.path.basename(order_path)
                self.log_p2(f"\n[{idx}/{len(order_files)}] Обробка: {fname}")
                order_num, order_date = extract_metadata_from_filename(fname)

                # Номер наказу може містити «/» (напр. «б/н», «123/45»), який
                # ОС трактує як роздільник шляху й створює вкладені папки,
                # тому назву папки треба чистити так само, як назву файлу.
                sub_folder_name = (
                    f"Наказ № {sanitize_filename(order_num)}"
                    if order_num
                    else sanitize_filename(os.path.splitext(fname)[0])
                )
                target_dir = os.path.join(out_root, sub_folder_name)
                os.makedirs(target_dir, exist_ok=True)

                out_copy_name = build_copy_two_filename(order_num, order_date, fname)
                target_file = os.path.join(target_dir, out_copy_name)

                doc = None
                working_copy = ""
                try:
                    self.log_p2("  Режим: збірка примірника із заготовки (без колонтитулів).")
                    values = back_page_tag_values(order_num, order_date)
                    # Зворотна сумісність для вже створених шаблонів.
                    if order_num:
                        values["{{номер}}"] = order_num
                    if order_date:
                        values["{{дата}}"] = order_date
                    values["{{примірник_номер}}"] = copy_title
                    values["{{засвідчення}}"] = "Згідно з оригіналом"

                    # Виконавець примірників; якщо не заповнено — беремо
                    # виконавця витягів, щоб не змушувати вводити двічі.
                    executor_value = (
                        self.p2_executor.get().strip() or self.executor.get().strip()
                    )
                    if executor_value:
                        values["{{виконавець}}"] = _slash_to_lines(executor_value)

                    # Засвідчувач примірників — окремі поля. За потреби
                    # користувач переносить значення з витягів окремою кнопкою.
                    cert_pos = _slash_to_lines(self.p2_certifier_position.get().strip())
                    cert_rank = self.p2_certifier_rank.get().strip()
                    cert_name = self.p2_certifier_name.get().strip()
                    values.update(certifier_tags(cert_pos, cert_rank, cert_name))
                    if cert_pos:
                        values["{{засвідчувач_посада}}"] = cert_pos
                        values["{{згідно_з_оригіналом_посада}}"] = cert_pos
                    if cert_rank:
                        values["{{засвідчувач_звання}}"] = cert_rank
                        values["{{згідно_з_оригіналом_звання}}"] = cert_rank
                    if cert_name:
                        values["{{засвідчувач_піб}}"] = cert_name
                        values["{{згідно_з_оригіналом_піб}}"] = cert_name
                        values["{{засвідчувач}}"] = cert_name

                    # Тег «Згідно з оригіналом» розгортається в СУЦІЛЬНИЙ блок
                    # без порожніх абзаців: заголовок, посада, а останнім
                    # рядком — звання та прізвище.
                    certifier_block = ["Згідно з оригіналом"]
                    if cert_pos:
                        certifier_block.append(cert_pos)
                    signature_line = " ".join(part for part in (cert_rank, cert_name) if part)
                    if signature_line:
                        certifier_block.append(signature_line)
                    values["{{згідно_з_оригіналом}}"] = "\r".join(certifier_block)
                    values["{{засвідчення}}"] = values["{{згідно_з_оригіналом}}"]

                    # Заготовка може бути у форматі Word 97-2003, тож робоча
                    # копія зберігає реальне розширення, а .docx дає SaveAs2.
                    working_copy = copy_template_for_editing(
                        back_page_abs,
                        os.path.join(target_dir, "_nat_tmpl.docx"),
                        label="заготовка примірника",
                    )
                    # Під кінець великого пакета Word може тимчасово відхиляти
                    # виклики — такий збій не є помилкою даних, тому повторюємо.
                    # Показник кроків СВІЙ на кожен наказ: інакше нумерація
                    # «Крок N/8» тривала б наскрізно через увесь пакет.
                    steps = PreviewSteps(
                        log=self.log_p2,
                        delay=preview_delay,
                        enabled=preview_on,
                        sleeper=self._preview_pause,
                    )
                    final_pages = retry_on_busy_word(
                        lambda: build_copy_document(
                            word,
                            os.path.abspath(order_path),
                            os.path.abspath(working_copy),
                            os.path.abspath(target_file),
                            values,
                            resolve_span=lambda source: self._order_body_context(
                                source, signer_as_tag=signer_tag_in_template
                            ),
                            log=self.log_p2,
                            preview=steps,
                        ),
                        log=self.log_p2,
                    )

                    sheets_1_copy = (final_pages + 1) // 2
                    self.log_p2(
                        f"  ✅ Створено: {out_copy_name} (сторінок: {final_pages} | "
                        f"аркушів для двостороннього друку: {sheets_1_copy} арк.)"
                    )
                    created_records.append((
                        idx,
                        out_copy_name,
                        order_num or "—",
                        order_date or "—",
                        final_pages,
                        sheets_1_copy,
                        target_file,
                    ))
                except Exception as order_error:
                    # Збій одного наказу не має зривати весь пакет.
                    traceback.print_exc()
                    explanation = explain_error(order_error)
                    self.log_p2(f"  ПОМИЛКА під час обробки «{fname}»:\n    "
                                + explanation.replace("\n", "\n    "))
                    failed_orders.append((fname, explanation))
                    try:
                        if doc is not None:
                            doc.Close(False)
                    except Exception:
                        pass
                finally:
                    if working_copy and os.path.exists(working_copy):
                        try:
                            os.remove(working_copy)
                        except OSError:
                            pass

            for rec in created_records:
                self.p2_tree.insert("", tk.END, values=rec)
            self._set_copy_two_sources([record[-1] for record in created_records])

            total_orders = len(created_records)
            total_pages = sum(r[4] for r in created_records if isinstance(r[4], int))
            total_sheets_1_copy = sum(r[5] for r in created_records if isinstance(r[5], int))

            p2_stats_msg = (
                f"📑 Кількість опрацьованих наказів: {total_orders} шт.\n"
                f"📄 Сторінок в 1 примірнику (сумарно): {total_pages} стор.\n"
                f"🖨️ Кількість фізичних аркушів паперу А4 (двосторонній друк, задня сторінка на звороті):\n"
                f"   • На сформовані примірники № 2: {total_sheets_1_copy} арк. А4"
            )

            self.log_p2(f"\n📊 СТАТИСТИКА ТА РОЗРАХУНОК ДРУКУ:\n{p2_stats_msg}")
            self.log_p2(f"\n🎉 Завершено! Успішно сформовано {total_orders} примірник(ів) № 2.")

            if failed_orders:
                details = "\n\n".join(
                    f"Наказ «{name}» — примірник не створено.\n{error}"
                    for name, error in failed_orders[:10]
                )
                if len(failed_orders) > 10:
                    details += (
                        f"\n\n… і ще {len(failed_orders) - 10} наказ(ів) "
                        "з такими самими збоями."
                    )
                self.log_p2(f"\nНе вдалося сформувати примірники: {len(failed_orders)} шт.")
                messagebox.showwarning(
                    "Примірники сформовано частково",
                    f"Сформовано {total_orders} з {len(order_files)} примірник(ів).\n\n"
                    f"Не вдалося ({len(failed_orders)}):\n{details}\n\n{p2_stats_msg}",
                )
            else:
                messagebox.showinfo(
                    "Успіх",
                    f"Успішно сформовано {total_orders} примірник(ів) № 2!\n\n"
                    f"{p2_stats_msg}"
                )

        except Exception as error:
            # Раніше винятку не було де перехопити: таблиця лишалась порожньою,
            # а користувач бачив лише traceback у консолі.
            traceback.print_exc()
            explanation = explain_error(error)
            self.log_p2(f"\n❌ ПОМИЛКА пакетної генерації:\n{explanation}")
            for rec in created_records:
                self.p2_tree.insert("", tk.END, values=rec)
            if created_records:
                self._set_copy_two_sources([record[-1] for record in created_records])
            messagebox.showerror(
                "Примірники сформовано не до кінця",
                f"{explanation}\n\n"
                f"Встигли сформувати: {len(created_records)} примірник(ів).",
            )
        finally:
            if word:
                force_quit_word(word)
            self.btn_run_p2.config(state=NORMAL)

    def select_all_copies(self):
        for item in self.p2_tree.get_children():
            self.p2_tree.selection_add(item)

    def deselect_all_copies(self):
        for item in self.p2_tree.selection():
            self.p2_tree.selection_remove(item)

    def transfer_selected_copy_to_extracts(self):
        selected = self.p2_tree.selection()
        if not selected:
            messagebox.showinfo("Вибір", "Виберіть сформований примірник у таблиці!")
            return

        valid_paths = []
        for sel_id in selected:
            item_vals = self.p2_tree.item(sel_id, "values")
            if item_vals and len(item_vals) >= 2:
                target_path = item_vals[-1]
                if os.path.exists(target_path):
                    valid_paths.append(target_path)

        if not valid_paths:
            messagebox.showwarning("Помилка", "Жоден із вибраних файлів не знайдено на диску.")
            return

        self._set_copy_two_sources(valid_paths)
        target_path = valid_paths[0]
        self._set_processing_order(target_path)
        self._refresh_order_signer()
        self.save_config()
        self.notebook.select(self.tab_extracts_page)
        self.log(f"\n📥 Передано {len(valid_paths)} примірників № 2 для розрахунку та витягів.")
        messagebox.showinfo(
            "Передано до Витягів",
            f"Позначено {len(valid_paths)} примірників № 2. У вкладці 1 можна зняти позначки з непотрібних файлів."
        )

    def open_p2_output_folder(self):
        folder = self.p2_out_folder.get()
        if folder and os.path.exists(folder):
            os.startfile(folder)
        else:
            messagebox.showwarning("Помилка", "Папка результату ще не створена.")

    def open_selected_copy_in_word(self):
        selected = self.p2_tree.selection()
        if not selected:
            messagebox.showinfo("Вибір", "Виберіть файл у таблиці!")
            return
        item_vals = self.p2_tree.item(selected[0], "values")
        if item_vals and len(item_vals) >= 2:
            file_path = item_vals[-1]
            if os.path.exists(file_path):
                os.startfile(file_path)
            else:
                messagebox.showwarning("Помилка", f"Файл не знайдено: {file_path}")


if __name__ == "__main__":
    # Tk-версії більше немає: запуск цього файлу відкриває Qt-оболонку.
    from generate_extracts_qt import main

    raise SystemExit(main())
