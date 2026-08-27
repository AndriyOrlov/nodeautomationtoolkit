import os
import sys
import json
import re
import shutil
import time
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

# ttkbootstrap для сучасного та привабливого інтерфейсу
try:
    import ttkbootstrap as tb
    from ttkbootstrap.constants import *
    HAS_TTKBOOTSTRAP = True
except ImportError:
    HAS_TTKBOOTSTRAP = False
    import tkinter.ttk as tb
    from tkinter.constants import *

# windnd для drag-and-drop перетягування файлів на вікно
try:
    import windnd
    HAS_WINDND = True
except ImportError:
    HAS_WINDND = False

import win32com.client
import openpyxl

# Додаємо src до sys.path для імпорту модулів nodeautomationtoolkit
project_root = os.path.dirname(os.path.abspath(__file__))
src_path = os.path.join(project_root, "src")
if src_path not in sys.path:
    sys.path.insert(0, src_path)

from nodeautomationtoolkit.builtin_nodes.recipient_mapping import (
    read_recipient_mapping,
    map_military_units,
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
from nodeautomationtoolkit.builtin_nodes.compare_window import DocxCompareWindow


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


def build_extracts_filename(order_num: str, order_date: str) -> str:
    """Назва зведеного файла витягів без вигаданих реквізитів наказу."""
    parts = ["Витяги наказу"]
    if order_num:
        safe_number = sanitize_filename(order_num)
        parts.append(f"№ {safe_number}")
    if order_date:
        safe_date = sanitize_filename(order_date)
        parts.append(f"від {safe_date}")
    return " ".join(parts) + ".docx"


def read_document_text(doc) -> str:
    """Текст документа, зібраний З АБЗАЦІВ, а не з `Content.Text`.

    `Content.Text` склеює цілий рядок таблиці в один рядок тексту, тоді як
    `doc.Paragraphs` рахує кожну комірку окремим абзацом. Через це нумерація
    рядків розходилася з нумерацією абзаців, і після будь-якої таблиці в тілі
    наказу пункти зіставлялися не з тими абзацами — частина пунктів губилася.

    Збираючи текст саме з абзаців, ми отримуємо відповідність «рядок ↔ абзац»
    за побудовою: обидві сторони розбиваються однаково.
    """
    return "\n".join(
        (doc.Paragraphs(index).Range.Text or "").rstrip("\r\x07")
        for index in range(1, doc.Paragraphs.Count + 1)
    )


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


def apply_ukrainian_typography(text: str) -> str:
    """Замінює пробіли після коротких прийменників, сполучників та скорочень на нерозривні (\u00A0)."""
    result = text or ""
    # Код ВОС має лишатися суцільним: «ВОС - 0602002» не можна розривати між
    # рядками, інакше номер зависає під самим краєм сторінки.
    result = re.sub(
        r"(?i)\b(ВОС)\s*(-|–|—)\s*(\d+)",
        lambda m: f"{m.group(1)} {m.group(2)} {m.group(3)}",
        result,
    )
    pattern = r"(?i)\b(з|із|зі|та|до|в|у|на|і|й|по|за|від|при|під|над|про|для|без|через|шпк|вос-?\d*|зс|р\.н\.|в/ч|в\.ч\.|№|п\.|пп\.|ст\.)\s+"
    return re.sub(pattern, lambda m: f"{m.group(1)} ", result)


def clean_duplicated_units(text: str) -> str:
    """Усуває повторення фраз 'військової частини' та однакових шифрів підряд із збереженням регістру."""
    def _rep_phrase(match):
        m_txt = match.group(0)
        letters = [c for c in m_txt if c.isalpha()]
        if letters and all(c.isupper() for c in letters):
            return "ВІЙСЬКОВОЇ ЧАСТИНИ "
        return "військової частини "

    t = re.sub(
        r"\b(?:військов(?:ої|а|у|ій|ою)\s+частин(?:и|а|у|і|ою)\s*){2,}",
        _rep_phrase,
        text or "",
        flags=re.IGNORECASE,
    )
    t = re.sub(
        r"\b(військов\w+\s+частин\w+\s+[АA]\d+)(?:\s+\1\b)+",
        r"\1",
        t,
        flags=re.IGNORECASE,
    )
    t = re.sub(
        r"\b(ВІЙСЬКОВ\w+\s+ЧАСТИН\w+\s+[АA]\d+)(?:\s+\1\b)+",
        r"\1",
        t,
    )
    return t


def ensure_blank_line_before_items(text: str) -> str:
    """Гарантує наявність рівно 1 порожнього рядка перед кожним пунктом наказу
    та рівно 2 порожніх рядків перед підписантом наказу.
    Якщо порожній рядок вже є — додатковий не вставляється, якщо їх кілька — згортається до потрібної кількості."""
    lines = (text or "").splitlines()
    result = []
    for idx, line in enumerate(lines):
        clean = line.strip()
        is_item = bool(re.match(r"^\d{1,3}(?:\.\d{1,3})*[\.\)]\s+", clean))
        is_signer = bool(_ORDER_SIGNER_START_RE.match(clean))

        if is_signer and idx > 0 and result:
            while result and result[-1].strip() == "":
                result.pop()
            result.append("")
            result.append("")
        elif is_item and idx > 0 and result:
            while len(result) > 1 and result[-1].strip() == "" and result[-2].strip() == "":
                result.pop()
            if result[-1].strip() != "":
                result.append("")
        result.append(line)
    return "\n".join(result)


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


_UNMATCHED_OPEN_UNIT_RE = re.compile(
    r"\b(?:"
    r"(?:\d{1,3}\s*(?:-?[а-яіїєґ]+)?\s*)?(?:окрем\w+\s+)+(?:механізован\w+|танков\w+|десантн\w+|артилерійськ\w+|піхотн\w+|єгерськ\w+|стрілецьк\w+|штурмов\w+|розвідувальн\w+|гірсько-штурмов\w+|десантно-штурмов\w+|аеромобільн\w+|повітряно-десантн\w+|зв['’]язку)?\s*(?:бригад\w*|полк\w*|батальйон\w*|дивізіон\w*|загін\w*|центр\w*)"
    r"|"
    r"\d{1,3}\s*(?:-?[а-яіїєґ]+)?\s*(?:механізован\w+|танков\w+|десантн\w+|артилерійськ\w+|піхотн\w+|стрілецьк\w+|штурмов\w+|десантно-штурмов\w+|аеромобільн\w+)?\s*(?:бригад\w*|полк\w*|армійськ\w+\s+корпус\w*|АК)"
    r"|"
    r"(?:армійськ\w+\s+корпус\w*|\b\d{1,3}\s*АК\b)"
    r")\b",
    re.IGNORECASE | re.UNICODE,
)

def find_unmatched_open_unit_spans(text: str) -> list[tuple[int, int]]:
    """Повертає діапазони відкритих назв частин, які лишилися після шифрування.
    
    Не підсвічує лінійні внутрішні батальйони/дивізіони, які вже належать закритій в/ч.
    """
    if not text:
        return []
    spans = []
    for match in _UNMATCHED_OPEN_UNIT_RE.finditer(text):
        start, end = match.start(), match.end()
        following_text = text[end:end + 60]
        if re.match(r"^\s*(?:військов\w+\s+частин\w*|в\s*/?\s*ч)\s+[АA]?\d+", following_text, re.IGNORECASE):
            continue
        spans.append((start, end))
    return spans


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
        cipher_text = standalone_cipher(entry)
        cipher = str(entry.get("cipher") or "").strip()
        if not recipient_to:
            return cipher_text
        if not cipher or cipher.casefold() in recipient_to.casefold():
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
    - лише ТЦК → «начальникам ТЦК»;
    - і частини, і ТЦК → «командирам військових частин та начальникам ТЦК».
    """
    has_units = bool(groups.get("corps") or groups.get("units"))
    has_tck = bool(groups.get("tck"))
    if has_units and has_tck:
        return "Командирам військових частин та Начальникам ТЦК"
    if has_tck:
        return "Начальникам ТЦК"
    if has_units:
        return "Командирам військових частин"
    return ""


# Написання тегу типу адресата, які розпізнаються в шаблонах повідомлень.
# Пошук у Word не враховує регістр, тому достатньо варіантів написання «чі/чи».
_MESSAGE_ADDRESSEE_KIND_TAGS = ("{{тцк чі вч}}", "{{тцк чи вч}}")


_ORDER_SIGNER_START_RE = re.compile(
    r"^\s*(?:т\.?\s*в\.?\s*о\.?|тимчасово\s+виконуюч(?:ий|а)?|"
    r"командувач|командир|начальник|заступник\s+командувача)\b",
    re.IGNORECASE | re.UNICODE,
)
_ORDER_SIGNER_RANK_RE = re.compile(
    r"^\s*(генерал(?:[-\s](?:майор|лейтенант|полковник))?|адмірал(?:[-\s]\w+)?|"
    r"полковник|підполковник|майор|капітан(?:[-\s](?:лейтенант|[1-3]\s+рангу))?|"
    r"старший\s+лейтенант|молодший\s+лейтенант|лейтенант|головний\s+сержант|"
    r"штаб[-\s]сержант|майстер[-\s]сержант|старший\s+сержант|молодший\s+сержант|"
    r"сержант|старшина|солдат|матрос)\b",
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


def copy_template_for_editing(template_path: str, output_path: str) -> str:
    """Створює робочу копію шаблону з коректним для його вмісту розширенням.

    Повертає шлях робочої копії. Якщо він відрізняється від `output_path`,
    викликач має після `SaveAs2` прибрати проміжний файл.
    """
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
    def __init__(self, root: tb.Window):
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

        # Конфігурація вкладки «Повідомлення».
        self.message_cover_template_path = tk.StringVar()
        self.message_content_template_path = tk.StringVar()
        self.message_out_folder = tk.StringVar()
        self.message_executor = tk.StringVar()

        # Налаштування теми
        self.current_theme = tk.StringVar(value="cosmo")
        self.config_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")

        self.load_config()
        self.create_widgets()

        # Підключення Drag-and-Drop
        if HAS_WINDND:
            try:
                windnd.hook_dropfiles(self.root, func=self.handle_drag_and_drop)
            except Exception as e:
                print(f"Помилка ініціалізації Drag-and-Drop: {e}")

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

        for fpath in files:
            fname = os.path.basename(fpath).lower()
            if os.path.isdir(fpath):
                if current_tab == TAB_COPIES:
                    self.p2_orders_folder.set(fpath)
                    self.p2_source_mode.set("folder")
                    self._on_p2_source_mode_changed()
                    self.log_p2(f"📥 [Drag-and-Drop] Папку наказів встановлено: {fpath}")
                elif current_tab == TAB_EXTRACTS:
                    self.out_folder.set(fpath)
                    self.log(f"📥 [Drag-and-Drop] Папку результату встановлено: {fpath}")
                elif current_tab == TAB_MESSAGES:
                    self.message_out_folder.set(fpath)
                    self.log(f"📥 [Drag-and-Drop] Папку результату повідомлень встановлено: {fpath}")
                handled_count += 1
            elif fname.endswith((".xlsx", ".xls")):
                self.excel_path.set(fpath)
                self.log(f"📥 [Drag-and-Drop] Словник Excel встановлено: {os.path.basename(fpath)}")
                handled_count += 1
            elif fname.endswith(".docx") and not fname.startswith("~$"):
                if current_tab == TAB_COPIES:
                    if "задн" in fname or "back" in fname or "шаблон" in fname:
                        self.p2_back_page_path.set(fpath)
                        self.log_p2(f"📥 [Drag-and-Drop] Шаблон «Задня сторінка» встановлено: {os.path.basename(fpath)}")
                    else:
                        self.p2_single_file.set(fpath)
                        self.p2_source_mode.set("single")
                        self._on_p2_source_mode_changed()
                        self.log_p2(f"📥 [Drag-and-Drop] Наказ для примірника встановлено: {os.path.basename(fpath)}")
                elif current_tab == TAB_EXTRACTS:
                    if "шаблон" in fname or "template" in fname or "зразок" in fname:
                        self.template_path.set(fpath)
                        self.log(f"📥 [Drag-and-Drop] Шаблон витягу встановлено: {os.path.basename(fpath)}")
                    else:
                        self.doc_path.set(fpath)
                        order_dir = os.path.dirname(fpath)
                        self.out_folder.set(os.path.join(order_dir, "Extracts_Output"))
                        self._refresh_order_signer()
                        self.log(f"📥 [Drag-and-Drop] Наказ завантажено: {os.path.basename(fpath)}")
                elif current_tab == TAB_MESSAGES:
                    if "зміст" in fname or "content" in fname:
                        self.message_content_template_path.set(fpath)
                        self.log(f"📥 [Drag-and-Drop] Шаблон змісту повідомлення встановлено: {os.path.basename(fpath)}")
                    elif "титул" in fname or "cover" in fname or "шаблон" in fname:
                        self.message_cover_template_path.set(fpath)
                        self.log(f"📥 [Drag-and-Drop] Шаблон супровідного повідомлення встановлено: {os.path.basename(fpath)}")
                    else:
                        self.doc_path.set(fpath)
                        self._refresh_order_signer()
                        self.log(f"📥 [Drag-and-Drop] Наказ для повідомлень завантажено: {os.path.basename(fpath)}")
                handled_count += 1

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
                    if "group_corps" in data:
                        self.group_corps_var.set(data["group_corps"])
                    if "duplex_2up_layout" in data:
                        self.duplex_2up_layout.set(data["duplex_2up_layout"])

                    self.p2_orders_folder.set(data.get("p2_orders_folder", ""))
                    self.p2_single_file.set(data.get("p2_single_file", ""))
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

                    self.message_cover_template_path.set(data.get("message_cover_template_path", ""))
                    self.message_content_template_path.set(data.get("message_content_template_path", ""))
                    self.message_out_folder.set(data.get("message_out_folder", ""))
                    self.message_executor.set(data.get("message_executor", ""))

                    theme = data.get("theme", "cosmo")
                    self.current_theme.set(theme)
                    if hasattr(self.root, "style"):
                        self.root.style.theme_use(theme)
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
            "group_corps": self.group_corps_var.get(),
            "duplex_2up_layout": self.duplex_2up_layout.get(),
            "p2_orders_folder": self.p2_orders_folder.get(),
            "p2_single_file": self.p2_single_file.get(),
            "p2_back_page_path": self.p2_back_page_path.get(),
            "p2_out_folder": self.p2_out_folder.get(),
            "p2_out_folder_manual": self.p2_out_folder_manual.get(),
            "p2_copy_title": self.p2_copy_title.get(),
            "p2_executor": self.p2_executor.get(),
            "p2_preview": self.p2_preview.get(),
            "p2_preview_delay": self.p2_preview_delay.get(),
            "message_cover_template_path": self.message_cover_template_path.get(),
            "message_content_template_path": self.message_content_template_path.get(),
            "message_out_folder": self.message_out_folder.get(),
            "message_executor": self.message_executor.get(),
            "theme": self.current_theme.get(),
        }
        try:
            with open(self.config_file, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
        except Exception:
            pass

    def toggle_theme(self):
        new_theme = "darkly" if self.current_theme.get() in ("cosmo", "flatly", "litera") else "cosmo"
        self.current_theme.set(new_theme)
        if HAS_TTKBOOTSTRAP and hasattr(self.root, "style"):
            try:
                self.root.style.theme_use(new_theme)
            except Exception:
                pass
        self.theme_btn.config(text="🌙 Темна тема" if new_theme == "cosmo" else "☀️ Світла тема")
        self.save_config()

    def create_widgets(self):
        main_container = tb.Frame(self.root, padding=12)
        main_container.pack(fill=BOTH, expand=True)

        # Верхній Header
        header_frame = tb.Frame(main_container)
        header_frame.pack(fill=X, pady=(0, 10))

        title_box = tb.Frame(header_frame)
        title_box.pack(side=LEFT, fill=Y)

        title_label = tb.Label(
            title_box,
            text="⚡ Node Automation Toolkit",
            font=("Segoe UI", 16, "bold"),
            bootstyle="primary",
        )
        title_label.pack(anchor=W)

        subtitle_label = tb.Label(
            title_box,
            text="Автоматизація діловодства: розрахунок розсилки, витяги та примірники наказів",
            font=("Segoe UI", 9),
            bootstyle="secondary",
        )
        subtitle_label.pack(anchor=W)

        self.theme_btn = tb.Button(
            header_frame,
            text="🌙 Темна тема" if self.current_theme.get() == "cosmo" else "☀️ Світла тема",
            bootstyle="secondary-outline",
            command=self.toggle_theme,
        )
        self.theme_btn.pack(side=RIGHT, padx=4)

        # Головні робочі вкладки
        self.notebook = tb.Notebook(main_container, bootstyle="primary")
        self.notebook.pack(fill=BOTH, expand=True)

        self.tab_copies = tb.Frame(self.notebook, padding=10)
        self.notebook.add(self.tab_copies, text="  📑 1. Примірники 2/3  ")

        self.tab_extracts = tb.Frame(self.notebook, padding=10)
        self.notebook.add(self.tab_extracts, text="  📄 2. Розрахунок та витяги  ")

        self.tab_messages = tb.Frame(self.notebook, padding=10)
        self.notebook.add(self.tab_messages, text="  💬 3. Повідомлення  ")

        self._build_tab_extracts()
        self._build_tab_copies()
        self._build_tab_messages()

    # =========================================================================
    # ВКЛАДКА 1: РОЗРАХУНОК ТА ВИТЯГИ
    # =========================================================================
    def _build_tab_extracts(self):
        tab = self.tab_extracts
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(3, weight=3)
        tab.rowconfigure(4, weight=1)

        # Блок 1: Вхідні файли
        files_card = tb.Labelframe(tab, text=" Вхідні файли та шаблони ", padding=10, bootstyle="info")
        files_card.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        files_card.columnconfigure(1, weight=1)

        tb.Label(files_card, text="1. Словник (Excel):").grid(row=0, column=0, sticky=W, padx=(0, 8), pady=3)
        tb.Entry(files_card, textvariable=self.excel_path).grid(row=0, column=1, sticky="ew", pady=3)
        tb.Button(files_card, text="📂 Вибрати", bootstyle="info-outline", command=self.select_excel).grid(
            row=0, column=2, padx=(6, 0), pady=3
        )

        tb.Label(files_card, text="2. Наказ (DOCX):").grid(row=1, column=0, sticky=W, padx=(0, 8), pady=3)
        tb.Entry(files_card, textvariable=self.doc_path).grid(row=1, column=1, sticky="ew", pady=3)
        tb.Button(files_card, text="📂 Вибрати", bootstyle="info-outline", command=self.select_doc).grid(
            row=1, column=2, padx=(6, 0), pady=3
        )

        tb.Label(files_card, text="3. Шаблон витягу:").grid(row=2, column=0, sticky=W, padx=(0, 8), pady=3)
        tb.Entry(files_card, textvariable=self.template_path).grid(row=2, column=1, sticky="ew", pady=3)
        tpl_btn_box = tb.Frame(files_card)
        tpl_btn_box.grid(row=2, column=2, padx=(6, 0), pady=3)
        tb.Button(tpl_btn_box, text="📂 Вибрати", bootstyle="info-outline", command=self.select_template).pack(side=LEFT)
        tb.Button(tpl_btn_box, text="✏️ Редагувати", bootstyle="secondary-outline", command=self.edit_template).pack(side=LEFT, padx=(4, 0))

        tb.Label(files_card, text="Папка результату:").grid(row=3, column=0, sticky=W, padx=(0, 8), pady=3)
        tb.Entry(files_card, textvariable=self.out_folder).grid(row=3, column=1, sticky="ew", pady=3)
        tb.Button(files_card, text="📂 Вибрати", bootstyle="secondary-outline", command=self.select_folder).grid(
            row=3, column=2, padx=(6, 0), pady=3
        )

        tb.Label(
            files_card,
            text="💡 Перетягніть сюди файл наказу (.docx), словник (.xlsx) чи шаблон (Drag-and-Drop)",
            font=("Segoe UI", 9, "italic"),
            bootstyle="info",
        ).grid(row=4, column=0, columnspan=3, sticky=W, pady=(4, 0))

        copy_sources = tb.Labelframe(
            files_card,
            text=" Примірники № 2 з останнього пакетного проходу ",
            padding=4,
            bootstyle="success",
        )
        copy_sources.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(7, 0))
        copy_sources.columnconfigure(0, weight=1)
        self.copy_two_tree = tb.Treeview(
            copy_sources,
            columns=("use", "file"),
            show="headings",
            height=3,
            selectmode="none",
            bootstyle="success",
        )
        self.copy_two_tree.heading("use", text="Використати")
        self.copy_two_tree.heading("file", text="Файл примірника № 2")
        self.copy_two_tree.column("use", width=100, minwidth=100, stretch=False, anchor="center")
        self.copy_two_tree.column("file", width=620, minwidth=250, stretch=True, anchor=W)
        self.copy_two_tree.grid(row=0, column=0, sticky="ew")
        self.copy_two_tree.bind("<Button-1>", self._toggle_copy_two_source)
        self.copy_two_tree.insert("", tk.END, values=("—", "Спершу сформуйте примірники № 2 у вкладці «Примірники 2/3»."))

        # Блок 2: Параметри та підписанти
        opts_card = tb.Labelframe(tab, text=" Параметри, підписант наказу та засвідчення витягів ", padding=10, bootstyle="secondary")
        opts_card.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        opts_card.columnconfigure(1, weight=1)

        tb.Checkbutton(
            opts_card,
            text="🖨️ Оптимізація під друк «2 сторінки на 1 аркуш» (вирівнювання та порожні сторінки)",
            variable=self.duplex_2up_layout,
            bootstyle="success-round-toggle",
        ).grid(row=0, column=0, columnspan=2, sticky=W, pady=(0, 4))

        tb.Checkbutton(
            opts_card,
            text="Групувати підпорядковані частини по Корпусах (Варіант 1)",
            variable=self.group_corps_var,
            bootstyle="primary-round-toggle",
        ).grid(row=1, column=0, columnspan=2, sticky=W, pady=(0, 4))

        tb.Label(opts_card, text="Виконавець:").grid(row=2, column=0, sticky=W, padx=(0, 8), pady=2)
        tb.Entry(opts_card, textvariable=self.executor).grid(row=2, column=1, sticky="ew", pady=2)

        # 1. Підписант наказу (Командувач/Командир — зчитується автоматично після пунктів)
        signer_box = tb.Labelframe(opts_card, text=" Підписант оригіналу наказу (зчитується автоматично після пунктів) ", padding=6, bootstyle="info")
        signer_box.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(6, 4))
        signer_box.columnconfigure(1, weight=1)

        tb.Label(signer_box, text="Посада:").grid(row=0, column=0, sticky=W, padx=(0, 8), pady=2)
        tb.Entry(signer_box, textvariable=self.order_signer_position).grid(row=0, column=1, sticky="ew", pady=2)

        row_sr = tb.Frame(signer_box)
        row_sr.grid(row=1, column=0, columnspan=2, sticky="ew", pady=2)
        row_sr.columnconfigure(1, weight=1)
        row_sr.columnconfigure(3, weight=1)

        tb.Label(row_sr, text="Звання:").grid(row=0, column=0, sticky=W, padx=(0, 6))
        tb.Entry(row_sr, textvariable=self.order_signer_rank).grid(row=0, column=1, sticky="ew", padx=(0, 12))

        tb.Label(row_sr, text="ПІБ:").grid(row=0, column=2, sticky=W, padx=(0, 6))
        tb.Entry(row_sr, textvariable=self.order_signer_name).grid(row=0, column=3, sticky="ew")

        # 2. Особа, яка засвідчує витяг («Згідно з оригіналом»)
        cert_box = tb.Labelframe(opts_card, text=" Особа, яка засвідчує витяг («Згідно з оригіналом») ", padding=6, bootstyle="success")
        cert_box.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(4, 0))
        cert_box.columnconfigure(1, weight=1)

        tb.Label(cert_box, text="Посада:").grid(row=0, column=0, sticky=W, padx=(0, 8), pady=2)
        tb.Entry(cert_box, textvariable=self.certifier_position).grid(row=0, column=1, sticky="ew", pady=2)

        row_cr = tb.Frame(cert_box)
        row_cr.grid(row=1, column=0, columnspan=2, sticky="ew", pady=2)
        row_cr.columnconfigure(1, weight=1)
        row_cr.columnconfigure(3, weight=1)

        tb.Label(row_cr, text="Звання:").grid(row=0, column=0, sticky=W, padx=(0, 6))
        tb.Entry(row_cr, textvariable=self.certifier_rank).grid(row=0, column=1, sticky="ew", padx=(0, 12))

        tb.Label(row_cr, text="ПІБ:").grid(row=0, column=2, sticky=W, padx=(0, 6))
        tb.Entry(row_cr, textvariable=self.certifier_name).grid(row=0, column=3, sticky="ew")

        # Блок 3: Кнопки дій
        actions_bar = tb.Frame(tab)
        actions_bar.grid(row=2, column=0, sticky="ew", pady=(0, 8))

        self.btn_calc = tb.Button(
            actions_bar,
            text="📊 1. Розрахувати розсилку",
            bootstyle="primary",
            command=self.run_rozrahunok_action,
        )
        self.btn_calc.pack(side=LEFT, padx=(0, 6))

        self.btn_extracts = tb.Button(
            actions_bar,
            text="📑 2. Створити витяги (Word)",
            bootstyle="success",
            command=self.run_extracts_action,
        )
        self.btn_extracts.pack(side=LEFT, padx=6)

        tb.Button(
            actions_bar,
            text="ℹ️ Теги шаблону {{…}}",
            bootstyle="info-outline",
            command=self.show_template_tags,
        ).pack(side=LEFT, padx=6)

        tb.Button(
            actions_bar,
            text="🔍 Порівняти з еталоном (Compare)",
            bootstyle="warning-outline",
            command=self.open_compare_extracts,
        ).pack(side=LEFT, padx=6)

        tb.Button(
            actions_bar,
            text="🧹 Очистити журнал",
            bootstyle="secondary-outline",
            command=self.clear_log,
        ).pack(side=RIGHT)

        # Блок 4: Результати аналізу
        results_card = tb.Labelframe(tab, text=" Результати аналізу ", padding=6, bootstyle="primary")
        results_card.grid(row=3, column=0, sticky="nsew", pady=(0, 6))
        results_card.columnconfigure(0, weight=1)
        results_card.rowconfigure(0, weight=1)

        self.results_notebook = tb.Notebook(results_card, bootstyle="info")
        self.results_notebook.grid(row=0, column=0, sticky="nsew")

        self.result_views: dict[str, ttk.Treeview] = {}
        self.result_tabs: dict[str, tb.Frame] = {}

        self._create_result_tab(
            "calculation", "📊 Розсилка",
            ["Військова частина / Відправник", "Пункти витягу", "Кількість"],
            "— Очікується розрахунок розсилки —"
        )
        self._create_result_tab(
            "unmatched", "⚠️ Пропущені",
            ["Пункт", "Текст пункту", "Причина"],
            "— Пропущені пункти відсутні —"
        )
        self._create_result_tab(
            "routing", "🔍 Контроль маршрутизації",
            ["Пункт", "Збіги з таблиці", "Застосовані правила", "Підсумкові адресати"],
            "— Очікується аналіз маршрутизації —"
        )
        self._create_result_tab(
            "layout", "📐 Макет витягів",
            ["Стан макета"],
            "— Очікується формування витягів —"
        )

        # Блок 5: Журнал (Log)
        log_card = tb.Labelframe(tab, text=" Журнал виконання ", padding=4, bootstyle="secondary")
        log_card.grid(row=4, column=0, sticky="nsew")
        log_card.columnconfigure(0, weight=1)
        log_card.rowconfigure(1, weight=1)

        log_toolbar = tb.Frame(log_card)
        log_toolbar.grid(row=0, column=0, sticky="ew", padx=2, pady=(0, 2))
        tb.Button(
            log_toolbar,
            text="📋 Скопіювати журнал",
            bootstyle="info-outline",
            command=lambda: self.copy_log(self.log_text),
        ).pack(side=LEFT, padx=(0, 4))
        tb.Button(
            log_toolbar,
            text="🧹 Очистити",
            bootstyle="secondary-outline",
            command=self.clear_log,
        ).pack(side=LEFT)

        self.log_text = ScrolledText(log_card, height=5, wrap=tk.WORD, font=("Segoe UI", 9))
        self.log_text.grid(row=1, column=0, sticky="nsew")
        self._setup_text_copy_menu(self.log_text)

        self.log("Готовий до роботи. Вкажіть файли та натисніть потрібну дію.")

    # =========================================================================
    # ВКЛАДКА 2: ПРИМІРНИКИ 2/3
    # =========================================================================
    def _build_tab_copies(self):
        tab = self.tab_copies
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(2, weight=3)
        tab.rowconfigure(3, weight=1)

        # Блок 1: Вхідні параметри
        copies_card = tb.Labelframe(tab, text=" Параметри генерації примірників ", padding=10, bootstyle="primary")
        copies_card.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        copies_card.columnconfigure(1, weight=1)

        mode_box = tb.Frame(copies_card)
        mode_box.grid(row=0, column=0, columnspan=3, sticky=W, pady=(0, 6))
        tb.Radiobutton(
            mode_box,
            text="Пакетна обробка папки з наказами",
            variable=self.p2_source_mode,
            value="folder",
            command=self._on_p2_source_mode_changed,
            bootstyle="primary",
        ).pack(side=LEFT, padx=(0, 14))
        tb.Radiobutton(
            mode_box,
            text="Окремий файл наказу",
            variable=self.p2_source_mode,
            value="file",
            command=self._on_p2_source_mode_changed,
            bootstyle="primary",
        ).pack(side=LEFT)

        self.p2_lbl_folder = tb.Label(copies_card, text="Папка з наказами (DOCX):")
        self.p2_lbl_folder.grid(row=1, column=0, sticky=W, padx=(0, 8), pady=3)
        self.p2_entry_folder = tb.Entry(copies_card, textvariable=self.p2_orders_folder)
        self.p2_entry_folder.grid(row=1, column=1, sticky="ew", pady=3)
        self.p2_btn_folder = tb.Button(
            copies_card, text="📂 Вибрати", bootstyle="primary-outline", command=self.select_p2_folder
        )
        self.p2_btn_folder.grid(row=1, column=2, padx=(6, 0), pady=3)

        self.p2_lbl_file = tb.Label(copies_card, text="Окремий файл наказу:")
        self.p2_lbl_file.grid(row=2, column=0, sticky=W, padx=(0, 8), pady=3)
        self.p2_entry_file = tb.Entry(copies_card, textvariable=self.p2_single_file)
        self.p2_entry_file.grid(row=2, column=1, sticky="ew", pady=3)
        self.p2_btn_file = tb.Button(
            copies_card, text="📂 Вибрати", bootstyle="primary-outline", command=self.select_p2_file
        )
        self.p2_btn_file.grid(row=2, column=2, padx=(6, 0), pady=3)

        tb.Label(copies_card, text="Шаблон «Задня сторінка»:").grid(row=3, column=0, sticky=W, padx=(0, 8), pady=3)
        tb.Entry(copies_card, textvariable=self.p2_back_page_path).grid(row=3, column=1, sticky="ew", pady=3)
        back_btn_box = tb.Frame(copies_card)
        back_btn_box.grid(row=3, column=2, padx=(6, 0), pady=3)
        tb.Button(back_btn_box, text="📂 Вибрати", bootstyle="primary-outline", command=self.select_p2_back_page).pack(side=LEFT)
        tb.Button(back_btn_box, text="✏️ Редагувати", bootstyle="secondary-outline", command=self.edit_p2_back_page).pack(side=LEFT, padx=(4, 0))

        tb.Label(copies_card, text="Папка результатів:").grid(row=4, column=0, sticky=W, padx=(0, 8), pady=3)
        tb.Entry(copies_card, textvariable=self.p2_out_folder).grid(row=4, column=1, sticky="ew", pady=3)
        tb.Button(copies_card, text="📂 Вибрати", bootstyle="secondary-outline", command=self.select_p2_out_folder).grid(
            row=4, column=2, padx=(6, 0), pady=3
        )

        tb.Label(copies_card, text="Виконавець:").grid(row=5, column=0, sticky=W, padx=(0, 8), pady=3)
        tb.Entry(copies_card, textvariable=self.p2_executor).grid(row=5, column=1, sticky="ew", pady=3)
        tb.Label(
            copies_card,
            text="порожньо = з витягів",
            bootstyle="secondary",
        ).grid(row=5, column=2, sticky=W, padx=(6, 0), pady=3)

        cert_p2 = tb.Labelframe(copies_card, text="Згідно з оригіналом (засвідчувач)", padding=8)
        cert_p2.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(8, 4))
        cert_p2.columnconfigure(1, weight=1)
        tb.Label(cert_p2, text="Посада:").grid(row=0, column=0, sticky=W, padx=(0, 8), pady=2)
        tb.Entry(cert_p2, textvariable=self.certifier_position).grid(row=0, column=1, sticky="ew", pady=2)
        tb.Label(cert_p2, text="Звання:").grid(row=1, column=0, sticky=W, padx=(0, 8), pady=2)
        tb.Entry(cert_p2, textvariable=self.certifier_rank).grid(row=1, column=1, sticky="ew", pady=2)
        tb.Label(cert_p2, text="ПІБ:").grid(row=2, column=0, sticky=W, padx=(0, 8), pady=2)
        tb.Entry(cert_p2, textvariable=self.certifier_name).grid(row=2, column=1, sticky="ew", pady=2)
        tb.Label(
            cert_p2,
            text="Спільні поля з вкладкою витягів — заповнюються один раз.",
            bootstyle="secondary",
        ).grid(row=3, column=0, columnspan=2, sticky=W, pady=(4, 0))

        preview_box = tb.Frame(copies_card)
        preview_box.grid(row=7, column=0, columnspan=3, sticky=W, pady=(6, 0))
        tb.Checkbutton(
            preview_box,
            text="🐢 Режим превʼю (повільно, з видимим Word)",
            variable=self.p2_preview,
            bootstyle="info-round-toggle",
        ).pack(side=LEFT, padx=(0, 10))
        tb.Label(preview_box, text="пауза, сек:").pack(side=LEFT, padx=(0, 4))
        tb.Entry(preview_box, textvariable=self.p2_preview_delay, width=6).pack(side=LEFT)
        tb.Label(
            preview_box,
            text="Показує кожен крок у журналі. Для великого пакета — довго.",
            bootstyle="secondary",
        ).pack(side=LEFT, padx=(10, 0))

        copy_type_box = tb.Frame(copies_card)
        copy_type_box.grid(row=8, column=0, columnspan=3, sticky=W, pady=(4, 0))
        tb.Label(copy_type_box, text="Формується:", font=("Segoe UI", 10, "bold")).pack(side=LEFT, padx=(0, 8))
        tb.Label(copy_type_box, text="Примірник № 2", bootstyle="success").pack(side=LEFT, padx=(0, 10))
        tb.Label(
            copy_type_box,
            text="Правила для примірника № 3 ще не погоджені.",
            bootstyle="secondary",
        ).pack(side=LEFT)

        tb.Label(
            copies_card,
            text="💡 Перетягніть сюди папку з наказами, окремий наказ чи шаблон (Drag-and-Drop)",
            font=("Segoe UI", 9, "italic"),
            bootstyle="info",
        ).grid(row=9, column=0, columnspan=3, sticky=W, pady=(4, 0))

        self._on_p2_source_mode_changed()

        # Блок 2: Кнопки запуску
        p2_actions = tb.Frame(tab)
        p2_actions.grid(row=1, column=0, sticky="ew", pady=(0, 8))

        self.btn_run_p2 = tb.Button(
            p2_actions,
            text="🚀 Сформувати примірники (заміна останньої сторінки)",
            bootstyle="primary",
            command=self.run_generate_copies,
        )
        self.btn_run_p2.pack(side=LEFT, padx=(0, 8))

        tb.Button(
            p2_actions,
            text="ℹ️ Правила примірника 2",
            bootstyle="info-outline",
            command=self.show_p2_info,
        ).pack(side=LEFT, padx=(0, 8))

        tb.Button(
            p2_actions,
            text="🔍 Порівняти з еталоном (Compare)",
            bootstyle="warning-outline",
            command=self.open_compare_copies,
        ).pack(side=LEFT)

        # Блок 3: Таблиця створених примірників
        copies_list_card = tb.Labelframe(tab, text=" Сформовані примірники наказів ", padding=6, bootstyle="success")
        copies_list_card.grid(row=2, column=0, sticky="nsew", pady=(0, 8))
        copies_list_card.columnconfigure(0, weight=1)
        copies_list_card.rowconfigure(0, weight=1)

        tree_frame = tb.Frame(copies_list_card)
        tree_frame.grid(row=0, column=0, sticky="nsew")
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)

        p2_cols = ["№", "Назва файлу", "Номер наказу", "Дата наказу", "Сторінок", "Арк. для друку", "Повний шлях"]
        self.p2_tree = tb.Treeview(tree_frame, columns=p2_cols, show="headings", selectmode="extended", bootstyle="success")
        p2_vscroll = tb.Scrollbar(tree_frame, orient=VERTICAL, command=self.p2_tree.yview)
        p2_hscroll = tb.Scrollbar(tree_frame, orient=HORIZONTAL, command=self.p2_tree.xview)
        self.p2_tree.configure(yscrollcommand=p2_vscroll.set, xscrollcommand=p2_hscroll.set)

        self.p2_tree.grid(row=0, column=0, sticky="nsew")
        p2_vscroll.grid(row=0, column=1, sticky="ns")
        p2_hscroll.grid(row=1, column=0, sticky="ew")

        widths = {
            "№": 45,
            "Назва файлу": 230,
            "Номер наказу": 110,
            "Дата наказу": 110,
            "Сторінок": 75,
            "Арк. для друку": 120,
            "Повний шлях": 350,
        }
        for col in p2_cols:
            self.p2_tree.heading(col, text=col)
            self.p2_tree.column(col, width=widths.get(col, 120), anchor=W)
        self.p2_tree.insert("", tk.END, values=("1", "— Очікується запуск формування примірників —", "—", "—", "—", "—", "—"))

        p2_selected_actions = tb.Frame(copies_list_card)
        p2_selected_actions.grid(row=1, column=0, sticky="ew", pady=(6, 0))

        tb.Button(
            p2_selected_actions,
            text="➡️ Передати вибраний примірник у Вкладку 2 (Витяги)",
            bootstyle="success",
            command=self.transfer_selected_copy_to_extracts,
        ).pack(side=LEFT, padx=(0, 6))

        tb.Button(
            p2_selected_actions,
            text="☑️ Вибрати всі",
            bootstyle="info-outline",
            command=self.select_all_copies,
        ).pack(side=LEFT, padx=(0, 6))

        tb.Button(
            p2_selected_actions,
            text="◻️ Зняти вибір",
            bootstyle="secondary-outline",
            command=self.deselect_all_copies,
        ).pack(side=LEFT, padx=(0, 6))

        tb.Button(
            p2_selected_actions,
            text="📂 Папка результату",
            bootstyle="secondary-outline",
            command=self.open_p2_output_folder,
        ).pack(side=LEFT, padx=(0, 6))

        tb.Button(
            p2_selected_actions,
            text="👁️ Відкрити у Word",
            bootstyle="info-outline",
            command=self.open_selected_copy_in_word,
        ).pack(side=LEFT)

        # Блок 4: Журнал для примірників
        p2_log_card = tb.Labelframe(tab, text=" Журнал примірників ", padding=4, bootstyle="secondary")
        p2_log_card.grid(row=3, column=0, sticky="nsew")
        p2_log_card.columnconfigure(0, weight=1)
        p2_log_card.rowconfigure(1, weight=1)

        p2_log_toolbar = tb.Frame(p2_log_card)
        p2_log_toolbar.grid(row=0, column=0, sticky="ew", padx=2, pady=(0, 2))
        tb.Button(
            p2_log_toolbar,
            text="📋 Скопіювати журнал",
            bootstyle="info-outline",
            command=lambda: self.copy_log(self.p2_log_text),
        ).pack(side=LEFT, padx=(0, 4))
        tb.Button(
            p2_log_toolbar,
            text="🧹 Очистити",
            bootstyle="secondary-outline",
            command=lambda: self.p2_log_text.delete(1.0, tk.END),
        ).pack(side=LEFT)

        self.p2_log_text = ScrolledText(p2_log_card, height=4, wrap=tk.WORD, font=("Segoe UI", 9))
        self.p2_log_text.grid(row=1, column=0, sticky="nsew")
        self._setup_text_copy_menu(self.p2_log_text)

    def _on_p2_source_mode_changed(self):
        if self.p2_source_mode.get() == "folder":
            self.p2_entry_folder.config(state=NORMAL)
            self.p2_btn_folder.config(state=NORMAL)
            self.p2_entry_file.config(state=DISABLED)
            self.p2_btn_file.config(state=DISABLED)
        else:
            self.p2_entry_folder.config(state=DISABLED)
            self.p2_btn_folder.config(state=DISABLED)
            self.p2_entry_file.config(state=NORMAL)
            self.p2_btn_file.config(state=NORMAL)

    # =========================================================================
    # ВКЛАДКА 3: ПОВІДОМЛЕННЯ ПРО ПРИЙНЯТТЯ
    # =========================================================================
    def _build_tab_messages(self):
        tab = self.tab_messages
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(2, weight=1)

        card = tb.Labelframe(tab, text=" Повідомлення про прийняття наказу ", padding=10, bootstyle="info")
        card.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        card.columnconfigure(1, weight=1)

        fields = (
            ("Словник (Excel):", self.excel_path, self.select_excel),
            ("Наказ (DOCX):", self.doc_path, self.select_doc),
            ("1. Шаблон супроводу:", self.message_cover_template_path, self.select_message_cover_template),
            ("2. Шаблон зі змістом:", self.message_content_template_path, self.select_message_content_template),
            ("Папка результату:", self.message_out_folder, self.select_message_output_folder),
        )
        for row, (label, variable, command) in enumerate(fields):
            tb.Label(card, text=label).grid(row=row, column=0, sticky=W, padx=(0, 8), pady=3)
            tb.Entry(card, textvariable=variable).grid(row=row, column=1, sticky="ew", pady=3)
            tb.Button(card, text="📂 Вибрати", bootstyle="info-outline", command=command).grid(
                row=row, column=2, padx=(6, 0), pady=3
            )

        executor_row = len(fields)
        tb.Label(card, text="Виконавець повідомлень:").grid(
            row=executor_row, column=0, sticky=W, padx=(0, 8), pady=3
        )
        tb.Entry(card, textvariable=self.message_executor).grid(
            row=executor_row, column=1, columnspan=2, sticky="ew", pady=3
        )

        tb.Label(
            card,
            text=(
                "Стандартні теги: {{номер_наказу}}, {{дата_наказу}}, {{кому_список}}, {{куди}}, "
                "{{тцк чі вч}}, {{виконавець}}. "
                "Кожен {{кому_список}} у таблиці = один унікальний адресат. "
                "У другому шаблоні: {{зміст_шифр}}. Невпізнані відкриті назви частин виділяються жовтим."
            ),
            font=("Segoe UI", 9, "italic"),
            bootstyle="secondary",
            wraplength=900,
            justify=LEFT,
        ).grid(row=executor_row + 1, column=0, columnspan=3, sticky=W, pady=(7, 0))

        actions = tb.Frame(tab)
        actions.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        self.btn_generate_messages = tb.Button(
            actions,
            text="✉️ Створити 2 повідомлення",
            bootstyle="success",
            command=self.run_generate_messages,
        )
        self.btn_generate_messages.pack(side=LEFT, padx=(0, 8))

        tb.Button(
            actions,
            text="🔍 Порівняти з еталоном (Compare)",
            bootstyle="warning-outline",
            command=self.open_compare_messages,
        ).pack(side=LEFT)
        tb.Button(actions, text="ℹ️ Теги повідомлень", bootstyle="info-outline", command=self.show_message_tags).pack(
            side=LEFT, padx=(8, 0)
        )

        # Журнал для повідомлень
        msg_log_card = tb.Labelframe(tab, text=" Журнал повідомлень ", padding=4, bootstyle="secondary")
        msg_log_card.grid(row=2, column=0, sticky="nsew")
        msg_log_card.columnconfigure(0, weight=1)
        msg_log_card.rowconfigure(1, weight=1)

        msg_log_toolbar = tb.Frame(msg_log_card)
        msg_log_toolbar.grid(row=0, column=0, sticky="ew", padx=2, pady=(0, 2))
        tb.Button(
            msg_log_toolbar,
            text="📋 Скопіювати журнал",
            bootstyle="info-outline",
            command=lambda: self.copy_log(self.log_text),
        ).pack(side=LEFT, padx=(0, 4))
        tb.Button(
            msg_log_toolbar,
            text="🧹 Очистити",
            bootstyle="secondary-outline",
            command=self.clear_log,
        ).pack(side=LEFT)

        msg_log_text = ScrolledText(msg_log_card, height=5, wrap=tk.WORD, font=("Segoe UI", 9))
        msg_log_text.grid(row=1, column=0, sticky="nsew")
        self._setup_text_copy_menu(msg_log_text)

    # =========================================================================
    # ДОПОМІЖНІ МЕТОДИ ІНТЕРФЕЙСУ
    # =========================================================================
    def _create_result_tab(
        self, key: str, title: str, default_cols: list[str] | None = None, placeholder_text: str = ""
    ):
        tab = tb.Frame(self.results_notebook, padding=4)
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(0, weight=1)
        tree = tb.Treeview(tab, show="headings", selectmode="browse", bootstyle="primary")
        vertical = tb.Scrollbar(tab, orient=VERTICAL, command=tree.yview)
        horizontal = tb.Scrollbar(tab, orient=HORIZONTAL, command=tree.xview)
        tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        self.results_notebook.add(tab, text=title)
        self.result_views[key] = tree
        self.result_tabs[key] = tab

        if default_cols:
            tree["columns"] = default_cols
            for col in default_cols:
                tree.heading(col, text=col)
                tree.column(col, width=160, minwidth=100, stretch=True, anchor=W)
            if placeholder_text:
                row_vals = [placeholder_text] + ["—"] * (len(default_cols) - 1)
                tree.insert("", tk.END, values=tuple(row_vals))

    def show_template_tags(self):
        messagebox.showinfo(
            "Теги шаблону витягу",
            "Додайте ці теги до DOCX-шаблону витягу:\n\n"
            "• {{кому}} — значення «Кому» з Excel; якщо порожнє — шифр.\n"
            "• {{куди}} — значення «Куди» з Excel.\n"
            "• {{номер_наказу}} — номер лише з назви файла наказу.\n"
            "• {{дата_наказу}} — дата лише з назви файла наказу.\n"
            "• {{пункти}} — перелік пунктів цього витягу.\n"
            "• {{зміст}} — повний текст пунктів з оригіналу наказу.\n\n"
            "Підписант оригіналу наказу (зчитується автоматично після пунктів):\n"
            "• {{підписант_посада}}, {{підписант_звання}}, {{підписант_піб}}\n\n"
            "Засвідчення («Згідно з оригіналом»):\n"
            "• {{згідно_з_оригіналом}} або {{засвідчення}} → «Згідно з оригіналом»\n"
            "• {{засвідчувач_посада}} (або {{згідно_з_оригіналом_посада}})\n"
            "• {{засвідчувач_звання}} (або {{згідно_з_оригіналом_звання}})\n"
            "• {{засвідчувач_піб}} (або {{згідно_з_оригіналом_піб}}, {{засвідчувач}})\n\n"
            "Службові реквізити:\n"
            "• {{виконавець}} — виконавець з форми програми.",
        )

    def show_p2_info(self):
        messagebox.showinfo(
            "Правила формування примірника № 2",
            "Примірник № 2 — це повна копія вихідного наказу, у якій останню сторінку "
            "замінено окремим односторінковим DOCX «Задня сторінка».\n\n"
            "Теги задньої сторінки:\n"
            "• {{згідно_з_оригіналом}} → «Згідно з оригіналом»\n"
            "• {{номер_наказу}} → номер із назви файла\n"
            "• {{дата_наказу}} → дата з назви файла\n"
            "• {{примірник}} → «Примірник № 2»\n\n"
            "Якщо номер або дату не знайдено в назві файла, відповідний тег лишається "
            "для ручного заповнення. Шаблон має містити рівно одну сторінку.",
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
        for row in rows:
            tree.insert("", tk.END, values=tuple("-" if value in (None, "") else str(value) for value in row))

    def show_analysis_results(self, map_result: dict):
        units_table = map_result.get("units_table")
        calculation_rows = []
        if units_table and hasattr(units_table, "rows"):
            calculation_rows = [(row[0], row[2], row[3]) for row in units_table.rows]
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

    def log(self, message: str):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def log_p2(self, message: str):
        self.p2_log_text.insert(tk.END, message + "\n")
        self.p2_log_text.see(tk.END)
        self.root.update_idletasks()

    def _preview_pause(self, seconds: float) -> None:
        """Пауза режиму превʼю, під час якої вікно лишається живим.

        Звичайний `time.sleep` заморожує цикл подій Tk: журнал не гортається,
        а вікно виглядає зависшим — тобто рівно протилежне до того, задля чого
        режим і потрібен. Тому чекаємо короткими відрізками, а між ними даємо
        Tk перемалюватись. Кнопку запуску на час прогону вимкнено, тож
        повторного входу це не спричиняє.
        """
        deadline = time.monotonic() + max(0.0, seconds)
        while True:
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                break
            time.sleep(min(0.05, remaining))
            self.root.update()

    def clear_log(self):
        self.log_text.delete(1.0, tk.END)

    def copy_log(self, text_widget=None):
        widget = text_widget or self.log_text
        content = widget.get("1.0", tk.END).strip()
        if content:
            self.root.clipboard_clear()
            self.root.clipboard_append(content)
            self.root.update()
            self.log("📋 Текст журналу успішно скопійовано в буфер обміну!")

    def _setup_text_copy_menu(self, text_widget):
        """Додає контекстне меню (ПКМ) та гарячі клавіші для легкого копіювання логів."""
        menu = tk.Menu(text_widget, tearoff=0)
        menu.add_command(label="Копіювати виділене (Ctrl+C)", command=lambda: text_widget.event_generate("<<Copy>>"))
        menu.add_command(label="Виділити все (Ctrl+A)", command=lambda: text_widget.tag_add(tk.SEL, "1.0", tk.END))
        menu.add_separator()
        menu.add_command(label="Скопіювати весь журнал", command=lambda: self.copy_log(text_widget))
        menu.add_command(label="Очистити", command=lambda: text_widget.delete("1.0", tk.END))

        def _popup(event):
            menu.tk_popup(event.x_root, event.y_root)

        text_widget.bind("<Button-3>", _popup)
        text_widget.bind("<Control-a>", lambda e: (text_widget.tag_add(tk.SEL, "1.0", tk.END), "break")[1])
        text_widget.bind("<Control-A>", lambda e: (text_widget.tag_add(tk.SEL, "1.0", tk.END), "break")[1])
        text_widget.bind("<Control-c>", lambda e: text_widget.event_generate("<<Copy>>"))
        text_widget.bind("<Control-C>", lambda e: text_widget.event_generate("<<Copy>>"))

    # ── Вибір файлів ─────────────────────────────────────────────────────────
    def select_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.excel_path.set(path)
            self.save_config()

    def select_doc(self):
        path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx *.doc")])
        if path:
            self.doc_path.set(path)
            output_folder = os.path.join(os.path.dirname(path), "Extracts_Output")
            try:
                os.makedirs(output_folder, exist_ok=True)
                self.out_folder.set(output_folder)
                self.log(f"Папка результату: {output_folder}")
            except OSError as error:
                messagebox.showwarning("Помилка папки", f"Не вдалося створити папку результату:\n{error}")
            self.save_config()
            self._refresh_order_signer()

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
            "{{номер_наказу}}, {{дата_наказу}}, {{кому_список}}, {{куди}}, {{виконавець}}.\n\n"
            "Розмістіть {{кому_список}} у кожному рядку таблиці окремо: один тег "
            "отримує одного унікального адресата. Якщо рядків замало, програма "
            "додає перед наступним блоком копії останнього рядка таблиці — штамп "
            "зсувається вниз і не перекривається. Порядок адресатів: корпуси, "
            "частини, лише обласні ТЦК.\n\n"
            "{{тцк чі вч}} — тип адресата, визначається автоматично:\n"
            "• лише військові частини → «командирам військових частин»;\n"
            "• лише ТЦК → «начальникам ТЦК»;\n"
            "• і те, і те → «командирам військових частин та начальникам ТЦК».\n\n"
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
        for tag, value in replacements.items():
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
        spans = [
            (content_range.Paragraphs(i).Range.Start, content_range.Paragraphs(i).Range.End)
            for i in range(1, content_range.Paragraphs.Count + 1)
        ]

        def paragraph_kind(text: str) -> str:
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

        kinds = []
        for start, end in spans:
            kinds.append(paragraph_kind(doc.Range(start, end).Text))

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
                paragraph_format.KeepWithNext = False
            elif kinds[index] == "heading":
                paragraph_format.KeepWithNext = True
            else:
                # Пункт тримає свій біографічний блок; новий пункт або шапка
                # починають окрему групу.
                paragraph_format.KeepWithNext = kinds[next_meaningful] == "continuation"

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
            for span_start, span_end in find_unmatched_open_unit_spans(ciphered):
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
        if signer.get("name"):
            values["{{підписант_піб}}"] = signer["name"]

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

    def _insert_plain_content(self, doc, tag_range, encrypted_content: str) -> int:
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
        for span_start, span_end in find_unmatched_open_unit_spans(formatted_content):
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
            raise RuntimeError(
                f"Файл «{os.path.basename(output_path)}» відкритий в іншій програмі "
                "(найімовірніше у Word). Закрийте його та повторіть генерацію."
            )

        # Робоча копія шаблону створюється у тимчасовій папці, а не за
        # кінцевим шляхом: так оригінал шаблону лишається недоторканим, а
        # заблокований результат не зриває обробку. Розширення копії
        # відповідає реальному вмісту шаблону (він може бути у форматі
        # Word 97-2003), інакше Word відмовиться її відкрити.
        temp_dir = os.path.join(os.path.dirname(os.path.abspath(output_path)), "_nat_temp")
        os.makedirs(temp_dir, exist_ok=True)
        working_path = copy_template_for_editing(
            template_path, os.path.join(temp_dir, os.path.basename(output_path))
        )
        doc = word.Documents.Open(os.path.abspath(working_path), ReadOnly=False)
        try:
            executor_blocks = self._replace_message_tags(doc, replacements)
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
                            doc, find_obj.Parent, encrypted_content
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
            order_text, _ = text_before_order_signer(source_text)
            order_num, order_date = extract_metadata_from_filename(os.path.basename(order_path))
            # У повідомленнях дата не розкривається словами (на відміну від витягів).
            order_date_formatted = format_message_date(order_date)

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
                f"Невпізнаних назв, виділених жовтим: {highlights}"
                + (f"\nУВАГА: не вмістилося адресатів: {recipient_overflow}" if recipient_overflow else ""),
            )
        except Exception as error:
            self.log(f"ПОМИЛКА повідомлень: {error}")
            messagebox.showerror("Помилка повідомлень", str(error))
        finally:
            if source_doc is not None:
                try:
                    source_doc.Close(False)
                except Exception:
                    pass
            if word:
                force_quit_word(word)
            self.btn_generate_messages.config(state=NORMAL)

    def open_compare_extracts(self):
        """Відкриває вікно порівняння для режиму витягів."""
        DocxCompareWindow(self.root, mode="витяги")

    def open_compare_copies(self):
        """Відкриває вікно порівняння для режиму примірників."""
        gen_path = self.p2_single_file.get() if self.p2_source_mode.get() == "file" else ""
        DocxCompareWindow(self.root, generated_path=gen_path, mode="примірник_2")

    def open_compare_messages(self):
        """Відкриває вікно порівняння для режиму повідомлень."""
        DocxCompareWindow(self.root, mode="повідомлення_зміст")

    def select_p2_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.p2_orders_folder.set(path)
            # Папку результату переобчислюємо під нове джерело, доки
            # користувач не вибрав її вручну.
            if not self.p2_out_folder_manual.get():
                self.p2_out_folder.set(os.path.join(path, "Примірники_2"))
            self.save_config()

    def select_p2_file(self):
        path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx *.doc")])
        if path:
            self.p2_single_file.set(path)
            if not self.p2_out_folder_manual.get():
                self.p2_out_folder.set(os.path.join(os.path.dirname(path), "Примірники_2"))
            self.save_config()

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
        word = None
        doc = None
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0  # wdAlertsNone: не показувати блокуючі діалоги (напр. "Зберегти зміни?")
            doc = word.Documents.Open(os.path.abspath(doc_path), ReadOnly=True)
            return read_document_text(doc)
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
            return
        mark = "☑" if selected else "☐"
        for index, path in enumerate(self.last_copy_two_paths):
            self.copy_two_tree.insert("", tk.END, iid=f"copy2_{index}", values=(mark, path))

    def _toggle_copy_two_source(self, event):
        item = self.copy_two_tree.identify_row(event.y)
        if not item or event.x > 115:
            return
        values = list(self.copy_two_tree.item(item, "values"))
        if len(values) != 2 or values[0] not in ("☐", "☑"):
            return
        values[0] = "☐" if values[0] == "☑" else "☑"
        self.copy_two_tree.item(item, values=values)
        return "break"

    def _selected_order_paths(self) -> tuple[list[str], bool]:
        """Повертає `(шляхи, чи це примірники № 2)`.

        Прапорець потрібен для вибору папки результату: для примірників,
        створених пакетом, витяги кладемо поряд із ними у папку наказу.
        """
        if self.last_copy_two_paths:
            paths = [
                values[1]
                for item in self.copy_two_tree.get_children()
                if (values := self.copy_two_tree.item(item, "values"))
                and len(values) == 2 and values[0] == "☑" and os.path.isfile(values[1])
            ]
            if paths:
                return paths, True
            # Галочки зняті з усіх примірників — працюємо з наказом,
            # обраним вручну, замість того щоб відмовляти користувачу.
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
    def run_rozrahunok_action(self):
        self.save_config()
        order_paths, from_copies = self._selected_order_paths()
        if not self.excel_path.get() or not order_paths:
            messagebox.showwarning(
                "Помилка",
                "Виберіть словник Excel і хоча б один наказ або примірник № 2.",
            )
            return

        self.btn_calc.config(state=DISABLED)
        self.btn_extracts.config(state=DISABLED)
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
            self.btn_calc.config(state=NORMAL)
            self.btn_extracts.config(state=NORMAL)

    def run_extracts_action(self):
        self.save_config()
        order_paths, from_copies = self._selected_order_paths()
        if not self.excel_path.get() or not self.template_path.get() or not order_paths:
            messagebox.showwarning(
                "Помилка",
                "Виберіть словник Excel, шаблон витягу і хоча б один наказ або примірник № 2.",
            )
            return

        self.btn_calc.config(state=DISABLED)
        self.btn_extracts.config(state=DISABLED)
        try:
            failures = self._run_batch(
                order_paths, self.run_extracts, "Витяги для", alongside=from_copies
            )
            self.save_config()
            self._report_batch_result(
                len(order_paths), failures, "Генерація витягів", "генерацію витягів"
            )
        finally:
            self.btn_calc.config(state=NORMAL)
            self.btn_extracts.config(state=NORMAL)

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
                self.log(f"  ПОМИЛКА ({name}): {error}")
                failures.append((name, str(error)))
        return failures

    def _report_batch_result(
        self, total: int, failures: list[tuple[str, str]], title: str, action_name: str
    ) -> None:
        """Показує підсумок пакета: скільки вдалося, а що саме — ні."""
        succeeded = total - len(failures)
        if not failures:
            messagebox.showinfo("Успіх", f"{title}: успішно оброблено {succeeded} з {total} файл(ів).")
            return

        details = "\n".join(f"• {name}: {error}" for name, error in failures[:10])
        if len(failures) > 10:
            details += f"\n… ще {len(failures) - 10}"
        self.log(f"\nНе вдалося виконати {action_name} для {len(failures)} файл(ів).")
        messagebox.showwarning(
            title,
            f"Оброблено {succeeded} з {total} файл(ів).\n\n"
            f"Не вдалося ({len(failures)}):\n{details}",
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

    def run_extracts(self):
        self.log("\n=== ГЕНЕРАЦІЯ ВИТЯГІВ ===")
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
        # Тимчасове правило користувача: у витягах блок підписанта оригіналу
        # наказу повністю вимкнений. Тіло однаково обрізається перед ним, але
        # реквізити не переносяться до шаблону. «Згідно з оригіналом» та
        # засвідчувач нижче лишаються.
        order_signer = {"position": "", "rank": "", "name": ""}
        self.log("Блок підписанта оригіналу у витягах тимчасово вимкнено.")
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

        units_data = map_res.get("unit_paragraphs", {})
        if not units_data:
            self.log("Жодної військової частини не знайдено!")
            messagebox.showwarning("Результат", "Жодної військової частини за словником не знайдено.")
            return

        template_path = os.path.abspath(self.template_path.get())
        out_file = os.path.join(
            self.out_folder.get(), build_extracts_filename(order_num, order_date)
        )
        os.makedirs(self.out_folder.get(), exist_ok=True)
        # Наприкінці генерації результат відкривається у Word, тому при
        # повторному запуску він може бути ще зайнятий. Перевіряємо одразу,
        # щоб не витрачати час на обробку й показати зрозумілу причину.
        if not is_path_writable(out_file):
            raise RuntimeError(
                f"Файл «{os.path.basename(out_file)}» відкритий в іншій програмі "
                "(найімовірніше у Word). Закрийте його та повторіть генерацію."
            )
        temp_dir = os.path.join(self.out_folder.get(), "_nat_temp")
        os.makedirs(temp_dir, exist_ok=True)

        temp_files = []
        layout_warnings = []

        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0  # wdAlertsNone: не показувати блокуючі діалоги (напр. "Зберегти зміни?")

        try:
            source_path = os.path.abspath(self.doc_path.get())
            source_doc = word.Documents.Open(source_path, ReadOnly=True)

            self.log("Індексуємо абзаци оригіналу наказу...")
            para_count = source_doc.Paragraphs.Count
            source_line_to_para = []
            for pi in range(1, para_count + 1):
                raw = source_doc.Paragraphs(pi).Range.Text
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
                    idx = doc.Paragraphs.Count
                    while idx > 1 and is_blank_paragraph(doc.Paragraphs(idx)):
                        doc.Paragraphs(idx).Range.Delete()
                        idx -= 1
                except Exception:
                    pass

            def normalize_signature_gap(signer_start, content_end_pos):
                """Нормалізує відступ перед підписантом: рівно 2 порожні абзаци.

                Після видалення/додавання абзаців позиції зсуваються, тому
                підписант шукається заново по непорожньому тексту після content_end_pos.
                """
                signer_index = None
                for paragraph_index in range(1, doc.Paragraphs.Count + 1):
                    if doc.Paragraphs(paragraph_index).Range.Start == signer_start:
                        signer_index = paragraph_index
                        break
                if signer_index is None:
                    return signer_start

                blank_indexes = []
                for paragraph_index in range(signer_index - 1, 0, -1):
                    paragraph = doc.Paragraphs(paragraph_index)
                    if not is_blank_paragraph(paragraph):
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

                # Після видалення/вставки — позиції зсунулись, шукаємо підписанта заново
                for paragraph_index in range(1, doc.Paragraphs.Count + 1):
                    paragraph = doc.Paragraphs(paragraph_index).Range
                    if paragraph.Start >= content_end_pos and not is_blank_paragraph(paragraph):
                        return paragraph.Start
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

            def apply_keep_rules(item_ranges, heading_ranges, heading_item_pairs, signer_start, executor_start=None):
                # 1. Заголовки (шапки): зв'язуються з наступними абзацами (KeepWithNext)
                for heading_range in heading_ranges:
                    for pi in range(1, heading_range.Paragraphs.Count + 1):
                        pf = heading_range.Paragraphs(pi).Range.ParagraphFormat
                        pf.KeepTogether = True
                        pf.KeepWithNext = True

                # 2. Пункти витягу: кожен пункт цілісний
                for item_index, item_range in enumerate(item_ranges):
                    paragraph_count = item_range.Paragraphs.Count
                    is_last_item = (item_index == len(item_ranges) - 1)
                    for pi in range(1, paragraph_count + 1):
                        pf = item_range.Paragraphs(pi).Range.ParagraphFormat
                        pf.KeepTogether = True
                        pf.KeepWithNext = True if is_last_item else (pi < paragraph_count)

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
                    chain_count = chain_range.Paragraphs.Count
                    for pi in range(1, chain_count + 1):
                        pf = chain_range.Paragraphs(pi).Range.ParagraphFormat
                        pf.KeepTogether = True
                        pf.KeepWithNext = (pi < chain_count)

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
                        for pi in range(1, doc.Paragraphs.Count + 1):
                            pr = doc.Paragraphs(pi).Range
                            if pr.Start <= b_pos < pr.End:
                                return pi
                    except Exception:
                        pass

                # 2. За текстом виконавця (фолбек, якщо закладку втрачено)
                exec_val = self.executor.get().strip()
                if exec_val:
                    lines = [ln.strip() for ln in exec_val.replace("/", "\n").splitlines() if ln.strip()]
                    if lines:
                        first_line = lines[0].lower()
                        for pi in range(doc.Paragraphs.Count, 0, -1):
                            pt = doc.Paragraphs(pi).Range.Text.strip().lower()
                            if first_line in pt or (len(first_line) > 4 and first_line[:4] in pt):
                                return pi

                # 3. Фолбек: останній непорожній абзац у документі
                for pi in range(doc.Paragraphs.Count, 0, -1):
                    if not is_blank_paragraph(doc.Paragraphs(pi)):
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

            def position_executor_at_page_bottom(bookmark_name=None, signer_start=None, needs_manual_review=False):
                """Вирівнює виконавця до низу сторінки, окрім випадків, де це неможливо
                без наставляння штучних порожніх абзаців — тоді лишаємо позицію зі зразка
                (правило 5.5 AGENT.md, виключення №2).

                needs_manual_review навмисно НЕ блокує спробу підштовхування: загальна
                проблема макета деінде у витягу (напр. інший пункт не вміщується) не
                означає, що саме область виконавця не можна коректно розмістити.
                Єдина причина відкату — фізична неможливість (перехід на іншу сторінку),
                що обробляється нижче в самому циклі."""
                executor_index = find_executor_paragraph_index(bookmark_name)
                if executor_index is None:
                    clean_redundant_blanks()
                    return

                # Запам'ятовуємо кількість порожніх абзаців зі зразка перед виконавцем,
                # щоб мати змогу відновити цю позицію, якщо автоматичне вирівнювання
                # донизу не вдасться виконати чисто (без переходу на іншу сторінку).
                original_blank_count = 0
                probe_index = executor_index - 1
                while probe_index > 0 and is_blank_paragraph(doc.Paragraphs(probe_index)):
                    original_blank_count += 1
                    probe_index -= 1

                # Спочатку очищаємо всі наявні порожні абзаци безпосередньо перед виконавцем
                while executor_index > 1 and is_blank_paragraph(doc.Paragraphs(executor_index - 1)):
                    try:
                        doc.Paragraphs(executor_index - 1).Range.Delete()
                        executor_index -= 1
                    except Exception:
                        break

                if executor_index > 1:
                    try:
                        doc.Paragraphs(executor_index - 1).Range.ParagraphFormat.KeepWithNext = False
                    except Exception:
                        pass

                clean_redundant_blanks()
                doc.Repaginate()

                def restore_sample_position(index):
                    """Відновлює рівно original_blank_count порожніх абзаців перед
                    виконавцем (позиція зі зразка) без штучного доштовхування донизу."""
                    while index > 1 and is_blank_paragraph(doc.Paragraphs(index - 1)):
                        try:
                            doc.Paragraphs(index - 1).Range.Delete()
                            index -= 1
                        except Exception:
                            break
                    for _ in range(original_blank_count):
                        pos = doc.Paragraphs(index).Range.Start
                        doc.Range(pos, pos).InsertBefore("\r")
                        index += 1
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
                def insert_enters(count):
                    nonlocal executor_index
                    for _ in range(count):
                        pos = doc.Paragraphs(executor_index).Range.Start
                        doc.Range(pos, pos).InsertBefore("\r")
                        executor_index += 1
                    doc.Repaginate()

                def remove_enters(count):
                    nonlocal executor_index
                    removed = 0
                    while (
                        removed < count
                        and executor_index > 1
                        and is_blank_paragraph(doc.Paragraphs(executor_index - 1))
                    ):
                        doc.Paragraphs(executor_index - 1).Range.Delete()
                        executor_index -= 1
                        removed += 1
                    doc.Repaginate()

                try:
                    start_page = page_of(doc.Paragraphs(executor_index).Range.Start)
                    if start_page is None:
                        self.log(f"  ⚠️ {cipher}: не вдалося визначити сторінку виконавця — лишено як у зразку.")
                    else:
                        added = 0
                        # Фаза 1 — грубе наближення порціями по 4 абзаци
                        # (щоб не робити десятки повільних Repaginate).
                        for _ in range(30):
                            insert_enters(4)
                            added += 4
                            if page_of(doc.Paragraphs(executor_index).Range.Start) != start_page:
                                remove_enters(4)
                                added -= 4
                                break
                        # Фаза 2 — точне доведення по одному абзацу.
                        for _ in range(8):
                            insert_enters(1)
                            added += 1
                            if page_of(doc.Paragraphs(executor_index).Range.Start) != start_page:
                                remove_enters(1)
                                added -= 1
                                break
                        self.log(f"  ↧ {cipher}: виконавця опущено вниз сторінки (+{added} ентер(ів)).")
                except Exception as exec_error:
                    self.log(f"  ⚠️ {cipher}: помилка опускання виконавця: {exec_error}")
                    try:
                        # Повертаємо відступ зі зразка, щоб не лишити виконавця
                        # без жодного порожнього абзацу після невдалої спроби.
                        restore_sample_position(executor_index)
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

            for idx, (cipher, data) in enumerate(units_data.items()):
                self.log(f"[{idx+1}/{len(units_data)}] Генеруємо витяг для: {cipher}")

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
                    template_path, os.path.join(temp_dir, f"extract_{idx:04d}.docx")
                )
                doc = word.Documents.Open(os.path.abspath(temp_path), ReadOnly=False)

                def replace_tag(tag, replacement_text, document=doc, collect_paragraphs=False,
                               highlight_red=False, bold_pattern=None):
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

                def remove_original_signer_template_block():
                    """Видаляє з шаблону весь блок тегів підписанта оригіналу.

                    Якщо тег стоїть у таблиці, видаляється відповідний рядок,
                    щоб після вимкненого блока не лишалися порожні комірки.
                    """
                    signer_tags = (
                        "{{підписант}}",
                        "{{підписант_посада}}",
                        "{{підписант_звання}}",
                        "{{підписант_піб}}",
                    )
                    paragraph_index = doc.Paragraphs.Count
                    while paragraph_index >= 1:
                        paragraph = doc.Paragraphs(paragraph_index).Range
                        paragraph_text = str(paragraph.Text or "")
                        if any(tag.casefold() in paragraph_text.casefold() for tag in signer_tags):
                            try:
                                if paragraph.Information(12):  # wdWithInTable
                                    paragraph.Cells(1).Row.Delete()
                                else:
                                    paragraph.Delete()
                            except Exception:
                                paragraph.Text = ""
                        paragraph_index = min(paragraph_index - 1, doc.Paragraphs.Count)

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
                    replace_tag("{{дата_наказу}}", order_date_formatted,
                                bold_pattern=r"\d+")
                if order_num:
                    # Жирним — лише номер, знак «№» лишається звичайним.
                    replace_tag("{{номер_наказу}}", f"№{order_num}",
                                bold_pattern=r"(?<=№).+")

                # Підписант оригіналу наказу тимчасово не переноситься.
                remove_original_signer_template_block()

                # Особа, яка засвідчує витяг («Згідно з оригіналом» / Засвідчувач)
                replace_tag("{{згідно_з_оригіналом}}", "Згідно з оригіналом")
                replace_tag("{{засвідчення}}", "Згідно з оригіналом")
                if self.certifier_position.get().strip():
                    cert_pos = _slash_to_lines(self.certifier_position.get().strip())
                    replace_tag("{{засвідчувач_посада}}", cert_pos)
                    replace_tag("{{згідно_з_оригіналом_посада}}", cert_pos)
                if self.certifier_rank.get().strip():
                    cert_rank = self.certifier_rank.get().strip()
                    replace_tag("{{засвідчувач_звання}}", cert_rank)
                    replace_tag("{{згідно_з_оригіналом_звання}}", cert_rank)
                if self.certifier_name.get().strip():
                    cert_name = self.certifier_name.get().strip()
                    replace_tag("{{засвідчувач_піб}}", cert_name)
                    replace_tag("{{згідно_з_оригіналом_піб}}", cert_name)
                    replace_tag("{{засвідчувач}}", cert_name)

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

                    def insert_source_span(start_line, end_line, fallback_text, kind):
                        nonlocal insert_point
                        paragraph_indexes = line_span_to_paragraphs(start_line, end_line)
                        # Обрізаємо порожні абзаци на початку та в кінці діапазону, але зберігаємо внутрішні ентери
                        while paragraph_indexes and is_blank_paragraph(source_doc.Paragraphs(paragraph_indexes[0])):
                            paragraph_indexes.pop(0)
                        while paragraph_indexes and is_blank_paragraph(source_doc.Paragraphs(paragraph_indexes[-1])):
                            paragraph_indexes.pop()

                        first_start = None
                        last_end = None
                        for paragraph_index in paragraph_indexes:
                            source_paragraph = source_doc.Paragraphs(paragraph_index)

                            if has_manual_page_break(source_paragraph):
                                self.log("Пропущено вихідний розрив сторінки; пагінація витягу буде побудована заново.")
                                continue
                            source_range = source_paragraph.Range.Duplicate
                            start = insert_point
                            destination_range = doc.Range(start, start)
                            destination_range.FormattedText = source_range.FormattedText
                            insert_point = destination_range.End

                            # Геометрію переносимо ЯВНО, а не покладаємось на
                            # FormattedText. Word не записує властивість, яка
                            # дорівнює типовій (Alignment=Left, FirstLineIndent=0),
                            # тому такий абзац у витягу успадковував стиль
                            # `Normal` ШАБЛОНА. Якщо в шаблоні стоїть «за
                            # шириною» та відступ 1.25 см — біографічний блок
                            # наказу (left 8 см, перший рядок 0, уліво) з'їжджав
                            # і розтягувався по ширині, і так само «плив» § .
                            # ParagraphFormat діапазону віддає ДІЮЧІ значення,
                            # тож так витяг повторює наказ незалежно від шаблону.
                            source_format = source_paragraph.Range.ParagraphFormat
                            geometry = {}
                            for prop in ("Alignment", "LeftIndent", "RightIndent", "FirstLineIndent"):
                                try:
                                    geometry[prop] = getattr(source_format, prop)
                                except Exception:
                                    pass
                            for dest_pi in range(1, doc.Range(start, insert_point).Paragraphs.Count + 1):
                                dest_format = doc.Range(start, insert_point).Paragraphs(dest_pi).Range.ParagraphFormat
                                dest_format.PageBreakBefore = False
                                for prop, value in geometry.items():
                                    try:
                                        setattr(dest_format, prop, value)
                                    except Exception:
                                        pass
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

                    for item in items:
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
                                heading_range = insert_source_span(
                                    heading_key[0], heading_key[1], item.get("parent_heading", ""), "заголовка"
                                )
                                if heading_range:
                                    inserted_heading_ranges.append(heading_range)
                                    new_heading_ranges.append(heading_range)
                                    # §, основна шапка та підшапка — окремі
                                    # елементи; після кожного лишається відступ.
                                    insert_empty_paragraph()
                        copied_heading_keys = heading_keys

                        item_range = insert_source_span(
                            item.get("source_start_line"), item.get("source_end_line"),
                            item.get("original_text") or item.get("text", ""), "пункту"
                        )
                        if item_range:
                            inserted_item_ranges.append(item_range)
                            item_label = item.get("label", "пункт")
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
                    upper = doc.Content.End
                    exec_idx = find_executor_paragraph_index(executor_bookmark)
                    if exec_idx:
                        try:
                            upper = doc.Paragraphs(exec_idx).Range.Start
                        except Exception:
                            upper = doc.Content.End
                    return content_start_pos, upper

                def apply_exact_line_spacing(points, blanks_too=True):
                    """Виставляє точний міжрядковий інтервал (wdLineSpaceExactly) в pt
                    виключно на абзаци вставленого змісту та підписанта."""
                    lower, upper = spacing_bounds()
                    if lower is None:
                        return
                    for p_idx in range(1, doc.Paragraphs.Count + 1):
                        p = doc.Paragraphs(p_idx)
                        p_range = p.Range
                        if p_range.Start < lower or p_range.Start >= upper:
                            continue
                        # Абзац із вбудованим зображенням ніколи не отримує точний
                        # інтервал — інакше Word обріже картинку.
                        try:
                            if p_range.InlineShapes.Count > 0:
                                continue
                        except Exception:
                            pass
                        if not blanks_too and is_blank_paragraph(p):
                            continue
                        p_fmt = p_range.ParagraphFormat
                        p_fmt.LineSpacingRule = 4  # wdLineSpaceExactly
                        p_fmt.LineSpacing = points
                        p_fmt.SpaceBefore = 0
                        p_fmt.SpaceAfter = 0

                # 1. Автопідбір для вміщення на 1 сторінку: якщо через невеликий
                # перебір вийшло 2 сторінки, шукаємо в діапазоні 16 → 14 пт
                # (точний міжрядковий інтервал, крок 0.5 пт) найбільше значення,
                # що ще вміщує весь витяг на 1 сторінку — замість жорсткого
                # одинарного інтервалу, який не завжди рятує.
                if pages_count == 2:
                    clean_redundant_blanks()
                    fitted_spacing = None
                    spacing = 16.0
                    while spacing >= 14.0 - 1e-6:
                        apply_exact_line_spacing(spacing)
                        doc.Repaginate()
                        if doc.ComputeStatistics(2) == 1:
                            fitted_spacing = spacing
                            break
                        spacing -= 0.5
                    if fitted_spacing is not None:
                        pages_count = 1
                        self.log(f"  ℹ️ {cipher}: точний інтервал {fitted_spacing} пт — вміщено на 1 сторінку.")
                    else:
                        # Навіть 14 пт не допомогло — це реальний перебір контенту,
                        # лишаємо 2 сторінки (потребує уваги, а не силового стиснення).
                        apply_exact_line_spacing(16.0)
                        doc.Repaginate()
                        pages_count = doc.ComputeStatistics(2)

                # 1б. БОНУС для БАГАТОСТОРІНКОВИХ витягів: не змінюючи кількості
                # сторінок, добираємо інтервал так, щоб на ПЕРШУ сторінку сіло
                # якнайбільше пунктів. Без цього інтервал лишався 16 пт —
                # найбільший, — і на першій сторінці часто стояв один пункт,
                # тоді як 2, 3, 4, 5 разом вміщувалися на другій. Кількість
                # сторінок не погіршуємо ніколи: це саме бонус, а не стиснення.
                if pages_count > 1 and inserted_item_ranges:

                    def items_on_first_page():
                        count = 0
                        for candidate in inserted_item_ranges:
                            start_page, end_page = range_pages(candidate)
                            if start_page == 1 and end_page == 1:
                                count += 1
                        return count

                    best_pages = pages_count
                    best_on_first = items_on_first_page()
                    best_multi_spacing = 16.0
                    spacing = 16.0 - 0.5
                    while spacing >= 14.0 - 1e-6:
                        apply_exact_line_spacing(spacing)
                        doc.Repaginate()
                        pages_now = doc.ComputeStatistics(2)
                        if pages_now <= best_pages:
                            on_first = items_on_first_page()
                            if pages_now < best_pages or on_first > best_on_first:
                                best_pages = pages_now
                                best_on_first = on_first
                                best_multi_spacing = spacing
                        spacing -= 0.5
                    apply_exact_line_spacing(best_multi_spacing)
                    doc.Repaginate()
                    pages_count = doc.ComputeStatistics(2)
                    if best_multi_spacing != 16.0:
                        self.log(
                            f"  ℹ️ {cipher}: інтервал {best_multi_spacing} пт — на першій "
                            f"сторінці пунктів: {best_on_first} (сторінок: {pages_count})."
                        )

                # 2. Правило «Максимально заповнена сторінка»: якщо витяг вміщується
                # на 1 сторінці, шукаємо в тому самому діапазоні 14 → 16 пт
                # НАЙБІЛЬШЕ значення точного інтервалу, що все ще утримує весь
                # вміст на 1 сторінці — для максимального гармонійного заповнення
                # (замість фіксованого +2 пт, який на практиці не заповнював сторінку).
                if pages_count == 1:
                    best_spacing = None
                    spacing = 16.0
                    while spacing >= 14.0 - 1e-6:
                        apply_exact_line_spacing(spacing, blanks_too=False)
                        doc.Repaginate()
                        if doc.ComputeStatistics(2) == 1:
                            best_spacing = spacing
                            break
                        spacing -= 0.25
                    if best_spacing is None:
                        apply_exact_line_spacing(14.0, blanks_too=False)
                        doc.Repaginate()

                # 3. Остаточна перевірка макета — ПІСЛЯ підбору міжрядкового
                # інтервалу (16→14 пт), який часто сам усуває розбіжність
                # (відірвану шапку, перебір на 2 сторінки тощо). Лише якщо
                # проблема лишається й після цього — витяг справді потребує
                # ручної перевірки, і виконавця НЕ підштовхуємо штучно.
                if inserted_item_ranges:
                    remaining_issues = layout_issues(
                        inserted_item_ranges, inserted_item_labels, heading_item_pairs, signer_start, executor_start
                    )
                    for issue in remaining_issues:
                        layout_warnings.append(f"{cipher}: {issue}.")
                    if remaining_issues:
                        extract_needs_manual_review = True

                # 4. Виконавець ЗАВЖДИ вирівнюється строго до низу поточної сторінки,
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
                temp_files.append((temp_path, pages_count, cipher))

            source_doc.Close(False)

            # Збираємо всі витяги в один фінальний документ
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
                    """ФАКТИЧНА кількість сторінок зібраного документа."""
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
                        index = target_doc.Paragraphs.Count
                        while index > 1:
                            paragraph = target_doc.Paragraphs(index).Range
                            if (paragraph.Text or "").strip(chr(13) + chr(7) + chr(11) + chr(12) + chr(32) + chr(9)):
                                break
                            paragraph.Delete()
                            index = min(index - 1, target_doc.Paragraphs.Count)
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
                    # до тієї, старої, ламала вирівнювання.
                    actual_pages = max(1, current_doc_pages - pages_before)
                    if actual_pages != t_pages:
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

                last_para = target_doc.Paragraphs(target_doc.Paragraphs.Count)
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
            stats_msg = chr(10).join([stats_msg] + missed_lines)

            self.log(f"Збережено файл: {out_file}")
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

            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir, ignore_errors=True)

            self.log("Відкриваємо згенерований документ...")
            final_doc = word.Documents.Open(os.path.abspath(out_file))
            word.Visible = True
            final_doc.Activate()
            word = None

        finally:
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir, ignore_errors=True)
            if word:
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

        order_files: list[str] = []
        if self.p2_source_mode.get() == "folder":
            folder = self.p2_orders_folder.get()
            if not folder or not os.path.exists(folder):
                messagebox.showwarning("Помилка", "Виберіть папку з наказами!")
                return
            for fname in sorted(os.listdir(folder)):
                candidate = os.path.join(folder, fname)
                if (
                    fname.lower().endswith(".docx")
                    and not fname.startswith("~$")
                    and not is_generated_copy_filename(fname)
                    and os.path.normcase(os.path.abspath(candidate)) != os.path.normcase(os.path.abspath(back_page))
                ):
                    order_files.append(candidate)
        else:
            single_file = self.p2_single_file.get()
            if not single_file or not os.path.exists(single_file):
                messagebox.showwarning("Помилка", "Виберіть файл наказу!")
                return
            order_files.append(single_file)

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
            # «Задня сторінка» завжди є окремим односторінковим шаблоном.
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
            signer_tag_in_template = "{{підписант}}" in template_text
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

                    # Засвідчувач («Згідно з оригіналом») — ті самі поля, що
                    # й у витягах, щоб дані не розходились між вкладками.
                    cert_pos = _slash_to_lines(self.certifier_position.get().strip())
                    cert_rank = self.certifier_rank.get().strip()
                    cert_name = self.certifier_name.get().strip()
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
                        back_page_abs, os.path.join(target_dir, "_nat_tmpl.docx")
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
                    self.log_p2(f"  ПОМИЛКА ({fname}): {order_error}")
                    failed_orders.append((fname, str(order_error)))
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
                details = "\n".join(f"• {name}: {error}" for name, error in failed_orders[:10])
                if len(failed_orders) > 10:
                    details += f"\n… ще {len(failed_orders) - 10}"
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
            self.log_p2(f"\n❌ ПОМИЛКА пакетної генерації: {error}")
            for rec in created_records:
                self.p2_tree.insert("", tk.END, values=rec)
            if created_records:
                self._set_copy_two_sources([record[-1] for record in created_records])
            messagebox.showerror(
                "Помилка формування примірників",
                f"Пакет перервано: {error}\n\n"
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
        self.notebook.select(self.tab_extracts)
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
    app_window = tb.Window(themename="cosmo")
    app = App(app_window)
    app_window.mainloop()
