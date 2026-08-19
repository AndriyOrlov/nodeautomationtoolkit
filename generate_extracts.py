import os
import sys
import json
import re
import shutil
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
from nodeautomationtoolkit.builtin_nodes.message_order import generate_decision_order
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
        values["{{номер_наказу}}"] = order_num
    if order_date:
        values["{{дата_наказу}}"] = format_ukr_date(order_date) or order_date
    return values


def build_copy_two_filename(order_num: str, order_date: str, source_filename: str) -> str:
    """Формує безпечну назву примірника № 2 без вигаданих реквізитів."""
    if order_num and order_date:
        safe_number = sanitize_filename(order_num)
        safe_date = sanitize_filename(order_date)
        return f"прим_2_{safe_date}_{safe_number}.docx"
    return f"прим_2_{os.path.basename(source_filename)}"


def apply_ukrainian_typography(text: str) -> str:
    """Замінює пробіли після коротких прийменників, сполучників та скорочень на нерозривні (\u00A0)."""
    pattern = r"(?i)\b(з|із|зі|та|до|в|у|на|і|й|по|за|від|при|під|над|про|для|без|через|шпк|вос-?\d*|зс|р\.н\.|в/ч|в\.ч\.|№|п\.|пп\.|ст\.)\s+"
    return re.sub(pattern, lambda m: f"{m.group(1)}\u00A0", text or "")


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
    if "р.н." in t or "року народження" in t:
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

_MESSAGE_OPEN_INFORMATION_LINES = (
    "ВІДКРИТА ІНФОРМАЦІЯ",
    "(Обмежено в розповсюдженні – лише для Збройних Сил України)",
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


def build_message_recipient_list(mapping: dict, routes: dict) -> list[str]:
    """Список для {{кому_список}} без втрати підпорядкування.

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
    return corps_recipients + unit_recipients + tck_recipients


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


def _find_order_signer(text: str) -> dict[str, str] | None:
    """Повертає реквізити й номер рядка початку підписанта (Командувача) в наказі, відсікаючи таблицю розсилки."""
    raw_lines = str(text or "").replace("\x07", "").splitlines()
    lines = [re.sub(r"\s+", " ", line).strip() for line in raw_lines]

    # Знаходимо межу відсікання службової таблиці/розсилки, шукаючи з кінця документа
    # щоб не обрізати наказ передчасно на випадковому входженні маркера в тілі пункту
    reference_index = len(lines)
    scan_limit = max(0, len(lines) - 120)
    for idx in range(len(lines) - 1, scan_limit - 1, -1):
        line_lower = lines[idx].casefold().strip()
        if any(line_lower.startswith(marker) or marker == line_lower for marker in _DISTRIBUTION_CUTOFF_MARKERS):
            reference_index = idx
            break

    search_start = max(0, reference_index - 80)

    for start_index in range(reference_index - 1, search_start - 1, -1):
        if not _ORDER_SIGNER_START_RE.match(lines[start_index]):
            continue

        position_lines = []
        for line_index in range(start_index, min(reference_index, start_index + 8)):
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


def text_before_order_signer(text: str) -> tuple[str, dict[str, str]]:
    """Відсікає підписанта й увесь службовий текст, який іде після нього."""
    signer = _find_order_signer(text)
    if not signer:
        return text, {"position": "", "rank": "", "name": ""}
    raw_lines = str(text or "").splitlines()
    start_line = signer["start_line"]
    clean_signer = {key: signer.get(key, "") for key in ("position", "rank", "name")}
    return "\n".join(raw_lines[:start_line]).rstrip(), clean_signer


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
        self.p2_copy_title = tk.StringVar(value="Примірник № 2")
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
                    # Правила для примірника № 3 ще не погоджені.
                    self.p2_copy_title.set("Примірник № 2")

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
            "p2_copy_title": self.p2_copy_title.get(),
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

        copy_type_box = tb.Frame(copies_card)
        copy_type_box.grid(row=5, column=0, columnspan=3, sticky=W, pady=(4, 0))
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
        ).grid(row=6, column=0, columnspan=3, sticky=W, pady=(4, 0))

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
                "Стандартні теги: {{номер_наказу}}, {{дата_наказу}}, {{кому_список}}, {{куди}}, {{виконавець}}. "
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
            "У шаблоні зі змістом:\n"
            "{{зміст_шифр}} — пункти наказу із заміною відкритих назв частин "
            "на шифри з Excel. Назви, які лишилися без збігу, виділяються жовтим.\n\n"
            "Після {{виконавець}} програма додає рядки «ВІДКРИТА ІНФОРМАЦІЯ» та "
            "«(Обмежено в розповсюджені – лише для Збройних Сил України)»; весь "
            "блок розміщується за 3 см від низу сторінки, якщо тег поза таблицею. "
            "Будь-який текст шаблону після {{виконавець}} видаляється.\n\n"
            "Номер і дата беруться лише з назви файла наказу."
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
                replacement = value
                if tag == "{{виконавець}}":
                    replacement = "\r".join((value, *_MESSAGE_OPEN_INFORMATION_LINES))
                find_obj.Parent.Text = replacement
                if tag == "{{виконавець}}":
                    # Блок виконавця є кінцем повідомлення. Прибираємо старий
                    # дрібний службовий хвіст шаблону після нього.
                    tail_start = replacement_start + len(replacement)
                    tail_end = max(tail_start, document.Content.End - 1)
                    if tail_end > tail_start:
                        document.Range(tail_start, tail_end).Delete()
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

    def _create_message_file(
        self,
        word,
        template_path: str,
        output_path: str,
        replacements: dict[str, str],
        encrypted_content: str = "",
        recipients: list[str] | None = None,
    ) -> tuple[int, int, int]:
        # Створюємо копію шаблону за цільовим шляхом, щоб ніколи не відкривати і не змінювати оригінал шаблону
        shutil.copy2(template_path, output_path)
        doc = word.Documents.Open(os.path.abspath(output_path), ReadOnly=False)
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
            if encrypted_content:
                find_obj = doc.Content.Find
                find_obj.Text = "{{зміст_шифр}}"
                if find_obj.Execute():
                    content_start = find_obj.Parent.Start
                    # 1. Очищення дублікатів, ентери перед пунктами та застосування нерозривних пробілів (типографіка)
                    cleaned_content = clean_duplicated_units(encrypted_content)
                    separated_content = ensure_blank_line_before_items(cleaned_content)
                    formatted_content = apply_ukrainian_typography(separated_content)
                    find_obj.Parent.Text = formatted_content
                    content_end = content_start + len(formatted_content)
                    content_range = doc.Range(content_start, content_end)

                    # 2. Форматування абзаців змісту
                    for pi in range(1, content_range.Paragraphs.Count + 1):
                        p = content_range.Paragraphs(pi)
                        p_text = p.Range.Text.strip()
                        if not p_text:
                            continue
                        p_format = p.Range.ParagraphFormat
                        if is_biographical_paragraph(p_text):
                            # Біографічні блоки (дата народження, освіта, служба, РНОКПП) — по середині аркуша
                            p_format.Alignment = 1  # wdAlignParagraphCenter
                            p_format.LeftIndent = 0
                            p_format.RightIndent = 0
                            p_format.FirstLineIndent = 0
                            p_format.SpaceBefore = 0
                            p_format.SpaceAfter = 0
                            p_format.LineSpacingRule = 0  # wdLineSpaceSingle
                        elif re.match(r"^\d{1,3}[\.\)]", p_text):
                            # Пункти наказу — вирівнювання по ширині, абзацний відступ 1.25 см
                            p_format.Alignment = 3  # wdAlignParagraphJustify
                            p_format.LeftIndent = 0
                            p_format.RightIndent = 0
                            p_format.FirstLineIndent = 35.45  # 1.25 cm
                            p_format.SpaceBefore = 6
                            p_format.SpaceAfter = 0
                        elif p_text.startswith("Відповідно до") or "ЗВІЛЬНИТИ" in p_text.upper() or "ПРИЗНАЧИТИ:" in p_text.upper():
                            p_format.Alignment = 3  # wdAlignParagraphJustify
                            p_format.LeftIndent = 0
                            p_format.RightIndent = 0
                            p_format.FirstLineIndent = 35.45
                            p_format.SpaceBefore = 6
                            p_format.SpaceAfter = 6
                        elif "Призначається на" in p_text or "шпк" in p_text.lower():
                            p_format.Alignment = 3  # wdAlignParagraphJustify
                            p_format.LeftIndent = 0
                            p_format.RightIndent = 0
                            p_format.FirstLineIndent = 35.45
                            p_format.SpaceBefore = 0
                            p_format.SpaceAfter = 6

                    # 3. Підсвічування відкритих назв ВЧ, які лишилися без шифрування
                    for start, end in find_unmatched_open_unit_spans(formatted_content):
                        doc.Range(content_start + start, content_start + end).HighlightColorIndex = 7  # wdYellow
                        highlighted_count += 1
                else:
                    self.log("УВАГА: у шаблоні зі змістом не знайдено тег {{зміст_шифр}}.")
            for bookmark_name in executor_bookmarks:
                self._position_message_executor_at_page_bottom(doc, bookmark_name)
                if doc.Bookmarks.Exists(bookmark_name):
                    doc.Bookmarks(bookmark_name).Delete()
            doc.SaveAs2(os.path.abspath(output_path), 16)
            return highlighted_count, recipient_slots, recipient_overflow
        finally:
            doc.Close(False)

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
        try:
            mapping = read_recipient_mapping(path=self.excel_path.get()).get("mapping", {})
            source_text = self._read_word_text(order_path)
            order_text, _ = text_before_order_signer(source_text)
            order_num, order_date = extract_metadata_from_filename(os.path.basename(order_path))
            order_date_formatted = format_ukr_date(order_date)

            routes = map_military_units(text=order_text, mapping=mapping)
            recipients = build_message_recipient_list(mapping, routes)
            destinations = []
            for data in routes.get("unit_paragraphs", {}).values():
                destination = str(data.get("destination_where") or "").strip()
                if destination and destination not in destinations:
                    destinations.append(destination)

            replacements = {}
            if order_num:
                replacements["{{номер_наказу}}"] = f"№ {order_num}"
            if order_date_formatted:
                replacements["{{дата_наказу}}"] = order_date_formatted
            if destinations:
                replacements["{{куди}}"] = "\r".join(destinations)
            if self.message_executor.get().strip():
                replacements["{{виконавець}}"] = self.message_executor.get().strip()

            decision = generate_decision_order(text=order_text, mapping=mapping, new_header="")
            encrypted_content = decision.get("decision_text", "")
            safe_number = re.sub(r'[\\/:*?"<>|]', "_", order_num) if order_num else ""
            suffix = f"_№{safe_number}" if safe_number else ""
            cover_output = os.path.join(out_folder, f"Повідомлення_супровід{suffix}.docx")
            content_output = os.path.join(out_folder, f"Повідомлення_шифрований_зміст{suffix}.docx")

            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0  # wdAlertsNone: не показувати блокуючі діалоги (напр. "Зберегти зміни?")
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
            if not self.p2_out_folder.get():
                self.p2_out_folder.set(os.path.join(path, "Примірники_2"))
            self.save_config()

    def select_p2_file(self):
        path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx *.doc")])
        if path:
            self.p2_single_file.set(path)
            if not self.p2_out_folder.get():
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
            self.save_config()

    def _read_word_text(self, doc_path: str) -> str:
        word = None
        doc = None
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0  # wdAlertsNone: не показувати блокуючі діалоги (напр. "Зберегти зміни?")
            doc = word.Documents.Open(os.path.abspath(doc_path), ReadOnly=True)
            text = doc.Content.Text
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

    def _selected_order_paths(self) -> list[str]:
        """Повертає позначені примірники № 2 або вручну обраний наказ."""
        if self.last_copy_two_paths:
            paths = [
                values[1]
                for item in self.copy_two_tree.get_children()
                if (values := self.copy_two_tree.item(item, "values"))
                and len(values) == 2 and values[0] == "☑" and os.path.isfile(values[1])
            ]
            return paths
        return [self.doc_path.get()] if self.doc_path.get() and os.path.isfile(self.doc_path.get()) else []

    def _set_processing_order(self, doc_path: str):
        self.doc_path.set(doc_path)
        self.out_folder.set(os.path.join(os.path.dirname(os.path.abspath(doc_path)), "Extracts_Output"))
        os.makedirs(self.out_folder.get(), exist_ok=True)

    # =========================================================================
    # ДІЇ: РОЗРАХУНОК ТА ВИТЯГИ
    # =========================================================================
    def run_rozrahunok_action(self):
        self.save_config()
        order_paths = self._selected_order_paths()
        if not self.excel_path.get() or not order_paths:
            messagebox.showwarning(
                "Помилка",
                "Виберіть словник Excel і хоча б один наказ або примірник № 2.",
            )
            return

        self.btn_calc.config(state=DISABLED)
        self.btn_extracts.config(state=DISABLED)
        try:
            for index, order_path in enumerate(order_paths, start=1):
                self._set_processing_order(order_path)
                self.log(f"\n[{index}/{len(order_paths)}] Розрахунок для: {os.path.basename(order_path)}")
                self.run_rozrahunok()
            self.save_config()
            messagebox.showinfo("Успіх", f"Розрахунок розсилки сформовано для {len(order_paths)} файл(ів).")
        except Exception as e:
            self.log(f"ПОМИЛКА: {str(e)}")
            import traceback
            tb_text = traceback.format_exc()
            traceback.print_exc()
            messagebox.showerror(
                "Помилка розрахунку розсилки",
                f"Під час розрахунку сталася помилка:\n\n{e}\n\n{tb_text[-1500:]}",
            )
        finally:
            self.btn_calc.config(state=NORMAL)
            self.btn_extracts.config(state=NORMAL)

    def run_extracts_action(self):
        self.save_config()
        order_paths = self._selected_order_paths()
        if not self.excel_path.get() or not self.template_path.get() or not order_paths:
            messagebox.showwarning(
                "Помилка",
                "Виберіть словник Excel, шаблон витягу і хоча б один наказ або примірник № 2.",
            )
            return

        self.btn_calc.config(state=DISABLED)
        self.btn_extracts.config(state=DISABLED)
        try:
            for index, order_path in enumerate(order_paths, start=1):
                self._set_processing_order(order_path)
                self.log(f"\n[{index}/{len(order_paths)}] Витяги для: {os.path.basename(order_path)}")
                self.run_extracts()
            self.save_config()
        except Exception as e:
            self.log(f"ПОМИЛКА: {str(e)}")
            import traceback
            tb_text = traceback.format_exc()
            traceback.print_exc()
            messagebox.showerror(
                "Помилка генерації витягів",
                f"Під час генерації витягів сталася помилка:\n\n{e}\n\n{tb_text[-1500:]}",
            )
        finally:
            self.btn_calc.config(state=NORMAL)
            self.btn_extracts.config(state=NORMAL)

    def run_rozrahunok(self):
        self.log("\n=== РОЗРАХУНОК РОЗСИЛКИ ===")
        self.log(f"Читаємо словник: {self.excel_path.get()}")
        excel_res = read_recipient_mapping(path=self.excel_path.get())
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
        excel_res = read_recipient_mapping(path=self.excel_path.get())
        mapping = excel_res.get("mapping", {})
        unique_entries = {id(value): value for value in mapping.values() if isinstance(value, dict)}
        abbreviation_count = sum(bool(str(value.get("abbreviation", "")).strip()) for value in unique_entries.values())
        self.log(
            f"Зчитано {len(unique_entries)} записів з Excel. "
            f"Скорочень з колонки C: {abbreviation_count}."
        )

        source_text = self._read_word_text(self.doc_path.get())
        text, order_signer = self._refresh_order_signer(source_text)
        # Якщо користувач вручну відредагував поля підписанта в GUI — використати їх
        gui_pos = self.order_signer_position.get().strip()
        gui_rank = self.order_signer_rank.get().strip()
        gui_name = self.order_signer_name.get().strip()
        if gui_pos or gui_rank or gui_name:
            order_signer = {
                "position": gui_pos.replace(" / ", "\n") if gui_pos else order_signer.get("position", ""),
                "rank": gui_rank if gui_rank else order_signer.get("rank", ""),
                "name": gui_name if gui_name else order_signer.get("name", ""),
            }
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

        for invalid_link in map_res.get("invalid_corps_links", []):
            self.log(
                "УВАГА: корпус не знайдено окремим рядком у таблиці; "
                f"частину залишено самостійним адресатом: {invalid_link.get('unit', '')} → {invalid_link.get('corps', '')}"
            )
        for tck_reference in map_res.get("unresolved_tck_references", []):
            self.log(f"УВАГА: ТЦК не визначено або його ОТЦК відсутній у таблиці; витяг не буде створено: {tck_reference}")

        routing_report = os.path.join(self.out_folder.get(), "Контроль_маршрутизації.xlsx")
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
                temp_path = os.path.join(temp_dir, f"extract_{idx:04d}.docx")
                # Робимо копію шаблону, щоб Word COM ніколи не відкривав і не змінював оригінальний файл шаблону
                shutil.copy2(template_path, temp_path)
                doc = word.Documents.Open(os.path.abspath(temp_path), ReadOnly=False)

                def replace_tag(tag, replacement_text, document=doc, collect_paragraphs=False, highlight_red=False):
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
                        if collect_paragraphs:
                            replaced_paragraphs.append(
                                document.Range(found_start, found_start).Paragraphs(1).Range.Duplicate
                            )
                        find_obj = document.Content.Find
                        find_obj.Text = tag
                    return replaced_paragraphs

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
                    replace_tag("{{дата_наказу}}", order_date_formatted)
                if order_num:
                    replace_tag("{{номер_наказу}}", f"№{order_num}")

                def _slash_to_lines(text_val: str) -> str:
                    """Конвертує слеші ' / ' або '/' (крім 'в/ч' та дат) у переноси рядків для Word."""
                    if not text_val:
                        return ""
                    t = re.sub(r"\b([вВ])\s*/\s*([чЧ])\b", r"\1_SLASH_TEMP_\2", str(text_val))
                    t = re.sub(r"(\d)\s*/\s*(\d)", r"\1_NUMSLASH_TEMP_\2", t)
                    t = re.sub(r"\s*/\s*", "\r", t)
                    t = t.replace("_SLASH_TEMP_", "/")
                    t = t.replace("_NUMSLASH_TEMP_", "/")
                    return t.replace("\n", "\r")

                # Підписант оригіналу наказу (Командувач/Командир)
                if order_signer["position"]:
                    replace_tag("{{підписант_посада}}", _slash_to_lines(order_signer["position"]))
                if order_signer["rank"]:
                    replace_tag("{{підписант_звання}}", order_signer["rank"])
                if order_signer["name"]:
                    replace_tag("{{підписант_піб}}", order_signer["name"])

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

                    def is_birth_date_paragraph(paragraph):
                        paragraph_text = paragraph.Range.Text.casefold()
                        return (
                            bool(re.search(r"\bр\s*\.\s*н\s*\.", paragraph_text))
                            or "року народження" in paragraph_text
                        )

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
                            for dest_pi in range(1, doc.Range(start, insert_point).Paragraphs.Count + 1):
                                doc.Range(start, insert_point).Paragraphs(dest_pi).Range.ParagraphFormat.PageBreakBefore = False
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
                    shutil.copy2(first_path, out_file)
            else:
                target_doc = word.Documents.Open(os.path.abspath(first_path))
                current_doc_pages = first_pages

                if enable_2up and first_pages > 1 and (first_pages % 2 != 0):
                    rng = target_doc.Content
                    rng.Collapse(0)
                    rng.InsertBreak(2)  # wdSectionBreakNextPage
                    current_doc_pages += 1
                    self.log(f"Додано порожню сторінку після витягу {first_cipher} ({first_pages} стор.) для вирівнювання аркуша.")

                for i in range(1, len(temp_files)):
                    t_path, t_pages, t_cipher = temp_files[i]

                    # Якщо увімкнено друк 2 на 1: багатосторінковий витяг (>1 стор.) обов'язково починається з нового фізичного аркуша (з непарної сторінки)
                    if enable_2up and t_pages > 1 and (current_doc_pages % 2 != 0):
                        rng = target_doc.Content
                        rng.Collapse(0)
                        rng.InsertBreak(2)  # wdSectionBreakNextPage
                        current_doc_pages += 1
                        self.log(f"Додано порожню сторінку перед багатосторінковим витягом {t_cipher} ({t_pages} стор.), щоб він почався з нового аркуша.")

                    rng = target_doc.Content
                    rng.Collapse(0)
                    rng.InsertBreak(2)  # wdSectionBreakNextPage

                    rng = target_doc.Content
                    rng.Collapse(0)
                    rng.InsertFile(os.path.abspath(t_path))
                    # Оновлюємо лічильник реальною кількістю сторінок, а не ручною арифметикою
                    target_doc.Repaginate()
                    current_doc_pages = target_doc.ComputeStatistics(2)  # wdStatisticPages

                    # Якщо цей багатосторінковий витяг непарний (наприклад, 3 сторінки), додаємо пусту сторінку після нього:
                    if enable_2up and t_pages > 1 and (t_pages % 2 != 0):
                        rng = target_doc.Content
                        rng.Collapse(0)
                        rng.InsertBreak(2)  # wdSectionBreakNextPage
                        current_doc_pages += 1
                        self.log(f"Додано порожню сторінку після багатосторінкового витягу {t_cipher} ({t_pages} стор.) для вирівнювання наступного аркуша.")

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
                    and "прим_" not in fname.lower()
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
        if not out_root:
            base_dir = self.p2_orders_folder.get() if self.p2_source_mode.get() == "folder" else os.path.dirname(order_files[0])
            out_root = os.path.join(base_dir, "Примірники_2")
            self.p2_out_folder.set(out_root)
        os.makedirs(out_root, exist_ok=True)
        self.save_config()

        self.btn_run_p2.config(state=DISABLED)
        self.p2_log_text.delete(1.0, tk.END)
        for item in self.p2_tree.get_children():
            self.p2_tree.delete(item)

        self.log_p2(f"Початок пакетного формування примірників № 2: {len(order_files)} наказ(ів)...")

        created_records = []
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0  # wdAlertsNone: не показувати блокуючі діалоги (напр. "Зберегти зміни?")

        try:
            # «Задня сторінка» завжди є окремим односторінковим шаблоном.
            test_tmpl = word.Documents.Open(back_page_abs, ReadOnly=True)
            test_tmpl.Repaginate()
            tmpl_pages = test_tmpl.ComputeStatistics(2)
            template_text = test_tmpl.Content.Text
            test_tmpl.Close(False)

            if tmpl_pages != 1:
                self.log_p2(f"❌ ПОМИЛКА: Файл «Задня сторінка» має {tmpl_pages} стор. (повинна бути рівно 1 сторінка)!")
                messagebox.showerror(
                    "Помилка шаблону",
                    f"Шаблон «Задня сторінка» містить {tmpl_pages} сторінок замість рівно 1 сторінки!\n\n"
                    f"Файл: {os.path.basename(back_page_abs)}\n\n"
                    "Для заміни останньої сторінки файл шаблону має займати рівно 1 сторінку."
                )
                return

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

                sub_folder_name = f"Наказ № {order_num}" if order_num else os.path.splitext(fname)[0]
                target_dir = os.path.join(out_root, sub_folder_name)
                os.makedirs(target_dir, exist_ok=True)

                out_copy_name = build_copy_two_filename(order_num, order_date, fname)
                target_file = os.path.join(target_dir, out_copy_name)

                self.log_p2("  Режим: повна копія наказу із заміною останньої сторінки.")
                shutil.copy2(order_path, target_file)
                doc = word.Documents.Open(os.path.abspath(target_file), ReadOnly=False)
                doc.Repaginate()
                total_pages = doc.ComputeStatistics(2)
                self.log_p2(f"  Кількість сторінок оригіналу: {total_pages}")
                last_start = doc.GoTo(1, 1, total_pages).Start
                content_end = max(last_start, doc.Content.End - 1)
                doc.Range(last_start, content_end).Delete()
                insert_pos = max(0, min(last_start, doc.Content.End - 1))
                doc.Range(insert_pos, insert_pos).InsertFile(back_page_abs)

                values = back_page_tag_values(order_num, order_date)
                # Зворотна сумісність для вже створених шаблонів.
                if order_num:
                    values["{{номер}}"] = order_num
                if order_date:
                    values["{{дата}}"] = order_date
                values["{{примірник_номер}}"] = copy_title
                values["{{засвідчення}}"] = "Згідно з оригіналом"
                for tag, value in values.items():
                    find_obj = doc.Content.Find
                    find_obj.Text = tag
                    while find_obj.Execute():
                        find_obj.Parent.Text = value
                        find_obj = doc.Content.Find
                        find_obj.Text = tag

                doc.Repaginate()
                final_pages = doc.ComputeStatistics(2)
                doc.Save()
                doc.Close(False)

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
            messagebox.showinfo(
                "Успіх",
                f"Успішно сформовано {total_orders} примірник(ів) № 2!\n\n"
                f"{p2_stats_msg}"
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
