"""Дати, номер і дата наказу з назви файлу, назви вихідних файлів.

Винесено з `generate_extracts.py` без змін логіки (26.10.2026).
"""
from __future__ import annotations

import os
import re
from datetime import datetime



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
