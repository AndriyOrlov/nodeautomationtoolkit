"""Адресати повідомлень і пункти, які в них не переносяться.

Винесено з `generate_extracts.py` без змін логіки (26.10.2026).
"""
from __future__ import annotations

import re
from nodeautomationtoolkit.builtin_nodes.recipient_mapping import (
    _table_cipher,
    _format_full_closed_unit_text,
)



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
        # Корпус — лише коли «корпус» стоїть у НАЗИВНОМУ («11 армійський
        # корпус») або скорочення «N АК». Частина, у стовпці A якої записано
        # підпорядкування («… бригада 11 армійського корпусу»), корпусом не є:
        # раніше вона йшла в групу корпусів, і група частин лишалась без неї.
        open_name = str(entry.get("open_name", ""))
        abbreviation = str(entry.get("abbreviation", ""))
        return bool(
            re.search(r"\bкорпус\b", open_name, re.IGNORECASE)
            or re.fullmatch(r"\s*\d{1,3}\s*АК\s*", abbreviation, re.IGNORECASE)
        )

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


#: Текст у {{кому_список}} супроводу, коли адресати — в окремому «Розрахунку розсилки».
MESSAGE_BY_CALCULATION_TEXT = "Згідно з розрахунком розсилки"


def message_skipped_item_lines(routes: dict) -> tuple[set[int], list[str]]:
    """Рядки наказу, які НЕ переносяться у зміст повідомлення.

    Це пункти внутрішнього переміщення в управлінні (з посади в управлінні на
    посаду в управлінні): повідомляти про них нікого. Разом із пунктом
    прибирається й шапка (§, «Відповідно до …:»), під якою не лишилось жодного
    іншого пункту, — інакше вона висіла б у повідомленні сама. Номери решти
    пунктів не змінюються: це номери пунктів наказу.
    """
    items = routes.get("item_spans") or []
    skipped = [item for item in items if item.get("internal_management_move")]
    if not skipped:
        return set(), []

    lines: set[int] = set()
    for item in skipped:
        lines.update(range(int(item["start_line"]), int(item["end_line"]) + 1))

    kept_heading_starts = {
        int(heading_range[0])
        for item in items
        if not item.get("internal_management_move")
        for heading_range in item.get("heading_ranges") or []
    }
    for item in skipped:
        for start, end in item.get("heading_ranges") or []:
            if int(start) not in kept_heading_starts:
                lines.update(range(int(start), int(end) + 1))
    return lines, [str(item.get("label", "")).strip() for item in skipped]


_MESSAGE_ITEM_BREAK_RE = re.compile(r"^\s*(?:\d{1,3}(?:\.\d{1,3})*[\.\)]\s|§|Відповідно\b)")


def message_item_contexts(paragraphs: list[str]) -> list[str]:
    """Для кожного абзацу — текст попередніх абзаців ТОГО САМОГО пункту.

    Зміст повідомлення шифрується поабзацно, а ТЦК може бути названо в одному
    абзаці пункту, а «цього самого центру» — в наступному. Пункт закінчується
    на порожньому абзаці, а новий починається з номера, «§» чи «Відповідно до»
    (номер автонумерації Word у тексті абзацу відсутній — тоді межею є
    порожній абзац між пунктами).
    """
    contexts = []
    current: list[str] = []
    for text in paragraphs:
        if not text.strip() or _MESSAGE_ITEM_BREAK_RE.match(text):
            current = []
        contexts.append("\n".join(current))
        if text.strip():
            current.append(text)
    return contexts


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
