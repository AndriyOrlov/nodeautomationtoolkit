from __future__ import annotations

import re

from nodeautomationtoolkit.core.definition import node
from nodeautomationtoolkit.core.table_types import DataTable

UKRAINIAN_MONTHS = (
    "січня|лютого|березня|квітня|травня|червня|липня|серпня|"
    "вересня|жовтня|листопада|грудня"
)

# База даних місяців із номерним зіставленням (1-12)
UKRAINIAN_MONTH_NUMBERS: dict[str, int] = {
    # 1 - Січень
    "січень": 1, "січня": 1, "січ": 1, "січ.": 1,
    # 2 - Лютий
    "лютий": 2, "лютого": 2, "лют": 2, "лют.": 2,
    # 3 - Березень
    "березень": 3, "березня": 3, "берез": 3, "берез.": 3, "бер": 3, "бер.": 3,
    # 4 - Квітень
    "квітень": 4, "квітня": 4, "квіт": 4, "квіт.": 4,
    # 5 - Травень
    "травень": 5, "травня": 5, "трав": 5, "трав.": 5,
    # 6 - Червень
    "червень": 6, "червня": 6, "черв": 6, "черв.": 6,
    # 7 - Липень
    "липень": 7, "липня": 7, "лип": 7, "лип.": 7,
    # 8 - Серпень
    "серпень": 8, "серпня": 8, "серп": 8, "серп.": 8,
    # 9 - Вересень
    "вересень": 9, "вересня": 9, "верес": 9, "верес.": 9, "вер": 9, "вер.": 9,
    # 10 - Жовтень
    "жовтень": 10, "жовтня": 10, "жовт": 10, "жовт.": 10, "жов": 10, "жов.": 10,
    # 11 - Листопад
    "листопад": 11, "листопада": 11, "листоп": 11, "листоп.": 11, "лист": 11, "лист.": 11,
    # 12 - Грудень
    "грудень": 12, "грудня": 12, "груд": 12, "груд.": 12,
}

MONTH_NAMES_BY_NUMBER: dict[int, dict[str, str]] = {
    1: {"name": "Січень", "genitive": "січня", "code": "01"},
    2: {"name": "Лютий", "genitive": "лютого", "code": "02"},
    3: {"name": "Березень", "genitive": "березня", "code": "03"},
    4: {"name": "Квітень", "genitive": "квітня", "code": "04"},
    5: {"name": "Травень", "genitive": "травня", "code": "05"},
    6: {"name": "Червень", "genitive": "червня", "code": "06"},
    7: {"name": "Липень", "genitive": "липня", "code": "07"},
    8: {"name": "Серпень", "genitive": "серпня", "code": "08"},
    9: {"name": "Вересень", "genitive": "вересня", "code": "09"},
    10: {"name": "Жовтень", "genitive": "жовтня", "code": "10"},
    11: {"name": "Листопад", "genitive": "листопада", "code": "11"},
    12: {"name": "Грудень", "genitive": "грудня", "code": "12"},
}


def parse_ukrainian_month(month_str: str) -> int | None:
    """Повертає порядковий номер місяця (1..12) за його назвою чи скороченням."""
    if not month_str:
        return None
    clean = month_str.strip().casefold()
    return UKRAINIAN_MONTH_NUMBERS.get(clean)


@node(
    name="Знайти у тексті",
    category="Текст",
    description=(
        "Знаходить текст або Regex і повертає збіги, позиції символів, рядок та контекст."
    ),
    type_id="builtin.text.find_detailed",
    outputs={
        "matches": "List",
        "count": "int",
        "first_position": "int",
        "preview": "str",
    },
)
def find_in_text(
    text: str = "",
    query: str = "",
    use_regex: bool = False,
    ignore_case: bool = True,
) -> dict:
    if not query:
        return {"matches": [], "count": 0, "first_position": -1, "preview": ""}
    pattern = query if use_regex else re.escape(query)
    flags = re.IGNORECASE if ignore_case else 0
    matches = []
    for match in re.finditer(pattern, text, flags):
        line = text.count("\n", 0, match.start()) + 1
        line_start = text.rfind("\n", 0, match.start()) + 1
        context_start = max(line_start, match.start() - 50)
        line_end = text.find("\n", match.end())
        context_end = len(text) if line_end < 0 else min(line_end, match.end() + 70)
        matches.append(
            {
                "text": match.group(0),
                "start": match.start(),
                "end": match.end(),
                "line": line,
                "column": match.start() - line_start + 1,
                "context": text[context_start:context_end].strip(),
            }
        )
    preview = "\n".join(
        f"Рядок {item['line']}, позиція {item['start']}: {item['context']}"
        for item in matches[:5]
    )
    return {
        "matches": matches,
        "count": len(matches),
        "first_position": matches[0]["start"] if matches else -1,
        "preview": preview,
    }


@node(
    name="Замінити у тексті",
    category="Текст",
    description="Замінює звичайний текст або Regex, один чи всі збіги.",
    type_id="builtin.text.replace_advanced",
    outputs={"text": "str", "replacements": "int"},
)
def replace_in_text(
    text: str = "",
    find: str = "",
    replacement_text: str = "",
    use_regex: bool = False,
    ignore_case: bool = True,
    replace_all: bool = True,
) -> dict:
    if not find:
        return {"text": text, "replacements": 0}
    pattern = find if use_regex else re.escape(find)
    flags = re.IGNORECASE if ignore_case else 0
    result, count = re.subn(
        pattern,
        replacement_text,
        text,
        count=0 if replace_all else 1,
        flags=flags,
    )
    return {"text": result, "replacements": count}


@node(
    name="Реквізити наказу",
    category="Наказ",
    description=(
        "Шукає дату, номер наказу, номер примірника та рядки зі службовими посадами."
    ),
    type_id="builtin.order.extract_fields",
    outputs={
        "order_date": "str",
        "order_number": "str",
        "copy_number": "str",
        "positions": "List",
        "summary": "str",
    },
)
def extract_order_fields(
    text: str = "",
    position_keywords: str = (
        "командир,начальник,заступник,керівник,голова,директор,"
        "т.в.о.,тимчасово виконуючий"
    ),
) -> dict:
    numeric_date = r"\b\d{1,2}[./]\d{1,2}[./]\d{2,4}\b"
    written_date = rf"\b\d{{1,2}}\s+(?:{UKRAINIAN_MONTHS})\s+\d{{4}}\s*(?:року|р\.)?"
    date_match = re.search(rf"(?:від\s+)?({written_date}|{numeric_date})", text, re.I)
    number_match = re.search(r"№\s*([A-Za-zА-Яа-яІіЇїЄєҐґ0-9./_-]+)", text)
    copy_match = re.search(
        r"(?:примірник|прим\.)\s*(?:№\s*)?([A-Za-zА-Яа-яІіЇїЄєҐґ0-9./_-]+)",
        text,
        re.I,
    )
    keywords = [item.strip().casefold() for item in position_keywords.split(",") if item.strip()]
    positions = []
    for line in text.splitlines():
        compact = " ".join(line.split())
        candidate = compact.casefold()
        if compact and any(keyword in candidate for keyword in keywords):
            positions.append(compact)
    positions = list(dict.fromkeys(positions))
    order_date = date_match.group(1).strip() if date_match else ""
    order_number = number_match.group(1).strip() if number_match else ""
    copy_number = copy_match.group(1).strip() if copy_match else ""
    summary = (
        f"Дата: {order_date or 'не знайдено'}\n"
        f"№ наказу: {order_number or 'не знайдено'}\n"
        f"Примірник: {copy_number or 'не знайдено'}\n"
        f"Посад: {len(positions)}"
    )
    return {
        "order_date": order_date,
        "order_number": order_number,
        "copy_number": copy_number,
        "positions": positions,
        "summary": summary,
    }


@node(
    name="Знайти місця в макеті наказу",
    category="Наказ",
    description=(
        "Знаходить у непідготовленому документі місце дати, позицію після № та "
        "початок причини. Показує всі кандидати, якщо місце неоднозначне."
    ),
    type_id="builtin.order.locate_layout_anchors",
    outputs={
        "anchors": "Dictionary",
        "candidates": "DataTable",
        "date_position": "int",
        "number_position": "int",
        "reason_position": "int",
        "summary": "str",
    },
)
def locate_order_layout_anchors(
    text: str = "",
    date_pattern: str = (
        rf"(?:[«“\"]\s*[»”\"]\s*)?(?:\d{{1,2}}\s+)?"
        rf"(?:{UKRAINIAN_MONTHS})\s+\d{{4}}\s*(?:року|р\.)?"
    ),
    number_pattern: str = r"№\s*",
    reason_pattern: str = (
        r"(?im)^\s*(?:відповідно|підстава|причина|у\s+зв'язку|"
        r"на\s+підставі|на\s+виконання)\b"
    ),
) -> dict:
    rules = {
        "Дата": _compile_line_pattern(date_pattern),
        "Номер": _compile_line_pattern(number_pattern),
        "Початок причини": _compile_line_pattern(reason_pattern),
    }
    anchors: dict[str, list[dict]] = {}
    rows = []
    for label, pattern in rules.items():
        found = []
        for match in pattern.finditer(text):
            line = text.count("\n", 0, match.start()) + 1
            line_start = text.rfind("\n", 0, match.start()) + 1
            line_end = text.find("\n", match.end())
            if line_end < 0:
                line_end = len(text)
            excerpt = " ".join(text[line_start:line_end].split())
            item = {
                "start": match.start(),
                "end": match.end(),
                "line": line,
                "text": match.group(0),
                "excerpt": excerpt,
            }
            found.append(item)
            rows.append((label, len(found), line, match.start(), excerpt))
        anchors[label] = found

    def unique_position(label: str, *, at_end: bool = False) -> int:
        found = anchors[label]
        if len(found) != 1:
            return -1
        return int(found[0]["end"] if at_end else found[0]["start"])

    table = DataTable(
        ("Тип", "Кандидат", "Рядок", "Позиція", "Контекст"),
        tuple(rows),
        "Місця в макеті наказу",
    )
    summary = " · ".join(f"{label}: {len(found)}" for label, found in anchors.items())
    return {
        "anchors": anchors,
        "candidates": table,
        "date_position": unique_position("Дата"),
        "number_position": unique_position("Номер", at_end=True),
        "reason_position": unique_position("Початок причини"),
        "summary": summary,
    }


def _compile_line_pattern(pattern: str) -> re.Pattern:
    try:
        return re.compile(pattern, re.I | re.M)
    except re.error as error:
        raise ValueError(f"Некоректний Regex розбивки: {error}") from error


@node(
    name="Розбити наказ на блоки",
    category="Наказ",
    description=(
        "Послідовно розбиває текст на параграфи, причини, пункти та звичайний текст. "
        "Regex-правила можна змінити під інший тип наказу."
    ),
    type_id="builtin.order.split_blocks",
    outputs={
        "blocks": "List",
        "sections": "List",
        "reasons": "List",
        "action_headers": "List",
        "items": "List",
        "header": "str",
        "body": "str",
        "signature": "str",
        "preview": "str",
    },
)
def split_order_blocks(
    text: str = "",
    section_pattern: str = r"^\s*(?:§\s*\d+|параграф\s+\d+|розділ\s+[IVXLC\d]+)\b",
    reason_pattern: str = (
        r"^\s*(?:відповідно|підстава|причина|у\s+зв'язку|на\s+підставі|"
        r"на\s+виконання)\b"
    ),
    action_pattern: str = (
        r"^\s*(?:звільнити(?:\s+(?:і|та)\s+призначити)?|"
        r"призначити(?:\s+(?:і|та)\s+звільнити)?|перемістити|"
        r"зарахувати|виключити)\b.*(?::|\bдо\b)"
    ),
    item_pattern: str = r"^\s*\d+(?:\.\d+)*[.)]\s*",
    signature_pattern: str = (
        r"^\s*(?:командир|командувач|начальник|заступник|керівник|голова|директор|"
        r"т\.в\.о\.|тимчасово\s+виконуючий)\b"
    ),
) -> dict:
    patterns = {
        "section": _compile_line_pattern(section_pattern),
        "reason": _compile_line_pattern(reason_pattern),
        "action": _compile_line_pattern(action_pattern),
        "item": _compile_line_pattern(item_pattern),
    }
    signature_regex = _compile_line_pattern(signature_pattern)
    starts = [
        match.start()
        for pattern in patterns.values()
        if (match := pattern.search(text)) is not None
    ]
    body_start = min(starts) if starts else 0
    signature_match = signature_regex.search(text, body_start)
    signature_start = signature_match.start() if signature_match else len(text)
    header = text[:body_start].strip()
    body = text[body_start:signature_start].strip()
    signature = text[signature_start:].strip()
    blocks: list[dict] = []
    current_type = "text"
    current_lines: list[str] = []

    def flush() -> None:
        nonlocal current_lines
        value = "\n".join(current_lines).strip()
        if value:
            blocks.append(
                {"index": len(blocks), "type": current_type, "text": value}
            )
        current_lines = []

    for line in body.splitlines():
        matched_type = next(
            (kind for kind, pattern in patterns.items() if pattern.search(line)),
            None,
        )
        if matched_type is not None:
            flush()
            current_type = matched_type
            current_lines = [line]
        elif current_lines:
            current_lines.append(line)
        elif line.strip():
            current_type = "text"
            current_lines = [line]
    flush()

    sections = [item["text"] for item in blocks if item["type"] == "section"]
    reasons = [item["text"] for item in blocks if item["type"] == "reason"]
    action_headers = [item["text"] for item in blocks if item["type"] == "action"]
    items = [item["text"] for item in blocks if item["type"] == "item"]
    preview = "\n".join(
        f"{item['type']}: {' '.join(item['text'].split())[:120]}" for item in blocks[:8]
    )
    return {
        "blocks": blocks,
        "sections": sections,
        "reasons": reasons,
        "action_headers": action_headers,
        "items": items,
        "header": header,
        "body": body,
        "signature": signature,
        "preview": preview,
    }


@node(
    name="Згрупувати пункти за мітками",
    category="Наказ",
    description=(
        "Формує окремий текст для кожної мітки-відправника. Зберігає найближчі "
        "параграф і причину; пункт із двома мітками потрапляє у дві групи."
    ),
    type_id="builtin.order.group_by_markers",
    outputs={
        "groups": "Dictionary",
        "names": "List",
        "counts": "Dictionary",
        "summary": "str",
    },
)
def group_items_by_markers(
    blocks: list | None = None,
    markers_text: str = "",
    markers: list | None = None,
    ignore_case: bool = True,
) -> dict:
    resolved_markers = [str(item).strip() for item in (markers or []) if str(item).strip()]
    if not resolved_markers:
        resolved_markers = [line.strip() for line in markers_text.splitlines() if line.strip()]
    resolved_markers = list(dict.fromkeys(resolved_markers))
    if not resolved_markers:
        return {"groups": {}, "names": [], "counts": {}, "summary": "Мітки не задані"}
    groups: dict[str, list[str]] = {marker: [] for marker in resolved_markers}
    group_counts: dict[str, int] = dict.fromkeys(resolved_markers, 0)
    context_section = ""
    context_reason = ""
    context_action = ""
    context_section_markers: set[str] = set()
    context_reason_markers: set[str] = set()
    context_action_markers: set[str] = set()
    emitted_context: dict[str, tuple[str, str, str]] = {
        marker: ("", "", "") for marker in resolved_markers
    }
    for block in blocks or []:
        if not isinstance(block, dict):
            continue
        block_type = str(block.get("type", "text"))
        block_text = str(block.get("text", ""))
        if block_type == "section":
            context_section = block_text
            context_reason = ""
            context_action = ""
            candidate = block_text.casefold() if ignore_case else block_text
            context_section_markers = {
                marker
                for marker in resolved_markers
                if (marker.casefold() if ignore_case else marker) in candidate
            }
            context_reason_markers = set()
            context_action_markers = set()
            continue
        if block_type == "reason":
            context_reason = block_text
            context_action = ""
            candidate = block_text.casefold() if ignore_case else block_text
            context_reason_markers = {
                marker
                for marker in resolved_markers
                if (marker.casefold() if ignore_case else marker) in candidate
            }
            context_action_markers = set()
            continue
        if block_type == "action":
            context_action = block_text
            candidate = block_text.casefold() if ignore_case else block_text
            context_action_markers = {
                marker
                for marker in resolved_markers
                if (marker.casefold() if ignore_case else marker) in candidate
            }
            continue
        if block_type != "item":
            continue
        candidate = block_text.casefold() if ignore_case else block_text
        direct_markers = {
            marker
            for marker in resolved_markers
            if (marker.casefold() if ignore_case else marker) in candidate
        }
        matched_markers = (
            context_section_markers
            | context_reason_markers
            | context_action_markers
            | direct_markers
        )
        for marker in resolved_markers:
            if marker not in matched_markers:
                continue
            previous_section, previous_reason, previous_action = emitted_context[marker]
            if context_section and context_section != previous_section:
                groups[marker].append(context_section)
                previous_section = context_section
                previous_reason = ""
                previous_action = ""
            if context_reason and context_reason != previous_reason:
                groups[marker].append(context_reason)
                previous_reason = context_reason
                previous_action = ""
            if context_action and context_action != previous_action:
                groups[marker].append(context_action)
                previous_action = context_action
            groups[marker].append(block_text)
            group_counts[marker] += 1
            emitted_context[marker] = (
                previous_section,
                previous_reason,
                previous_action,
            )
    compact_groups = {
        marker: "\n".join(parts) for marker, parts in groups.items() if parts
    }
    summary = "\n".join(
        f"{name}: {group_counts[name]} пунктів"
        for name in compact_groups
    )
    return {
        "groups": compact_groups,
        "names": list(compact_groups),
        "counts": {name: group_counts[name] for name in compact_groups},
        "summary": summary,
    }


@node(
    name="Візуалізувати макет наказу",
    category="Наказ",
    description=(
        "Генерує кольорове схематичне зображення (PNG) та розмітку розпізнаних "
        "блоків наказу (параграфи, причини, шапки дій, пункти, підпис)."
    ),
    type_id="builtin.order.visualize_layout",
    outputs={
        "image_path": "str",
        "blocks_summary": "str",
        "table": "DataTable",
        "preview": "str",
    },
)
def visualize_order_layout(
    text: str = "",
    blocks: list | None = None,
    output_image_path: str = "",
) -> dict:
    import tempfile
    from pathlib import Path
    from PIL import Image, ImageDraw
    from nodeautomationtoolkit.core.table_types import DataTable

    if not blocks and text:
        blocks = split_order_blocks(text)["blocks"]
    blocks = blocks or []

    type_styles = {
        "header": {"bg": (30, 58, 138), "border": (59, 130, 246), "label": "ШАПКА НАКАЗУ", "badge": "🔵"},
        "section": {"bg": (30, 58, 138), "border": (59, 130, 246), "label": "ПАРАГРАФ", "badge": "🔵"},
        "reason": {"bg": (88, 28, 135), "border": (168, 85, 247), "label": "ПРИЧИНА / ПІДСТАВА", "badge": "🟣"},
        "action": {"bg": (6, 95, 70), "border": (52, 211, 153), "label": "ШАПКА ДІЇ", "badge": "🟢"},
        "item": {"bg": (120, 53, 15), "border": (251, 191, 36), "label": "НУМЕРОВАНИЙ ПУНКТ", "badge": "🟡"},
        "signature": {"bg": (136, 19, 55), "border": (244, 63, 94), "label": "ПІДПИС", "badge": "🔴"},
        "text": {"bg": (49, 46, 129), "border": (129, 140, 248), "label": "ЗВИЧАЙНИЙ ТЕКСТ / ЗАСВІДЧЕННЯ", "badge": "🟣"},
    }

    table_rows = []
    summary_lines = []
    for idx, b in enumerate(blocks, start=1):
        b_type = str(b.get("type", "text"))
        b_text = str(b.get("text", ""))
        style = type_styles.get(b_type, type_styles["text"])
        excerpt = " ".join(b_text.split())[:80]
        table_rows.append((idx, b_type, style["label"], style["badge"], excerpt))
        summary_lines.append(f"{style['badge']} Блок {idx} [{style['label']}]: {excerpt}")

    width = 840
    card_height = 80
    gap = 14
    margin = 30
    header_h = 100
    img_height = max(500, header_h + margin + len(blocks) * (card_height + gap) + margin)

    img = Image.new("RGB", (width, img_height), color=(15, 23, 42))
    draw = ImageDraw.Draw(img)

    draw.rectangle([0, 0, width, header_h], fill=(30, 41, 59))
    draw.text((margin, 25), "СТРУКТУРНИЙ МАКЕТ НАКАЗУ", fill=(248, 250, 252))
    draw.text((margin, 60), f"Розпізнано блоків: {len(blocks)}", fill=(148, 163, 184))

    y = header_h + margin
    for idx, b in enumerate(blocks, start=1):
        b_type = str(b.get("type", "text"))
        b_text = str(b.get("text", ""))
        style = type_styles.get(b_type, type_styles["text"])

        draw.rounded_rectangle(
            [margin, y, width - margin, y + card_height],
            radius=8,
            fill=style["bg"],
            outline=style["border"],
            width=2,
        )
        draw.rounded_rectangle(
            [margin, y, margin + 12, y + card_height],
            radius=4,
            fill=style["border"],
        )
        header_str = f"БЛОК {idx}  |  {style['label']}"
        draw.text((margin + 24, y + 12), header_str, fill=(255, 255, 255))
        snippet = " ".join(b_text.split())
        if len(snippet) > 85:
            snippet = snippet[:84] + "..."
        draw.text((margin + 24, y + 42), snippet, fill=(226, 232, 240))
        y += card_height + gap

    if output_image_path.strip():
        out_path = Path(output_image_path).expanduser().resolve()
    else:
        temp_dir = Path(tempfile.gettempdir()) / "nat_layout_previews"
        temp_dir.mkdir(parents=True, exist_ok=True)
        out_path = temp_dir / f"order_layout_{len(blocks)}_blocks.png"

    img.save(out_path)

    table = DataTable(
        ("№", "Тип", "Назва блока", "Мітка", "Контекст"),
        tuple(table_rows),
        "Розпізнані блоки наказу",
    )
    blocks_summary = "\n".join(summary_lines)
    preview = f"Зображення створено: {out_path.name}\nБлоків: {len(blocks)}\n\n" + blocks_summary[:300]

    return {
        "image_path": str(out_path),
        "blocks_summary": blocks_summary,
        "table": table,
        "preview": preview,
    }

