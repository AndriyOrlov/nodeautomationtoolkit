from __future__ import annotations

import csv
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from nodeautomationtoolkit.core.definition import node
from nodeautomationtoolkit.core.table_types import DataTable

_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"


def _cell_column(reference: str) -> int:
    letters = re.match(r"[A-Z]+", reference.upper())
    result = 0
    for char in letters.group(0) if letters else "A":
        result = result * 26 + ord(char) - 64
    return result - 1


def _read_xlsx(path: Path, sheet_name: str = "") -> list[list[str]]:
    with zipfile.ZipFile(path) as package:
        shared: list[str] = []
        if "xl/sharedStrings.xml" in package.namelist():
            root = ET.fromstring(package.read("xl/sharedStrings.xml"))
            for item in root.findall(f"{{{_MAIN}}}si"):
                shared.append("".join(node.text or "" for node in item.iter(f"{{{_MAIN}}}t")))

        workbook = ET.fromstring(package.read("xl/workbook.xml"))
        relationships = ET.fromstring(package.read("xl/_rels/workbook.xml.rels"))
        targets = {
            item.attrib["Id"]: item.attrib["Target"]
            for item in relationships.findall(f"{{{_PKG_REL}}}Relationship")
        }
        sheets = workbook.find(f"{{{_MAIN}}}sheets")
        if sheets is None or not list(sheets):
            return []
        selected = next(
            (item for item in sheets if item.attrib.get("name") == sheet_name),
            list(sheets)[0],
        )
        rel_id = selected.attrib.get(f"{{{_REL}}}id", "")
        target = targets.get(rel_id, "worksheets/sheet1.xml").lstrip("/")
        target = target if target.startswith("xl/") else "xl/" + target
        sheet = ET.fromstring(package.read(target))
        rows: list[list[str]] = []
        for row in sheet.findall(f".//{{{_MAIN}}}row"):
            values: list[str] = []
            for cell in row.findall(f"{{{_MAIN}}}c"):
                column = _cell_column(cell.attrib.get("r", "A1"))
                while len(values) <= column:
                    values.append("")
                kind = cell.attrib.get("t", "")
                if kind == "inlineStr":
                    value = "".join(
                        item.text or "" for item in cell.iter(f"{{{_MAIN}}}t")
                    )
                else:
                    element = cell.find(f"{{{_MAIN}}}v")
                    value = element.text if element is not None and element.text else ""
                    if kind == "s" and value:
                        value = shared[int(value)]
                values[column] = value.strip()
            rows.append(values)
        return rows


def _read_rows(path: Path, sheet_name: str = "") -> list[list[str]]:
    if path.suffix.casefold() == ".xlsx":
        return _read_xlsx(path, sheet_name)
    if path.suffix.casefold() != ".csv":
        raise ValueError("Таблиця має бути CSV або XLSX")
    raw = path.read_text(encoding="utf-8-sig", errors="replace")
    delimiter = ";" if raw.count(";") >= raw.count(",") else ","
    reader = csv.reader(raw.splitlines(), delimiter=delimiter)
    return [[cell.strip() for cell in row] for row in reader]


@node(
    name="Прочитати таблицю відповідностей",
    category="Наказ",
    description=(
        "Читає CSV/XLSX: відкрите найменування, шифр та куди направляється. "
        "Вміст залишається локально й далі передається по дротах графа."
    ),
    type_id="builtin.order.read_recipient_mapping",
    outputs={"mapping": "Dictionary", "markers": "List", "table": "DataTable", "count": "int"},
)
def read_recipient_mapping(
    path: str = "",
    open_name_column: str = "Відкрите найменування",
    cipher_column: str = "Шифр",
    destination_column: str = "Куди направляється",
    sheet_name: str = "",
) -> dict:
    source = Path(path).expanduser()
    if not source.is_file():
        raise FileNotFoundError(f"Таблицю не знайдено: {path or '(шлях порожній)'}")
    rows = _read_rows(source, sheet_name)
    if not rows:
        raise ValueError("Таблиця порожня")
    requested = [open_name_column, cipher_column, destination_column]
    header_index = -1
    normalized: dict[str, int] = {}
    for index, row in enumerate(rows[:20]):
        candidate = {
            name.strip().casefold(): column
            for column, name in enumerate(row)
            if name.strip()
        }
        if all(name.casefold() in candidate for name in requested):
            header_index = index
            normalized = candidate
            break
    if header_index < 0:
        headers = [cell.strip() for cell in rows[0]]
        normalized = {name.casefold(): index for index, name in enumerate(headers)}
    missing = [name for name in requested if name.casefold() not in normalized]
    if missing:
        raise ValueError("У таблиці немає колонок: " + ", ".join(missing))
    indexes = [normalized[name.casefold()] for name in requested]
    mapping: dict[str, dict[str, str]] = {}
    table_rows = []
    for row in rows[header_index + 1 :]:
        values = [row[index].strip() if index < len(row) else "" for index in indexes]
        open_name, cipher, destination = values
        if not open_name:
            continue
        mapping[open_name] = {
            "open_name": open_name,
            "cipher": cipher,
            "destination": destination,
        }
        table_rows.append(tuple(values))
    table = DataTable(tuple(requested), tuple(table_rows), "Таблиця відповідностей")
    return {
        "mapping": mapping,
        "markers": list(mapping),
        "table": table,
        "count": len(mapping),
    }


@node(
    name="Перетворити групи на шифри",
    category="Наказ",
    description=(
        "Зіставляє знайдені відкриті найменування з таблицею, об'єднує однакові "
        "шифри та готує дані для пакета DOCX і звіту."
    ),
    type_id="builtin.order.groups_to_ciphers",
    outputs={
        "documents": "Dictionary",
        "report": "DataTable",
        "missing": "List",
        "summary": "str",
    },
)
def groups_to_ciphers(
    groups: dict | None = None,
    counts: dict | None = None,
    mapping: dict | None = None,
) -> dict:
    documents: dict[str, dict] = {}
    report_rows = []
    missing = []
    for open_name, content in (groups or {}).items():
        entry = (mapping or {}).get(open_name)
        count = int((counts or {}).get(open_name, 0))
        if not isinstance(entry, dict) or not str(entry.get("cipher", "")).strip():
            missing.append(str(open_name))
            report_rows.append((open_name, "", "", count, "Немає відповідності"))
            continue
        cipher = str(entry["cipher"]).strip()
        destination = str(entry.get("destination", "")).strip()
        current = documents.setdefault(
            cipher,
            {
                "name": cipher,
                "cipher": cipher,
                "sender": cipher,
                "destination": destination,
                "content": "",
                "count": 0,
                "open_names": [],
            },
        )
        if current["content"] and str(content).strip():
            current["content"] += "\n"
        current["content"] += str(content).strip()
        current["count"] += count
        current["open_names"].append(str(open_name))
        if not current["destination"]:
            current["destination"] = destination
        report_rows.append((open_name, cipher, destination, count, "Готово"))
    report = DataTable(
        ("Відкрите найменування", "Шифр", "Куди направляється", "Знайдено пунктів", "Стан"),
        tuple(report_rows),
        "Результат групування",
    )
    summary = f"Документів: {len(documents)} · без відповідності: {len(missing)}"
    return {"documents": documents, "report": report, "missing": missing, "summary": summary}


def _normalize_unit_name(name: str) -> str:
    """Нормалізує назву ВЧ для порівняння: видаляє зайві пробіли, переводить у нижній регістр."""
    return re.sub(r"\s+", " ", name.strip().casefold())


def _build_unit_fuzzy_pattern(open_name: str) -> re.Pattern:
    """
    Будує нечіткий Regex-патерн для пошуку ВЧ у тексті наказу.

    Принцип — мінімальна сигнатура з числа + перших 3 літер кожного значущого слова:

        "160 окрема механізована бригада"
        → [160] [окр...] [мех...] [бриг...]
        → патерн:  160.{0,80}окр\\w*.{0,80}мех\\w*.{0,80}бриг\\w*

    Це знаходить ВЧ у БУДЬ-ЯКОМУ відмінку та з будь-якими вставками між словами
    (артиклі, дефіси, порядкові числівники тощо).

    Прийменники (та, і, з, на, в, у, до, від) — пропускаються, не входять до сигнатури.
    Абревіатури (ОМБр, ОТЦК) — беруть перші 3 літери так само.
    Числа — обов'язковий якір, точний збіг.

    Ручний короткий запис теж підтримується:
        "160 окр"   → знайде "160 окремої механізованої бригади"
        "167 мех"   → знайде "167 окрема механізована бригада"
    """
    STOP_WORDS = {
        "та", "і", "й", "з", "зі", "із", "на", "в", "у", "до", "від", "по", "при", "за",
        "the", "of", "and", "а",
    }

    tokens = re.split(r"[\s\-–—,]+", open_name.strip())

    # Виокремлюємо значущі частини: числа + перші 3 букви слів
    anchors: list[str] = []
    for token in tokens:
        if not token:
            continue
        # Числа → точний збіг (якір)
        if re.match(r"^\d+$", token):
            anchors.append(re.escape(token))
        # Прийменники / сполучники → пропускаємо
        elif token.casefold() in STOP_WORDS:
            continue
        # Будь-яке слово (в т.ч. абревіатуру) → беремо перші 3 символи + \w*
        elif len(token) >= 2:
            prefix = re.escape(token[:3])
            anchors.append(prefix + r"\w*")

    if not anchors:
        return re.compile(re.escape(open_name), re.IGNORECASE)

    # З'єднуємо якорі: між ними допускаємо до 80 довільних символів
    # (прийменники, відмінкові закінчення, дефіси, порядкові числівники тощо)
    pattern_str = r".{0,80}?".join(anchors)
    return re.compile(pattern_str, re.IGNORECASE | re.UNICODE | re.DOTALL)



@node(
    name="Картування та пошук військових частин",
    category="Наказ",
    description=(
        "Шукає відкриті найменування військових частин у тексті наказу, замінює їх на закриті "
        "найменування (шифри / в/ч), зберігає шапки, преамбули (§) та повні блоки пунктів з Підставами. "
        "Підтримує нечіткий відмінковий пошук: знаходить 'бригади' замість 'бригада', "
        "'механізованої' замість 'механізована' тощо."
    ),
    type_id="builtin.order.map_military_units",
    outputs={
        "processed_text": "str",
        "units_table": "DataTable",
        "units_list": "List",
        "unit_paragraphs": "Dictionary",
        "match_report": "DataTable",
        "summary": "str",
    },
)
def map_military_units(
    text: str = "",
    mapping: dict | None = None,
    default_prefix: str = "військова частина ",
    fuzzy_match: bool = True,
) -> dict:
    if not text.strip():
        return {
            "processed_text": "",
            "units_table": DataTable(("Закрита ВЧ", "Відкрита назва", "Пункти наказу", "Згадок"), ()),
            "units_list": [],
            "unit_paragraphs": {},
            "match_report": DataTable(("Шаблон пошуку", "Знайдений збіг", "Шифр", "Кількість"), ()),
            "summary": "Порожній текст наказу",
        }

    mapping_dict = mapping or {}
    lines = [line.rstrip() for line in text.splitlines()]

    # ── Визначаємо шапку наказу ────────────────────────────────────────────────
    header_lines = []
    content_start_idx = len(lines)
    for idx, line in enumerate(lines):
        clean = line.strip()
        if (
            clean.startswith("§")
            or re.match(r"^\d+[\.\)]", clean)
            or "НАКАЗУЮ" in clean.upper()
            or "ПРИЗНАЧИТИ" in clean.upper()
        ):
            content_start_idx = idx
            break
        if clean:
            header_lines.append(line)

    if not header_lines and lines:
        header_lines = lines[:3]

    # ── Компілюємо патерни для кожної ВЧ ──────────────────────────────────────
    unit_patterns: list[tuple[str, str, re.Pattern]] = []
    for open_name, mapped_val in mapping_dict.items():
        if isinstance(mapped_val, dict):
            closed_code = str(
                mapped_val.get("cipher")
                or mapped_val.get("closed_name")
                or open_name
            )
        else:
            closed_code = str(mapped_val)

        if fuzzy_match:
            pattern = _build_unit_fuzzy_pattern(open_name)
        else:
            pattern = re.compile(re.escape(open_name), re.IGNORECASE)

        unit_patterns.append((open_name, closed_code, pattern))

    # ── Парсимо тіло наказу у блоки (§-параграфи та пронумеровані пункти) ─────
    blocks = []
    current_parent_heading = ""
    current_block = None

    for line in lines[content_start_idx:]:
        clean = line.strip()
        if not clean:
            if current_block:
                current_block["lines"].append("")
            continue

        is_section_marker = clean.startswith("§") or (
            ("Відповідно до" in clean or "Згідно з" in clean or clean.endswith(":"))
            and not re.match(r"^\d+[\.\d]*", clean)
            and not clean.startswith("Підстава")
        )

        is_new_item = (
            not is_section_marker
            and not ("р.н." in clean or "р. н." in clean or clean.startswith("Підстава"))
            and bool(re.match(r"^\d{1,3}(?:\.\d{1,3})*[\.\)]\s+", clean))
        )

        if is_section_marker:
            if clean.startswith("§") or not current_parent_heading:
                current_parent_heading = line
            else:
                current_parent_heading += "\n\n" + line
            current_block = {
                "type": "section",
                "heading": current_parent_heading,
                "label": clean.split()[0] if clean.split() else "Розділ",
                "lines": [line],
            }
            blocks.append(current_block)
        elif is_new_item:
            match = re.match(r"^(\d+[\.\d]*)", clean)
            label = f"Пункт {match.group(1)}" if match else "Пункт"
            current_block = {
                "type": "item",
                "heading": current_parent_heading,
                "label": label,
                "lines": [line],
            }
            blocks.append(current_block)
        else:
            if current_block:
                current_block["lines"].append(line)
            else:
                current_block = {
                    "type": "item",
                    "heading": current_parent_heading,
                    "label": "Основний текст",
                    "lines": [line],
                }
                blocks.append(current_block)

    # ── Шукаємо ВЧ у кожному пункті ───────────────────────────────────────────
    unit_data_map: dict[str, dict] = {}
    unit_counts: dict[str, int] = {}
    unit_open_names: dict[str, set[str]] = {}
    match_report_rows = []

    processed_lines = list(lines)

    for block in blocks:
        if block["type"] != "item":
            continue

        block_raw_text = "\n".join(block["lines"])
        block_replaced_lines = list(block["lines"])
        matched_units_in_block: set[tuple[str, str]] = set()

        for open_name, closed_code, pattern in unit_patterns:
            # Знаходимо всі збіги (для звіту)
            all_matches = pattern.findall(block_raw_text)
            if not all_matches:
                continue

            # Підраховуємо унікальні знайдені форми
            for found_form in set(str(m) if isinstance(m, str) else str(m[0]) for m in all_matches):
                existing = next(
                    (r for r in match_report_rows if r[0] == open_name and r[1] == found_form),
                    None,
                )
                if existing:
                    idx = match_report_rows.index(existing)
                    match_report_rows[idx] = (
                        existing[0], existing[1], existing[2], existing[3] + 1
                    )
                else:
                    match_report_rows.append((open_name, found_form, closed_code, 1))

            # Замінюємо знайдені форми на закритий шифр
            block_replaced_lines = [pattern.sub(closed_code, ln) for ln in block_replaced_lines]
            matched_units_in_block.add((closed_code, open_name))

        if matched_units_in_block:
            full_item_text = "\n".join(block_replaced_lines).strip()
            for closed_code, open_name in matched_units_in_block:
                unit_entry = unit_data_map.setdefault(
                    closed_code,
                    {
                        "unit_code": closed_code,
                        "header_lines": header_lines,
                        "items": [],
                    },
                )
                unit_entry["items"].append(
                    {
                        "parent_heading": block["heading"],
                        "label": block["label"],
                        "text": full_item_text,
                    }
                )
                unit_counts[closed_code] = unit_counts.get(closed_code, 0) + 1
                unit_open_names.setdefault(closed_code, set()).add(open_name)

    # ── Будуємо вихідні таблиці ────────────────────────────────────────────────
    table_rows = []
    for closed_code, count in unit_counts.items():
        open_str = ", ".join(sorted(unit_open_names.get(closed_code, [])))
        entry = unit_data_map[closed_code]
        p_labels = ", ".join(dict.fromkeys(item["label"] for item in entry["items"]))
        table_rows.append((closed_code, open_str, p_labels, count))

    table = DataTable(
        ("Закрита назва (ВЧ)", "Відкрита назва", "Пункти наказу", "Згадок"),
        tuple(table_rows),
        "Розпізнані військові частини",
    )
    match_report = DataTable(
        ("Шаблон пошуку", "Знайдена форма у тексті", "Замінено на (шифр)", "Кількість"),
        tuple(match_report_rows),
        "Звіт відмінкових збігів",
    )

    final_text = "\n".join(processed_lines)
    summary = f"Знайдено частин: {len(unit_counts)} · Загалом згадок: {sum(unit_counts.values())} · Відмінкових форм розпізнано: {len(match_report_rows)}"

    return {
        "processed_text": final_text,
        "units_table": table,
        "units_list": list(unit_counts.keys()),
        "unit_paragraphs": unit_data_map,
        "match_report": match_report,
        "summary": summary,
    }

def map_military_units(
    text: str = "",
    mapping: dict | None = None,
    default_prefix: str = "військова частина ",
) -> dict:
    if not text.strip():
        return {
            "processed_text": "",
            "units_table": DataTable(("Закрита ВЧ", "Відкрита назва", "Пункти наказу", "Згадок"), ()),
            "units_list": [],
            "unit_paragraphs": {},
            "summary": "Порожній текст наказу",
        }

    mapping_dict = mapping or {}
    lines = [line.rstrip() for line in text.splitlines()]

    header_lines = []
    content_start_idx = len(lines)
    for idx, line in enumerate(lines):
        clean = line.strip()
        if clean.startswith("§") or re.match(r"^\d+[\.\)]", clean) or "НАКАЗУЮ" in clean.upper() or "ПРИЗНАЧИТИ" in clean.upper():
            content_start_idx = idx
            break
        if clean:
            header_lines.append(line)

    if not header_lines and lines:
        header_lines = lines[:3]

    blocks = []
    current_parent_heading = ""
    current_block = None

    for line in lines[content_start_idx:]:
        clean = line.strip()
        if not clean:
            if current_block:
                current_block["lines"].append("")
            continue

        is_section_marker = clean.startswith("§") or (("Відповідно до" in clean or "Згідно з" in clean or clean.endswith(":")) and not re.match(r"^\d+[\.\d]*", clean) and not clean.startswith("Підстава"))
        
        is_new_item = (
            not is_section_marker
            and not ("р.н." in clean or "р. н." in clean or clean.startswith("Підстава"))
            and bool(re.match(r"^\d{1,3}(?:\.\d{1,3})*[\.\)]\s+", clean))
        )

        if is_section_marker:
            if clean.startswith("§") or not current_parent_heading:
                current_parent_heading = line
            else:
                current_parent_heading += "\n\n" + line
            current_block = {
                "type": "section",
                "heading": current_parent_heading,
                "label": clean.split()[0] if clean.split() else "Розділ",
                "lines": [line],
            }
            blocks.append(current_block)
        elif is_new_item:
            match = re.match(r"^(\d+[\.\d]*)", clean)
            label = f"Пункт {match.group(1)}" if match else "Пункт"
            current_block = {
                "type": "item",
                "heading": current_parent_heading,
                "label": label,
                "lines": [line],
            }
            blocks.append(current_block)
        else:
            if current_block:
                current_block["lines"].append(line)
            else:
                current_block = {
                    "type": "item",
                    "heading": current_parent_heading,
                    "label": "Основний текст",
                    "lines": [line],
                }
                blocks.append(current_block)

    unit_data_map: dict[str, dict] = {}
    unit_counts: dict[str, int] = {}
    unit_open_names: dict[str, set[str]] = {}

    processed_lines = list(lines)

    for block in blocks:
        if block["type"] != "item":
            continue
        
        block_raw_text = "\n".join(block["lines"])
        block_replaced_lines = list(block["lines"])
        matched_units_in_block = set()

        for open_name, mapped_val in mapping_dict.items():
            if isinstance(mapped_val, dict):
                closed_code = str(mapped_val.get("cipher") or mapped_val.get("closed_name") or open_name)
            else:
                closed_code = str(mapped_val)

            if open_name.casefold() in block_raw_text.casefold():
                pattern = re.compile(re.escape(open_name), re.IGNORECASE)
                block_replaced_lines = [pattern.sub(closed_code, l) for l in block_replaced_lines]
                matched_units_in_block.add((closed_code, open_name))

        if matched_units_in_block:
            full_item_text = "\n".join(block_replaced_lines).strip()
            for closed_code, open_name in matched_units_in_block:
                unit_entry = unit_data_map.setdefault(closed_code, {
                    "unit_code": closed_code,
                    "header_lines": header_lines,
                    "items": [],
                })
                unit_entry["items"].append({
                    "parent_heading": block["heading"],
                    "label": block["label"],
                    "text": full_item_text,
                })
                unit_counts[closed_code] = unit_counts.get(closed_code, 0) + 1
                unit_open_names.setdefault(closed_code, set()).add(open_name)

    table_rows = []
    for closed_code, count in unit_counts.items():
        open_str = ", ".join(sorted(unit_open_names.get(closed_code, [])))
        entry = unit_data_map[closed_code]
        p_labels = ", ".join(dict.fromkeys(item["label"] for item in entry["items"]))
        table_rows.append((closed_code, open_str, p_labels, count))

    table = DataTable(
        ("Закрита назва (ВЧ)", "Відкрита назва", "Пункти наказу", "Згадок"),
        tuple(table_rows),
        "Розпізнані військові частини",
    )

    final_text = "\n".join(processed_lines)
    summary = f"Знайдено частин: {len(unit_counts)} · Загалом згадок: {sum(unit_counts.values())}"

    return {
        "processed_text": final_text,
        "units_table": table,
        "units_list": list(unit_counts.keys()),
        "unit_paragraphs": unit_data_map,
        "summary": summary,
    }
