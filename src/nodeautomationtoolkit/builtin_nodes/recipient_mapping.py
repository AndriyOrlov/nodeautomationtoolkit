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
        "Читає CSV/XLSX таблицю за колонками A (Відкрита назва), B (Закрите найменування / в/ч), "
        "C (Скорочення), D (Корпус). Якщо колонка D вказана — відправником є Корпус; якщо порожня — частина."
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

    # ── Визначаємо індекси колонок або позиційний доступ A, B, C, D, E ──────────
    header_index = -1
    col_a_idx, col_b_idx = 0, 1
    col_c_idx, col_d_idx, col_e_idx = -1, -1, -1
    has_explicit_abbreviation = False

    for index, row in enumerate(rows[:10]):
        row_norm = {cell.strip().casefold(): i for i, cell in enumerate(row) if cell.strip()}
        if any("відкрит" in k or "найменування" in k for k in row_norm) or any("шифр" in k for k in row_norm):
            header_index = index
            for k, i in row_norm.items():
                if "відкрит" in k or "назва" in k or "найменування" in k:
                    col_a_idx = i
                elif "шифр" in k or "закрит" in k:
                    col_b_idx = i
                elif "скороч" in k or "абрев" in k:
                    col_c_idx = i
                    has_explicit_abbreviation = True
                elif "корпус" in k:
                    col_d_idx = i
                elif "дислокац" in k or "адрес" in k or "витяг" in k or "куди" in k:
                    col_e_idx = i
            break

    data_rows = rows[header_index + 1 :] if header_index >= 0 else rows

    mapping: dict[str, dict[str, str]] = {}
    table_rows = []

    for row in data_rows:
        if not any(row):
            continue
        open_name = row[col_a_idx].strip() if col_a_idx < len(row) else ""
        cipher = row[col_b_idx].strip() if col_b_idx < len(row) else ""

        # Якщо таблиця має 4+ колонок (A, B, C, D, E):
        if len(row) >= 4 or has_explicit_abbreviation:
            c_pos = col_c_idx if col_c_idx >= 0 else 2
            d_pos = col_d_idx if col_d_idx >= 0 else 3
            e_pos = col_e_idx if col_e_idx >= 0 else 4
            abbreviation = row[c_pos].strip() if c_pos < len(row) else ""
            corps = row[d_pos].strip() if d_pos < len(row) else ""
            dislocation = row[e_pos].strip() if e_pos < len(row) else ""
            destination = dislocation if dislocation else (corps if corps else (cipher or open_name))
        else:
            abbreviation = ""
            corps = ""
            dislocation = ""
            destination = row[2].strip() if len(row) > 2 else (cipher or open_name)

        if not cipher and not abbreviation and not corps and not dislocation:
            continue
        if not open_name:
            open_name = abbreviation or cipher

        entry = {
            "open_name": open_name,
            "cipher": cipher or open_name,
            "abbreviation": abbreviation,
            "corps": corps,  # Якщо D порожнє -> "", відправник сама частина!
            "dislocation": dislocation,  # Колонка E: Дислокація / куди направляти у витяги
            "destination": destination,
        }

        mapping[open_name] = entry
        if abbreviation and abbreviation != open_name:
            mapping[abbreviation] = entry

        table_rows.append((open_name, cipher, abbreviation, corps, dislocation or destination))

    table = DataTable(
        ("Відкрита назва (A)", "Закрите найменування (B)", "Скорочення (C)", "Корпус (D)", "Дислокація / Куди направляти (E)"),
        tuple(table_rows),
        "Таблиця відповідностей ВЧ",
    )
    return {
        "mapping": mapping,
        "markers": list(mapping),
        "table": table,
        "count": len(table_rows),
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

    pattern_str = r".{0,80}?".join(anchors)
    return re.compile(pattern_str, re.IGNORECASE | re.UNICODE | re.DOTALL)


_CORPS_RE = re.compile(
    r"\b(\d{1,3})[-\s]*(?:-?й|-?го|-?му|-?м)?\s*(?:армійськ\w*\s+корпус\w*|АК)\b",
    re.IGNORECASE | re.UNICODE,
)


def _extract_army_corps(text: str) -> str | None:
    """
    Знаходить армійський корпус у тексті (наприклад: '10-го армійського корпусу', '3 АК').
    Повертає нормалізовану назву: '10 армійський корпус'.
    """
    match = _CORPS_RE.search(text)
    if match:
        num = match.group(1)
        return f"{num} армійський корпус"
    return None


_TCK_OBL_RE = re.compile(
    r"([А-ЯІЇЄa-ua-z]+сько\w*)\s+област\w*",
    re.IGNORECASE | re.UNICODE,
)


def _extract_tck_sender(text: str) -> str | None:
    """
    Якщо у тексті згадується ТЦК (районний / міський), відправником завжди є ОБЛАСТЬ (ОТЦК).
    Наприклад: 'Ковельського районного ТЦК та СП Волинської області' -> 'Волинський ОТЦК та СП'.
    """
    if "ТЦК" not in text.upper() and "комплектування" not in text.lower():
        return None

    match = _TCK_OBL_RE.search(text)
    if match:
        region = match.group(1).strip()
        region_base = re.sub(r"ської$", "ський", region, flags=re.IGNORECASE)
        region_base = re.sub(r"скої$", "ський", region_base, flags=re.IGNORECASE)
        return f"{region_base} ОТЦК та СП"

    return "Обласний ТЦК та СП"



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
    unit_patterns: list[tuple[str, str, str, re.Pattern]] = []
    for open_name, mapped_val in mapping_dict.items():
        corps_col = ""
        if isinstance(mapped_val, dict):
            closed_code = str(
                mapped_val.get("cipher")
                or mapped_val.get("closed_name")
                or open_name
            )
            corps_col = str(mapped_val.get("corps", "")).strip()
            abbreviation = str(mapped_val.get("abbreviation", "")).strip()
        else:
            closed_code = str(mapped_val)
            abbreviation = ""

        if fuzzy_match:
            pattern = _build_unit_fuzzy_pattern(open_name)
        else:
            pattern = re.compile(re.escape(open_name), re.IGNORECASE)

        unit_patterns.append((open_name, closed_code, corps_col, pattern))
        if abbreviation and abbreviation != open_name:
            abbr_pattern = _build_unit_fuzzy_pattern(abbreviation) if fuzzy_match else re.compile(re.escape(abbreviation), re.IGNORECASE)
            unit_patterns.append((abbreviation, closed_code, corps_col, abbr_pattern))

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

        # 1. Зіставлення за патернами з колонок A (повна назва) та C (скорочення)
        for open_name, closed_code, corps_col, pattern in unit_patterns:
            all_matches = pattern.findall(block_raw_text)
            if not all_matches:
                continue

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

            block_replaced_lines = [pattern.sub(closed_code, ln) for ln in block_replaced_lines]
            # Якщо колонка D (corps_col) вказана -> відправником є Корпус з D.
            # Якщо колонка D порожня -> відправником є сама частина (не успадковує корпус з тексту)!
            sender_key = corps_col if corps_col else closed_code
            matched_units_in_block.add((sender_key, open_name))

        # 2. Якщо в таблиці немає відповідностей — перевіряємо ТЦК (районний/міський -> Область)
        if not matched_units_in_block:
            tck_sender = _extract_tck_sender(block_raw_text)
            if tck_sender:
                matched_units_in_block = {(tck_sender, "ТЦК та СП (Область)")}

        # 3. Якщо в таблиці немає відповідностей — перевіряємо чи в тексті явно є АК
        if not matched_units_in_block:
            corps_name = _extract_army_corps(block_raw_text)
            if corps_name:
                matched_units_in_block = {(corps_name, corps_name)}

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




@node(
    name="Аналіз відправників та пунктів наказу",
    category="Наказ",
    description=(
        "Аналізує текст відкритого наказу, знаходить усіх відправників / адресатів "
        "(військові частини, ТЦК та СП, штаби тощо) та прив'язує кожен пункт "
        "до відповідного відправника."
    ),
    type_id="builtin.order.analyze_senders",
    outputs={
        "sender_paragraphs": "dict",
        "senders_list": "list",
        "table": "DataTable",
        "summary": "str",
    },
)
def analyze_senders(
    text: str = "",
    mapping: dict | None = None,
    default_sender_prefix: str = "Відправник ",
) -> dict:
    """Аналізує відкритий наказ і формує відповідність [Відправник / ВЧ] -> [Список пунктів]."""
    res = map_military_units(text=text, mapping=mapping, default_prefix=default_sender_prefix)
    units_map = res.get("unit_paragraphs", {})
    sender_paragraphs: dict[str, list[str]] = {}
    table_rows = []

    for sender, data in units_map.items():
        if isinstance(data, dict) and "items" in data:
            items = [item["text"] for item in data["items"] if "text" in item]
            p_labels = ", ".join(dict.fromkeys(item.get("label", "") for item in data["items"]))
        else:
            items = [str(data)]
            p_labels = "-"
        sender_paragraphs[sender] = items
        table_rows.append((sender, p_labels, len(items)))

    table = DataTable(
        ("Відправник / ВЧ", "Пункти наказу", "Кількість пунктів"),
        tuple(table_rows),
        "Аналіз відправників та пунктів",
    )
    summary = f"Виявлено відправників: {len(sender_paragraphs)} · Опрацьовано пунктів: {sum(len(v) for v in sender_paragraphs.values())}"
    return {
        "sender_paragraphs": sender_paragraphs,
        "senders_list": list(sender_paragraphs.keys()),
        "table": table,
        "summary": summary,
    }


@node(
    name="Розділити наказ за відправниками",
    category="Наказ",
    description=(
        "Формує окремі текстові витяги або блоки наказу для кожного розпізнаного відправника "
        "(ВЧ), зберігаючи шапку та преамбулу (§) оригінального наказу."
    ),
    type_id="builtin.order.split_by_senders",
    outputs={
        "blocks": "dict",
        "senders_count": "int",
        "summary": "str",
    },
)
def split_by_senders(
    text: str = "",
    mapping: dict | None = None,
    header: str = "",
) -> dict:
    """Розділяє відкритий наказ на окремі блоки / витяги за відправниками."""
    res = map_military_units(text=text, mapping=mapping)
    units_map = res.get("unit_paragraphs", {})
    blocks: dict[str, str] = {}

    for sender, data in units_map.items():
        header_text = header
        if not header_text and isinstance(data, dict):
            header_text = "\n".join(data.get("header_lines", []))
        items_text = ""
        if isinstance(data, dict) and "items" in data:
            items_text = "\n\n".join(item.get("text", "") for item in data["items"])
        elif isinstance(data, str):
            items_text = data

        full_doc = f"{header_text}\n\n{items_text}".strip() if header_text else items_text.strip()
        blocks[sender] = full_doc

    summary = f"Створено витягів за відправниками: {len(blocks)}"
    return {
        "blocks": blocks,
        "senders_count": len(blocks),
        "summary": summary,
    }
