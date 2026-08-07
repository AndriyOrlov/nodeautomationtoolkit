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
    execution_inputs=("exec",),
    execution_outputs=("then",),
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


_MILITARY_TYPO_DICTIONARY = [
    # Описи та оддруки у слові "бригада / бригади"
    (re.compile(r"\bбригд[иаоеемся]?\b", re.IGNORECASE), "бригади"),
    (re.compile(r"\bбригаи\b", re.IGNORECASE), "бригади"),
    (re.compile(r"\bбриади?\b", re.IGNORECASE), "бригади"),
    (re.compile(r"\bбргади?\b", re.IGNORECASE), "бригади"),
    # Описи у слові "механізована / механізованої"
    (re.compile(r"\bмеханизоаної\b", re.IGNORECASE), "механізованої"),
    (re.compile(r"\bмеханізоаної\b", re.IGNORECASE), "механізованої"),
    (re.compile(r"\bмеханизованої\b", re.IGNORECASE), "механізованої"),
    (re.compile(r"\bмеханизована\b", re.IGNORECASE), "механізована"),
    # Описи у слові "військова частина / в/ч"
    (re.compile(r"\bв\s*\\\s*ч\b", re.IGNORECASE), "в/ч"),
    (re.compile(r"\bв\.?\s*ч\.?\b", re.IGNORECASE), "в/ч"),
    (re.compile(r"\bзв[’'`ʻ]?язку\b", re.IGNORECASE), "зв'язку"),
    (re.compile(r"\bсвязку\b", re.IGNORECASE), "зв'язку"),
    # Описи у типі бригад
    (re.compile(r"\bгірсько\s+штурмов\w*\b", re.IGNORECASE), "гірсько-штурмової"),
    (re.compile(r"\bгірськоштурмов\w*\b", re.IGNORECASE), "гірсько-штурмової"),
    (re.compile(r"\bдесантно\s+штурмов\w*\b", re.IGNORECASE), "десантно-штурмової"),
    (re.compile(r"\bдесантноштурмов\w*\b", re.IGNORECASE), "десантно-штурмової"),
    # Описи для центр та рекрутинг
    (re.compile(r"\bцнтр[уомівіаб]?\b", re.IGNORECASE), "центру"),
    (re.compile(r"\bрекрутинг[уомівіаб]?\b", re.IGNORECASE), "рекрутингу"),
]


def _fix_military_typos(text: str) -> str:
    """Нормалізує поширені описки та варіанти військових найменувань перед обробкою."""
    if not text:
        return ""
    res = text
    for pat, repl in _MILITARY_TYPO_DICTIONARY:
        res = pat.sub(repl, res)
    return res


def _stem_ukrainian_word(word: str) -> str:
    """Видаляє закінчення українських відмінків та стійка до оддруків (БРИГДИ, БРИГАДИ, МЕХАНИЗОВАНОЇ)."""
    w = word.casefold().strip()
    w = re.sub(r"ін$", "он", w)
    w = re.sub(r"іна$", "она", w)

    # Спеціальна обробка військових термінів для захисту від оддруків у документах
    if w.startswith("бриг"):
        return "бриг"
    if w.startswith("механ"):
        return "механ"
    if w.startswith("гірськ") or w.startswith("гірск") or w.startswith("гірш"):
        return "гірс"
    if w.startswith("десант"):
        return "десант"
    if w.startswith("артил"):
        return "артил"
    if w.startswith("баталь") or w.startswith("батальон"):
        return "бат"
    if w.startswith("центр") or w.startswith("цнтр"):
        return "центр"
    if w.startswith("рекрутинг") or w.startswith("рекрут"):
        return "рекрут"

    endings = [
        "ий", "ій", "ого", "ому", "им", "ім", "ої", "ою", "их", "ними", "ими",
        "ами", "ям", "ями", "ях", "ах", "ом", "ем", "єм", "ів", "ев", "єв",
        "и", "і", "а", "я", "у", "ю", "е", "є", "о"
    ]
    for end in sorted(endings, key=len, reverse=True):
        if w.endswith(end) and len(w) - len(end) >= 2:
            w = w[:-len(end)]
            break
    return w


def _build_unit_fuzzy_pattern(open_name: str) -> re.Pattern:
    """
    Будує нечіткий регулярний вираз для пошуку назви військової частини у будь-яких відмінках.
    Враховує відмінювання (полк, полку, полком, бригади, бригадою), апострофи та розширену відстань між словами.
    """
    STOP_WORDS = {
        "та", "і", "й", "з", "зі", "із", "на", "в", "у", "до", "від", "по", "при", "за",
        "the", "of", "and", "а", "ім", "імені",
        "командний", "командного", "командним", "командному", "пункт", "пункту", "пунктом", "пункті",
        "передовий", "передового", "передовим", "передовому", "запасний", "запасного", "запасним", "запасному",
        "зкп", "гкп", "ппу", "кп", "управління",
    }

    clean_name = re.sub(r"[^\w\s\d]", "'", open_name)
    tokens = [t for t in re.split(r"[^\w\d]+", clean_name) if t]

    anchors: list[str] = []
    for token in tokens:
        if not token:
            continue
        if re.match(r"^\d+$", token):
            anchors.append(re.escape(token))
        elif token.casefold() in STOP_WORDS:
            continue
        elif len(token) >= 2:
            stem = _stem_ukrainian_word(token)
            escaped_stem = re.escape(stem).replace("\\'", r"[^\w\s]?").replace("'", r"[^\w\s]?")
            pattern_part = escaped_stem + r"\w*"
            anchors.append(pattern_part)

    if not anchors:
        return re.compile(re.escape(open_name), re.IGNORECASE)

    pattern_str = r".{0,180}?".join(anchors)
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


_RECRUITING_CENTER_9_RE = re.compile(
    r"\b9[-\s]*(?:-?й|-?го|-?му|-?м)?\s*(?:центр\w*|цнтр\w*)\s+рекрутинг\w*\b",
    re.IGNORECASE | re.UNICODE,
)


_TCK_KEYWORDS_RE = re.compile(
    r"(?:територіальн\w*\s+центр\w*\s+комплектування\w*|ТЦК\w*|РТЦК\w*|МТЦК\w*|ОТЦК\w*)",
    re.IGNORECASE | re.UNICODE,
)

_TCK_OBLAST_EXPLICIT_RE = re.compile(
    r"\b([А-ЯІЇЄа-яіїє'ʼ-]+?(?:ськ|цьк|зьк)\w*)\s+област",
    re.IGNORECASE | re.UNICODE,
)

_TCK_REGION_BEFORE_RE = re.compile(
    r"\b([А-ЯІЇЄа-яіїє'ʼ-]+?(?:ськ|цьк|зьк)\w*)"
    r"\s+(?:районн\w*|міськ\w*|обласн\w*|)?\s*"
    r"(?:територіальн\w*\s+центр\w*\s+комплектування\w*|[РМОО]?ТЦК\w*)",
    re.IGNORECASE | re.UNICODE,
)

_RAYON_TO_OBLAST_MAP = {
    "броварськ": "Київський",
    "білоцерківськ": "Київський",
    "бориспільськ": "Київський",
    "вишгородськ": "Київський",
    "фастівськ": "Київський",
    "обухівськ": "Київський",
    "бучанськ": "Київський",
    "ковельськ": "Волинський",
    "володимирськ": "Волинський",
    "камінь-каширськ": "Волинський",
    "дрогобицьк": "Львівський",
    "стрийськ": "Львівський",
    "самбірськ": "Львівський",
    "червоноградськ": "Львівський",
    "золочівськ": "Львівський",
    "яворівськ": "Львівський",
    "криворізьк": "Дніпропетровський",
    "нікопольськ": "Дніпропетровський",
    "павлоградськ": "Дніпропетровський",
    "новомосковськ": "Дніпропетровський",
    "ізмаїльськ": "Одеський",
    "болградськ": "Одеський",
    "білгород-дністровськ": "Одеський",
    "ізюмськ": "Харківський",
    "куп'янськ": "Харківський",
    "чугуївськ": "Харківський",
    "кременчуцьк": "Полтавський",
    "лубенськ": "Полтавський",
    "миргородськ": "Полтавський",
    "бердичівськ": "Житомирський",
    "коростенськ": "Житомирський",
    "новоград-волинськ": "Житомирський",
    "звагельськ": "Житомирський",
    "ужгородськ": "Закарпатський",
    "мукачівськ": "Закарпатський",
    "уманськ": "Черкаський",
    "прилуцьк": "Чернігівський",
    "ніжинськ": "Чернігівський",
    "конотопськ": "Сумський",
    "шосткинськ": "Сумський",
    "сарненськ": "Рівненський",
    "дубенськ": "Рівненський",
    "хмільницьк": "Вінницький",
    "жмеринськ": "Вінницький",
}

_UKRAINE_OBLAST_STEMS = {
    "вінниц": "Вінницький",
    "волин": "Волинський",
    "дніпропетровськ": "Дніпропетровський",
    "донец": "Донецький",
    "житомир": "Житомирський",
    "закарпат": "Закарпатський",
    "запорізьк": "Запорізький",
    "івано-франківськ": "Івано-Франківський",
    "київськ": "Київський",
    "кіровоград": "Кіровоградський",
    "луган": "Луганський",
    "львів": "Львівський",
    "миколаїв": "Миколаївський",
    "одеськ": "Одеський",
    "полтав": "Полтавський",
    "рівнен": "Рівненський",
    "сум": "Сумський",
    "тернопіл": "Тернопільський",
    "харків": "Харківський",
    "херсон": "Херсонський",
    "хмельниц": "Хмельницький",
    "черкас": "Черкаський",
    "чернівец": "Чернівецький",
    "чернігів": "Чернігівський",
}


def _normalize_region_to_nominative(region_raw: str) -> str:
    """Нормалізує прикметник регіону до називного відмінку чоловічого роду (напр. Камінь-Каширського → Камінь-Каширський)."""
    parts = region_raw.strip().split("-")
    norm_parts = []
    for part in parts:
        p = part.strip()
        p = re.sub(r"(?:ського|ської|ському|ським|ською)$", "ський", p, flags=re.IGNORECASE)
        p = re.sub(r"(?:цького|цької|цькому|цьким|цькою)$", "цький", p, flags=re.IGNORECASE)
        p = re.sub(r"(?:зького|зької|зькому|зьким|зькою)$", "зький", p, flags=re.IGNORECASE)
        if not (p.lower().endswith("ський") or p.lower().endswith("цький") or p.lower().endswith("зький") or p.lower().endswith("ін")):
            p += "ський"
        norm_parts.append(p[0].upper() + p[1:] if p else "")
    return "-".join(norm_parts)



def _extract_tck_sender(text: str) -> str | None:
    """Розпізнає ТЦК та СП, а також 9 Центр рекрутингу (тимчасове правило розсилання на ТЦК)."""
    is_tck = bool(_TCK_KEYWORDS_RE.search(text))
    is_recruiting_9 = bool(_RECRUITING_CENTER_9_RE.search(text))

    if not (is_tck or is_recruiting_9):
        return None

    # Пріоритет 1: Якщо у тексті прямо вказано назву області (напр: Київської області)
    oblast_match = _TCK_OBLAST_EXPLICIT_RE.search(text)
    if oblast_match:
        reg_nom = _normalize_region_to_nominative(oblast_match.group(1))
        for stem, obl_name in _UKRAINE_OBLAST_STEMS.items():
            if reg_nom.lower().startswith(stem):
                return f"{obl_name} ОТЦК та СП"
        return f"{reg_nom} ОТЦК та СП"

    # Пріоритет 2: Шукаємо прикметник БЕЗПОСЕРЕДНЬО перед ТЦК (напр: Броварського РТЦК)
    region_match = _TCK_REGION_BEFORE_RE.search(text)
    if region_match:
        reg_nom = _normalize_region_to_nominative(region_match.group(1))
        low = reg_nom.lower()
        # Перевіряємо чи це область
        for stem, obl_name in _UKRAINE_OBLAST_STEMS.items():
            if low.startswith(stem):
                return f"{obl_name} ОТЦК та СП"
        # Перевіряємо чи це район відомої області
        for r_stem, obl_name in _RAYON_TO_OBLAST_MAP.items():
            if low.startswith(r_stem):
                return f"{obl_name} ОТЦК та СП"
        return f"{reg_nom} ОТЦК та СП"

    return "Обласний ТЦК та СП"

def _extract_corps_abbr(corps_str: str) -> str:
    """Витягує скорочену назву корпусу (напр. '25 армійський корпус' -> '25АК')."""
    clean = str(corps_str).strip()
    match = re.search(r"\b(\d{1,3})\s*(?:армійськ\w*\s+корпус\w*|АК)\b", clean, re.IGNORECASE)
    if match:
        return f"{match.group(1)}АК"
    return clean


def _normalize_key(raw_k: str, canonical_map: dict[str, str] | None = None) -> str:
    k = str(raw_k).strip()
    cmap = canonical_map or {}
    if k in cmap:
        return cmap[k]
    short_k = _short_closed_code(k)
    if short_k in cmap:
        return cmap[short_k]
    c_abbr = _extract_corps_abbr(k)
    if c_abbr in cmap:
        return cmap[c_abbr]
    return k


def _auto_abbreviate_unit_name(open_name: str) -> str:
    """Автоматично генерує скорочення ВЧ з повної назви (напр. '15 окрема механізована бригада' -> '15омбр')."""
    clean = str(open_name).strip()
    match_num = re.search(r"\b(\d{1,4})\b", clean)
    num_str = match_num.group(1) if match_num else ""

    low = clean.lower()
    type_abbr = ""
    if "механізован" in low or "омбр" in low:
        type_abbr = "омбр"
    elif "танков" in low or "отбр" in low:
        type_abbr = "отбр"
    elif "десантно-штурмов" in low or "одшбр" in low:
        type_abbr = "одшбр"
    elif "гірсько-штурмов" in low or "огшбр" in low:
        type_abbr = "огшбр"
    elif "артилерійськ" in low or "оабр" in low:
        type_abbr = "оабр"
    elif "полк зв" in low or "зв’язку" in low or "зв'язку" in low or "опз" in low:
        type_abbr = "опз"
    elif "розвідувальн" in low or "орб" in low:
        type_abbr = "орб"
    elif "інженерн" in low or "оібр" in low:
        type_abbr = "оібр"
    elif "автомобільн" in low or "оаб" in low:
        type_abbr = "оаб"
    elif "центр рекрутинг" in low or "цр" in low:
        type_abbr = "ЦР"
    elif "батальйон" in low:
        type_abbr = "ОБ"

    if num_str and type_abbr:
        return f"{num_str}{type_abbr}"
    return clean


def _find_corps_entry(corps_col_val: str, corps_abbr_val: str, mapping_dict: dict) -> dict | None:
    """Шукає рядок корпусу у таблиці за ключем, скороченням або повною назвою."""
    entry = mapping_dict.get(corps_col_val) or mapping_dict.get(corps_abbr_val)
    if isinstance(entry, dict) and entry.get("cipher"):
        return entry
    for _name, _val in mapping_dict.items():
        if not isinstance(_val, dict):
            continue
        abbr = str(_val.get("abbreviation") or _val.get("abbr") or "").strip()
        if abbr and _extract_corps_abbr(abbr) == corps_abbr_val and _val.get("cipher"):
            return _val
        if _extract_corps_abbr(_name) == corps_abbr_val and _val.get("cipher"):
            return _val
    return None


def _build_sender_key(closed_code: str, abbreviation: str, corps_col: str, corps_resolved_cipher: dict, mapping_dict: dict) -> tuple[str, str]:
    """
    Будує точний sender_key для ВЧ/Корпусу, повністю зберігаючи ТЕКСТ З КЛІТИНКИ СКОРОЧЕНОЇ НАЗВИ (колонка C).
    Усуває дублювання шифрів та усуває дублі у підсумковому виводі.
    """
    short_cipher = _short_closed_code(closed_code)

    if corps_col:
        corps_abbr = _extract_corps_abbr(corps_col)
        corps_cipher = corps_resolved_cipher.get(corps_abbr, short_cipher)
        corps_entry = _find_corps_entry(corps_col, corps_abbr, mapping_dict)
        if corps_entry:
            c_abbr = str(corps_entry.get("abbreviation") or corps_entry.get("abbr") or corps_abbr).strip()
            if c_abbr:
                if re.search(r"\bА\s*\d{4}\b", c_abbr, re.IGNORECASE) or (corps_cipher and corps_cipher.casefold() in c_abbr.casefold()):
                    sender_key = c_abbr
                elif corps_cipher:
                    sender_key = f"{c_abbr} {corps_cipher}"
                else:
                    sender_key = c_abbr
            else:
                sender_key = f"{corps_abbr} {corps_cipher}".strip() if corps_cipher else corps_abbr
        else:
            sender_key = f"{corps_abbr} {corps_cipher}".strip() if corps_cipher else corps_abbr
        return sender_key, corps_abbr

    abbr = str(abbreviation).strip()
    if not abbr or abbr.isdigit():
        return short_cipher, short_cipher

    # Для ТЦК та СП — повертаємо скорочення як є
    if "ТЦК" in abbr.upper() or "ОТЦК" in abbr.upper() or "РТЦК" in abbr.upper() or "МТЦК" in abbr.upper():
        return abbr, abbr

    # Якщо у клітинці скорочення (колонка C) ВЖЕ є шифр (напр. '24ОМБр А2424') — копіюємо ВСЕ з клітинки
    if re.search(r"\bА\s*\d{4}\b", abbr, re.IGNORECASE) or (short_cipher and short_cipher.casefold() in abbr.casefold()):
        return abbr, abbr

    # Інакше додаємо шифр до скорочення з колонки C: '24ОМБр А2424'
    return f"{abbr} {short_cipher}", abbr


def _short_closed_code(code: str, abbreviation: str = "") -> str:
    """Формує коротке закрите найменування частини або ТЦК (усуває дублювання довгих назв ТЦК)."""
    clean_code = str(code).strip()
    if clean_code.lower().startswith("в/ч "):
        cipher = clean_code[4:].strip()
    elif clean_code.lower().startswith("в/ч"):
        cipher = clean_code[3:].strip()
    else:
        cipher = clean_code

    abbr = str(abbreviation).strip()

    # Для ТЦК та СП — повертаємо лаконічну назву (напр. 'Київський ОТЦК та СП') без дублювання довгої назви
    if "ТЦК" in abbr.upper() or "ТЦК" in cipher.upper() or "ТЕРИТОРІАЛЬН" in cipher.upper() or "ЦЕНТЕР" in cipher.upper() or "ЦЕНТР" in cipher.upper():
        if abbr:
            return abbr
        reg = re.search(r"\b([А-ЯІЇЄ][а-яіїє'ʼ]+(?:ськ|цьк|зьк)\w*)\b", cipher, re.IGNORECASE)
        if reg:
            reg_nom = _normalize_region_to_nominative(reg.group(1))
            return f"{reg_nom} ОТЦК та СП"
        return "Обласний ТЦК та СП"

    if abbr and abbr != cipher and abbr.casefold() not in cipher.casefold():
        return f"{abbr} {cipher}"
    return cipher


def _format_item_numbers_range(labels: list[str]) -> str:
    """Форматує перелік пунктів з використанням дефісів для послідовних діапазонів (напр. 1,3,4-7,9,10,11-25)."""
    nums = set()
    non_nums = []

    for label in labels:
        clean = str(label).strip()
        matches = re.findall(r"\b\d+\b", clean)
        if matches:
            nums.add(int(matches[0]))
        elif clean:
            non_nums.append(clean)

    sorted_nums = sorted(nums)
    if not sorted_nums:
        return ",".join(non_nums) if non_nums else "-"

    ranges = []
    start = sorted_nums[0]
    end = sorted_nums[0]

    for n in sorted_nums[1:]:
        if n == end + 1:
            end = n
        else:
            if end > start + 1:
                ranges.append(f"{start}-{end}")
            elif end == start + 1:
                ranges.append(f"{start},{end}")
            else:
                ranges.append(f"{start}")
            start = n
            end = n

    if end > start + 1:
        ranges.append(f"{start}-{end}")
    elif end == start + 1:
        ranges.append(f"{start},{end}")
    else:
        ranges.append(f"{start}")

    result_str = ",".join(ranges)
    if non_nums:
        result_str += "," + ",".join(non_nums)
    return result_str



def _get_item_main_text(lines: list[str]) -> str:
    """
    Витягує тільки основний абзац пункту (до опису року народження 'р.н.' або 'Підстава').
    Рядки з р.н., РНОКПП та Підстава повністю ігноруються при визначенні ВЧ для витягів.
    """
    main_lines = []
    for line in lines:
        clean = line.strip()
        low = clean.lower()
        if "р.н." in low or "р. н." in low or "року народження" in low or clean.startswith("Підстава") or clean.startswith("підстава"):
            break
        main_lines.append(line)
    return "\n".join(main_lines) if main_lines else "\n".join(lines)


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

    text = _fix_military_typos(text)
    mapping_raw = mapping or {}
    mapping_dict = {}
    for k, v in mapping_raw.items():
        clean_k = _fix_military_typos(str(k))
        if isinstance(v, dict):
            clean_v = dict(v)
            if "open_name" in clean_v:
                clean_v["open_name"] = _fix_military_typos(str(clean_v["open_name"]))
            mapping_dict[clean_k] = clean_v
        else:
            mapping_dict[clean_k] = v

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
            or "НАПРАВИТИ" in clean.upper()
            or "ВІДРЯДИТИ" in clean.upper()
            or "ЗВІЛЬНИТИ" in clean.upper()
            or "ВІЙСЬКОВОСЛУЖБОВЦІВ" in clean.upper()
        ):
            content_start_idx = idx
            break
        if clean:
            header_lines.append(line)

    if not header_lines and lines:
        header_lines = lines[:3]

    # ── Компілюємо патерни для кожної ВЧ ──────────────────────────────────────
    unit_patterns: list[tuple[str, str, str, re.Pattern]] = []
    unit_abbr_map: dict[str, str] = {}

    canonical_key_map: dict[str, str] = {}
    cipher_digits_map: dict[str, tuple[str, str]] = {}

    # ── Прохід 1: визначаємо ЄДИНИЙ шифр для кожного Корпусу ─────────────────
    corps_resolved_cipher: dict[str, str] = {}  # corps_abbr -> шифр корпусу

    for open_name, mapped_val in mapping_dict.items():
        if not isinstance(mapped_val, dict):
            continue
        corps_col = str(mapped_val.get("corps", "")).strip()
        if corps_col:
            corps_abbr = _extract_corps_abbr(corps_col)
            if corps_abbr not in corps_resolved_cipher:
                corps_entry = _find_corps_entry(corps_col, corps_abbr, mapping_dict)
                if corps_entry:
                    corps_resolved_cipher[corps_abbr] = _short_closed_code(str(corps_entry["cipher"]))
                else:
                    # Якщо корпус не має окремого рядка — не підставляємо шифр підпорядкованої частини!
                    corps_resolved_cipher[corps_abbr] = ""

    # ── Прохід 2: компілюємо патерни з єдиним sender_key ─────────────────────
    for open_name, mapped_val in mapping_dict.items():
        corps_col = ""
        if isinstance(mapped_val, dict):
            closed_code = str(
                mapped_val.get("cipher")
                or mapped_val.get("closed_name")
                or open_name
            )
            corps_col = str(mapped_val.get("corps", "")).strip()
            abbreviation = str(
                mapped_val.get("abbreviation")
                or mapped_val.get("abbr")
                or mapped_val.get("short_name")
                or mapped_val.get("скорочення")
                or mapped_val.get("скорочена_назва")
                or ""
            ).strip()
        else:
            closed_code = str(mapped_val)
            abbreviation = ""

        low_open = open_name.lower()
        is_tck_entry = "тцк" in low_open or "територіальн" in low_open or "центр" in low_open
        if isinstance(mapped_val, dict) and not abbreviation and not corps_col and not is_tck_entry:
            abbreviation = _auto_abbreviate_unit_name(open_name)

        short_cipher = _short_closed_code(closed_code)
        sender_key, resolved_abbr = _build_sender_key(closed_code, abbreviation, corps_col, corps_resolved_cipher, mapping_dict)
        unit_abbr_map[sender_key] = resolved_abbr or unit_abbr_map.get(sender_key, "")

        for variant in [open_name, closed_code, f"в/ч {short_cipher}", short_cipher, abbreviation, corps_col]:
            if variant:
                canonical_key_map[variant] = sender_key
                if corps_col:
                    corps_abbr = _extract_corps_abbr(corps_col)
                    canonical_key_map[corps_abbr] = sender_key

        # Реєструємо 4-значні цифри шифру для миттєвого пошуку прямо за шифром в/ч у тексті
        c_digits = re.findall(r"\d{4}", closed_code)
        if c_digits:
            cipher_digits_map[c_digits[0]] = (sender_key, open_name)



        low_open = open_name.lower()
        is_tck_entry = "тцк" in low_open or "територіальн" in low_open or "центр" in low_open
        if fuzzy_match or is_tck_entry:
            pattern = _build_unit_fuzzy_pattern(open_name)
        else:
            pattern = re.compile(re.escape(open_name), re.IGNORECASE)

        unit_patterns.append((open_name, closed_code, corps_col, sender_key, pattern))
        if abbreviation and abbreviation != open_name:
            abbr_pattern = _build_unit_fuzzy_pattern(abbreviation) if fuzzy_match else re.compile(re.escape(abbreviation), re.IGNORECASE)
            unit_patterns.append((abbreviation, closed_code, corps_col, sender_key, abbr_pattern))

        if closed_code and closed_code != open_name:
            code_pattern = _build_unit_fuzzy_pattern(closed_code) if fuzzy_match else re.compile(re.escape(closed_code), re.IGNORECASE)
            unit_patterns.append((closed_code, closed_code, corps_col, sender_key, code_pattern))

        if short_cipher and short_cipher != open_name and short_cipher != closed_code:
            short_pattern = re.compile(r"\b" + re.escape(short_cipher) + r"\b", re.IGNORECASE)
            unit_patterns.append((short_cipher, closed_code, corps_col, sender_key, short_pattern))

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
            ("Відповідно до" in clean or "Згідно з" in clean or clean.endswith(":") or "ВІЙСЬКОВОСЛУЖБОВЦІВ" in clean.upper())
            and not re.match(r"^\d+[\.\d]*", clean)
            and not clean.startswith("Підстава")
            and not clean.startswith("підстава")
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
    active_section_units: set[tuple[str, str]] = set()

    for block in blocks:
        if block["type"] == "section":
            section_raw_text = "\n".join(block["lines"])
            sec_units: set[tuple[str, str]] = set()

            # Шукаємо ЗВІДКИ (вихідна ВЧ) та КУДИ (цільова ВЧ) у шапці розділу
            match_kudy = re.search(
                r"(?:призначити\s+до|направити\s+до|відрядити\s+до|у\s+розпорядження)\s+([^:\n\.]+)",
                section_raw_text,
                re.IGNORECASE,
            )
            kudy_text = match_kudy.group(1).strip() if match_kudy else ""

            # 1. Знаходимо ВЧ напрямку КУДИ
            kudy_units: set[tuple[str, str]] = set()
            if kudy_text and "цього саг" not in kudy_text.lower() and "цієї саг" not in kudy_text.lower() and "того ж" not in kudy_text.lower():
                for open_name, closed_code, corps_col, sender_key, pattern in unit_patterns:
                    m = pattern.search(kudy_text)
                    if m:
                        matched_str = m.group(0)
                        low_open = open_name.lower()
                        is_tck_entry = "тцк" in low_open or "територіальн" in low_open or "центр" in low_open
                        target_name = sender_key if is_tck_entry else (corps_col if (corps_col and ("корпус" in matched_str.lower() or "ак" in matched_str.lower())) else open_name)
                        kudy_units.add((sender_key, target_name))
                if not kudy_units:
                    sec_tck = _extract_tck_sender(kudy_text)
                    if sec_tck:
                        kudy_units.add((sec_tck, "ТЦК та СП"))

            # 2. Знаходимо ВЧ напрямку ЗВІДКИ (або загальну ВЧ у шапці)
            zvidky_units: set[tuple[str, str]] = set()
            for open_name, closed_code, corps_col, sender_key, pattern in unit_patterns:
                m = pattern.search(section_raw_text)
                if m:
                    matched_str = m.group(0)
                    low_open = open_name.lower()
                    is_tck_entry = "тцк" in low_open or "територіальн" in low_open or "центр" in low_open
                    target_name = sender_key if is_tck_entry else (corps_col if (corps_col and ("корпус" in matched_str.lower() or "ак" in matched_str.lower())) else open_name)
                    zvidky_units.add((sender_key, target_name))
                    break

            sec_units = (zvidky_units | kudy_units)
            if len(sec_units) > 2:
                sec_units = set(list(sec_units)[:2])

            if sec_units:
                active_section_units = sec_units
            continue

        if block["type"] != "item":
            continue

        # Очищаємо варіанти префіксів командних пунктів для точного зіставлення
        raw_text_clean = re.sub(
            r"\b(?:передов\w*|запасн\w*|головн\w*|)?\s*"
            r"(?:командн\w*\s+пункт\w*|пункт\w*\s+управління|ЗКП|ГКП|ППУ|КП)\s+",
            "",
            "\n".join(block["lines"]),
            flags=re.IGNORECASE | re.UNICODE,
        )
        block_raw_text = _get_item_main_text(block["lines"])
        block_replaced_lines = list(block["lines"])
        matched_units_in_block: set[tuple[str, str]] = set()

        # Якщо є активний заголовок розділу з ВЧ — пункти за замовчуванням мають її адресата
        if active_section_units:
            matched_units_in_block = set(active_section_units)

        if not matched_units_in_block:
            # 1. Зіставлення за патернами з колонок A (повна назва) та C (скорочення)
            for open_name, closed_code, corps_col, sender_key, pattern in unit_patterns:
                all_matches = pattern.findall(block_raw_text) or pattern.findall(raw_text_clean)
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
                matched_str = str(all_matches[0]) if all_matches else ""
                low_open = open_name.lower()
                is_tck_entry = "тцк" in low_open or "територіальн" in low_open or "центр" in low_open
                if is_tck_entry:
                    target_name = sender_key
                elif corps_col and ("корпус" in matched_str.lower() or "ак" in matched_str.lower()):
                    target_name = corps_col
                else:
                    target_name = open_name
                matched_units_in_block.add((sender_key, target_name))
                break  # 1 ПУНКТ НАКАЗУ НАЛЕЖИТЬ МАКСИМУМ 1 ОДНОМУ АДРЕСАТУ!



        if matched_units_in_block and len(matched_units_in_block) > 2:
            matched_units_in_block = set(list(matched_units_in_block)[:2])

        # 2. Пошук за цифровими шифрами в/ч (наприклад: військової частини А2424 або А 2828)
        if not matched_units_in_block:
            found_ciphers = re.findall(r"\bА\s*(\d{4})\b", block_raw_text, re.IGNORECASE)
            for digit in found_ciphers:
                if digit in cipher_digits_map:
                    s_key, o_name = cipher_digits_map[digit]
                    matched_units_in_block.add((s_key, o_name))

        # 3. Якщо в таблиці немає відповідностей — перевіряємо ТЦК (районний/міський -> Область)
        if not matched_units_in_block:
            tck_sender = _extract_tck_sender(block_raw_text)
            if tck_sender:
                matched_units_in_block = {(tck_sender, "ТЦК та СП (Область)")}

        # 4. Якщо в таблиці немає відповідностей — перевіряємо чи в тексті явно є АК
        if not matched_units_in_block:
            corps_name = _extract_army_corps(block_raw_text)
            if corps_name:
                corps_abbr = _extract_corps_abbr(corps_name)
                c_key = canonical_key_map.get(corps_name) or canonical_key_map.get(corps_abbr) or corps_abbr
                matched_units_in_block = {(c_key, corps_name)}

        # 5. Унаслідування адресата з активного заголовка розділу/шапки (§), якщо у пункті немає власної ВЧ
        if not matched_units_in_block and active_section_units:
            matched_units_in_block = set(active_section_units)

        if matched_units_in_block:
            full_item_text = "\n".join(block_replaced_lines).strip()
            # Дедуплікація: один пункт наказу додається лише ОДИН раз на кожен унікальний norm_code
            seen_norm_codes: set[str] = set()
            for raw_code, open_name in matched_units_in_block:
                norm_code = _normalize_key(raw_code, canonical_key_map)
                if norm_code in seen_norm_codes:
                    continue
                seen_norm_codes.add(norm_code)
                unit_entry = unit_data_map.setdefault(
                    norm_code,
                    {
                        "unit_code": norm_code,
                        "abbreviation": unit_abbr_map.get(norm_code, ""),
                        "header_lines": header_lines,
                        "items": [],
                    },
                )
                if "abbreviation" not in unit_entry or not unit_entry["abbreviation"]:
                    unit_entry["abbreviation"] = unit_abbr_map.get(norm_code, "")
                unit_entry["items"].append(
                    {
                        "parent_heading": block["heading"],
                        "label": block["label"],
                        "text": full_item_text,
                    }
                )
                unit_counts[norm_code] = unit_counts.get(norm_code, 0) + 1
                unit_open_names.setdefault(norm_code, set()).add(open_name)

    # ── Будуємо вихідні таблиці ────────────────────────────────────────────────
    table_rows = []
    for closed_code, count in unit_counts.items():
        open_str = ", ".join(sorted(unit_open_names.get(closed_code, [])))
        entry = unit_data_map[closed_code]
        raw_labels = [item.get("label", "") for item in entry["items"]]
        range_labels = _format_item_numbers_range(raw_labels)
        abbr = unit_abbr_map.get(closed_code, "")
        short_code = _short_closed_code(closed_code, abbreviation=abbr)
        table_rows.append((short_code, open_str, range_labels, count))

    table = DataTable(
        ("Закрита назва (ВЧ)", "Відкрита назва", "Номери пунктів витягу", "Кількість пунктів"),
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
        "unit_abbr_map": unit_abbr_map,
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
    abbr_map = res.get("unit_abbr_map", {})
    sender_paragraphs: dict[str, list[str]] = {}
    table_rows = []

    for sender, data in units_map.items():
        short_sender = sender
        if isinstance(data, dict) and "items" in data:
            items = [item["text"] for item in data["items"] if "text" in item]
            raw_labels = [item.get("label", "") for item in data["items"]]
            range_labels = _format_item_numbers_range(raw_labels)
        else:
            items = [str(data)]
            range_labels = "-"
        sender_paragraphs[sender] = items
        table_rows.append((short_sender, range_labels, len(items)))

    table = DataTable(
        ("Військова частина / Відправник", "Номери пунктів витягу", "Кількість пунктів"),
        tuple(table_rows),
        "Аналіз відправників та пунктів витягу",
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
        "table": "DataTable",
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
    abbr_map = res.get("unit_abbr_map", {})
    blocks: dict[str, str] = {}
    table_rows = []

    for sender, data in units_map.items():
        header_text = header
        if not header_text and isinstance(data, dict):
            header_text = "\n".join(data.get("header_lines", []))

        short_sender = sender
        items_text = ""
        range_labels = "-"
        count = 0
        if isinstance(data, dict) and "items" in data:
            # Групуємо пункти під відповідними преамбулами/заголовками розділів (§ 1 або ЗВІЛЬНИТИ...)
            section_groups: dict[str, list[str]] = {}
            raw_labels = []
            for item in data.get("items", []):
                heading = item.get("parent_heading", "").strip()
                item_text = item.get("text", "").strip()
                label = item.get("label", "")
                if label:
                    raw_labels.append(label)
                if item_text:
                    section_groups.setdefault(heading, []).append(item_text)

            range_labels = _format_item_numbers_range(raw_labels)
            count = len(data.get("items", []))

            body_sections = []
            for heading, item_texts in section_groups.items():
                if heading:
                    body_sections.append(f"{heading}\n\n" + "\n\n".join(item_texts))
                else:
                    body_sections.append("\n\n".join(item_texts))

            items_text = "\n\n".join(body_sections)
        elif isinstance(data, str):
            items_text = data
            count = 1

        full_doc = f"{header_text}\n\n{items_text}".strip() if header_text else items_text.strip()
        blocks[sender] = full_doc
        table_rows.append((short_sender, range_labels, count))

    table = DataTable(
        ("Військова частина (Адресат)", "Номери пунктів витягу", "Кількість пунктів"),
        tuple(table_rows),
        "Перелік відправників та пунктів витягу",
    )
    summary = f"Створено витягів за відправниками: {len(blocks)}"
    return {
        "blocks": blocks,
        "table": table,
        "senders_count": len(blocks),
        "summary": summary,
    }


# ── Правила заміни займенниково-іменникових зворотів ───────────────────────────
_UNIT_PHRASE_REPLACEMENTS = [
    # (цієї / цієї самої / зазначеної / вказаної / даної) (бригади / полку / батальйону ...)
    (
        re.compile(
            r"\b(цієї|цієї\s+самої|цього|цього\s+самого|зазначеної|зазначеного|вказаної|вказаного|даної|даного)\s+(бригади|полку|батальйону|дивізіону|загону|корпусу|центру)\b",
            re.IGNORECASE | re.UNICODE,
        ),
        lambda m: (
            "цієї самої військової частини" if "само" in m.group(0).lower()
            else ("цієї військової частини" if m.group(1).lower() in ("цієї", "цього")
            else ("зазначеної військової частини" if "зазнач" in m.group(1).lower()
            else ("вказаної військової частини" if "вказан" in m.group(1).lower()
            else "даної військової частини")))
        )
    ),
    # (цією / цією самою / цим / цим самим) (бригадою / полком ...)
    (
        re.compile(
            r"\b(цією|цією\s+самою|цим|цим\s+самим|зазначеною|зазначеним|вказаною|вказаним|даною|даним)\s+(бригадою|полком|батальйоном|дивізіоном|загоном|корпусом)\b",
            re.IGNORECASE | re.UNICODE,
        ),
        lambda m: (
            "цією самою військовою частиною" if "само" in m.group(0).lower()
            else ("цією військовою частиною" if m.group(1).lower() in ("цією", "цим")
            else ("зазначеною військовою частиною" if "зазнач" in m.group(1).lower()
            else ("вказаною військовою частиною" if "вказан" in m.group(1).lower()
            else "даною військовою частиною")))
        )
    ),
    # (у / в / по) (цій / цій самій / цьому / цьому самому) (бригаді / полку / батальйоні ...)
    (
        re.compile(
            r"\b(у|в|по)\s+(цій|цій\s+самій|цьому|цьому\s+самому|зазначеній|зазначеному|вказаній|вказаному|даній|даному)\s+(бригаді|полку|батальйоні|дивізіоні|загоні|корпусі)\b",
            re.IGNORECASE | re.UNICODE,
        ),
        lambda m: (
            f"{m.group(1)} цій самій військовій частині" if "само" in m.group(0).lower()
            else (f"{m.group(1)} цій військовій частині" if m.group(2).lower() in ("цій", "цьому")
            else (f"{m.group(1)} зазначеній військовій частині" if "зазнач" in m.group(2).lower()
            else (f"{m.group(1)} вказаній військовій частині" if "вказан" in m.group(2).lower()
            else f"{m.group(1)} даній військовій частині")))
        )
    ),
]


def _format_full_closed_unit_text(mapped_val: dict | str, mapping_dict: dict) -> str:
    """Форматує закриту назву ВЧ: 'військової частини АXXXX', а якщо вказано корпус — 'військової частини АXXXX військової частини АYYYY'. Для ТЦК зберігає назву без змін."""
    if isinstance(mapped_val, dict):
        open_name = str(mapped_val.get("open_name", "")).strip()
        cipher = str(mapped_val.get("cipher") or mapped_val.get("closed_name") or "").strip()
        corps_name = str(mapped_val.get("corps", "")).strip()
    else:
        open_name = ""
        cipher = str(mapped_val).strip()
        corps_name = ""

    # Перевірка на ТЦК (територіальний центр комплектування) — залишаємо як є
    is_tck = (
        "тцк" in open_name.lower()
        or "комплектування" in open_name.lower()
        or "тцк" in cipher.lower()
        or "комплектування" in cipher.lower()
    )
    if is_tck:
        return cipher or open_name

    def _to_unit_phrase(raw: str) -> str:
        clean = raw.strip()
        if not clean:
            return ""
        if clean.lower().startswith("військової частини") or clean.lower().startswith("військова частина"):
            return clean
        if clean.lower().startswith("в/ч"):
            code = clean[3:].strip()
            return f"військової частини {code}"
        if clean.upper().startswith("А") or clean.upper().startswith("A") or clean.isdigit():
            return f"військової частини {clean}"
        return f"військової частини {clean}"

    unit_code = _to_unit_phrase(cipher)

    if corps_name:
        corps_entry = mapping_dict.get(corps_name)
        if corps_entry:
            corps_cipher = str(corps_entry.get("cipher") if isinstance(corps_entry, dict) else corps_entry)
            corps_code = _to_unit_phrase(corps_cipher)
        else:
            corps_code = _to_unit_phrase(corps_name)
        return f"{unit_code} {corps_code}"

    return unit_code


def _match_case(original_matched_text: str, replacement_text: str) -> str:
    """Якщо оригінальний співпавший текст введений великими літерами (ALL CAPS), повертає заміну у ВЕЛИКИХ ЛІТЕРАХ."""
    if not original_matched_text or not replacement_text:
        return replacement_text
    letters = [ch for ch in original_matched_text if ch.isalpha()]
    if letters and all(ch.isupper() for ch in letters):
        return replacement_text.upper()
    return replacement_text


def _apply_custom_rules(text: str, rules: dict | list | str | None) -> tuple[str, int]:
    """Застосовує додаткові правила/виправлення, передані через нижній порт 'rules' / 'corrections'."""
    if not text or not rules:
        return text, 0

    count = 0
    rule_dict = {}

    if isinstance(rules, dict):
        rule_dict = dict(rules)
    elif isinstance(rules, list):
        for item in rules:
            if isinstance(item, dict):
                rule_dict.update(item)
            elif isinstance(item, str) and "->" in item:
                parts = item.split("->", 1)
                rule_dict[parts[0].strip()] = parts[1].strip()
    elif isinstance(rules, str):
        for line in rules.splitlines():
            if "->" in line:
                parts = line.split("->", 1)
                rule_dict[parts[0].strip()] = parts[1].strip()

    for old_val, new_val in rule_dict.items():
        if not old_val or not isinstance(new_val, str):
            continue
        pattern = re.compile(re.escape(str(old_val)), re.IGNORECASE)
        if pattern.search(text):
            matches = pattern.findall(text)
            count += len(matches)
            text = pattern.sub(lambda m: _match_case(m.group(0), new_val), text)

    return text, count


@node(
    name="Правила та виправлення (підключення знизу)",
    category="Наказ",
    description=(
        "Створює пакет додаткових правил та виправлень для підключення до нижнього порту 'rules' / 'corrections' "
        "нод опрацювання наказів. Підтримує виправлення у форматі 'старе -> нове'."
    ),
    type_id="builtin.order.create_rules",
    outputs={
        "rules": "dict",
        "count": "int",
        "summary": "str",
    },
)
def create_order_rules(
    text_rules: str = "",
    overrides: dict | None = None,
) -> dict:
    rule_dict = {}
    if overrides:
        rule_dict.update(overrides)
    if text_rules.strip():
        for line in text_rules.splitlines():
            if "->" in line:
                parts = line.split("->", 1)
                rule_dict[parts[0].strip()] = parts[1].strip()
    summary = f"Сформовано правил/виправлень: {len(rule_dict)}"
    return {
        "rules": rule_dict,
        "count": len(rule_dict),
        "summary": summary,
    }


@node(
    name="Генерація наказу про прийняття рішень (закритий)",
    category="Наказ",
    description=(
        "Генерує закритий наказ про прийняття рішень: видаляє/замінює відкриту шапку на закриту, "
        "конвертує всі відкриті назви частин у форматовані шифри ('військової частини АXXXX' чи 'військової частини АXXXX військової частини АYYYY' для корпусів), "
        "зберігає CAPS якщо оригінальний текст введений ВЕЛИКИМИ ЛІТЕРАМИ, замінює звороти ('цієї самої бригади' -> 'цієї самої військової частини'), "
        "застосовує виправлення з нижнього порту 'rules', та повертає таблицю виявлених частин і корпусів."
    ),
    type_id="builtin.order.generate_decision_order",
    execution_inputs=("exec",),
    execution_outputs=("then",),
    outputs={
        "decision_text": "str",
        "table": "DataTable",
        "replaced_count": "int",
        "summary": "str",
    },
)
def generate_decision_order(
    text: str = "",
    mapping: dict | None = None,
    new_header: str = "НАКАЗ командира військової частини А0000 (по стройовій частині)",
    fuzzy_match: bool = True,
    rules: dict | list | str | None = None,
) -> dict:
    """Генерує закритий наказ про прийняття рішень."""
    if not text.strip():
        return {
            "decision_text": "",
            "table": DataTable(("Відкрита назва", "Закритий код ВЧ", "Армійський корпус", "Форматована заміна"), ()),
            "replaced_count": 0,
            "summary": "Порожній текст наказу",
        }

    text = _fix_military_typos(text)
    mapping_raw = mapping or {}
    mapping_dict = {}
    for k, v in mapping_raw.items():
        clean_k = _fix_military_typos(str(k))
        if isinstance(v, dict):
            clean_v = dict(v)
            if "open_name" in clean_v:
                clean_v["open_name"] = _fix_military_typos(str(clean_v["open_name"]))
            mapping_dict[clean_k] = clean_v
        else:
            mapping_dict[clean_k] = v

    lines = [line.rstrip() for line in text.splitlines()]

    # 1. Знаходимо початок змістовної частини (після шапки наказу)
    content_start_idx = 0
    for idx, line in enumerate(lines):
        clean = line.strip()
        if (
            clean.startswith("§")
            or re.match(r"^\d+[\.\)]", clean)
            or "НАКАЗУЮ" in clean.upper()
            or "ПРИЗНАЧИТИ" in clean.upper()
            or "НАПРАВИТИ" in clean.upper()
            or "ВІДРЯДИТИ" in clean.upper()
            or "ЗВІЛЬНИТИ" in clean.upper()
            or "ВІЙСЬКОВОСЛУЖБОВЦІВ" in clean.upper()
        ):
            content_start_idx = idx
            break

    body_lines = lines[content_start_idx:]
    body_text = "\n".join(body_lines)

    replaced_count = 0
    report_rows = []

    # 2. Замінюємо відкриті назви частин на закриті формовані назви (із корпусом) та збереженням CAPS
    for open_name, mapped_val in mapping_dict.items():
        closed_code = _format_full_closed_unit_text(mapped_val, mapping_dict)
        if isinstance(mapped_val, dict):
            raw_cipher = str(mapped_val.get("cipher", ""))
            corps_info = str(mapped_val.get("corps", ""))
            abbreviation = str(mapped_val.get("abbreviation", "")).strip()
        else:
            raw_cipher = str(mapped_val)
            corps_info = ""
            abbreviation = ""

        matched = False

        # Зіставляємо за сигнатурою відкриту назву
        pattern = _build_unit_fuzzy_pattern(open_name) if fuzzy_match else re.compile(re.escape(open_name), re.IGNORECASE)
        matches = pattern.findall(body_text)
        if matches:
            replaced_count += len(matches)
            body_text = pattern.sub(lambda m: _match_case(m.group(0), closed_code), body_text)
            matched = True

        # Зіставляємо скорочення
        if abbreviation and abbreviation != open_name:
            abbr_pattern = _build_unit_fuzzy_pattern(abbreviation) if fuzzy_match else re.compile(re.escape(abbreviation), re.IGNORECASE)
            abbr_matches = abbr_pattern.findall(body_text)
            if abbr_matches:
                replaced_count += len(abbr_matches)
                body_text = abbr_pattern.sub(lambda m: _match_case(m.group(0), closed_code), body_text)
                matched = True

        if matched:
            report_rows.append((open_name, raw_cipher, corps_info or "—", closed_code))

    # 3. Замінюємо звороти "цієї самої бригади", "цього самого полку" тощо (із збереженням CAPS)
    for pattern, replacer in _UNIT_PHRASE_REPLACEMENTS:
        matches = pattern.findall(body_text)
        if matches:
            replaced_count += len(matches)
            def _make_phrase_rep(r_func):
                return lambda m: _match_case(m.group(0), r_func(m))
            body_text = pattern.sub(_make_phrase_rep(replacer), body_text)

    # 4. Застосовуємо додаткові правила/виправлення з нижнього порту 'rules'
    if rules:
        body_text, custom_count = _apply_custom_rules(body_text, rules)
        replaced_count += custom_count

    # 5. Збираємо фінальний закритий наказ (нова шапка + закрите тіло)
    header_str = new_header.strip() if new_header.strip() else ""
    final_text = f"{header_str}\n\n{body_text}".strip() if header_str else body_text.strip()
    table = DataTable(
        ("Відкрита назва", "Закритий код ВЧ", "Армійський корпус", "Форматована заміна у тексті"),
        tuple(report_rows),
        "Звіт частин та корпусів наказу",
    )
    summary = f"Сформовано закритий наказ. Знайдено ВЧ/корпусів: {len(report_rows)}, замін: {replaced_count}"

    return {
        "decision_text": final_text,
        "table": table,
        "replaced_count": replaced_count,
        "summary": summary,
    }


# ── БЛОЧНИЙ КОНСТРУКТОР НАКАЗУ ───────────────────────────────────────────────

@node(
    name="Розібрати наказ на блоки (Конструктор)",
    category="Наказ",
    description=(
        "Розбирає текст наказу на структуровані блоки-конструктори: шапка (header), "
        "розділи (§ / section), пронумеровані пункти (item), підстави (basis) та підписи (footer)."
    ),
    type_id="builtin.order.parse_to_blocks",
    execution_inputs=("exec",),
    execution_outputs=("then",),
    outputs={
        "blocks": "List",
        "table": "DataTable",
        "count": "int",
        "summary": "str",
    },
)
def parse_to_blocks(text: str = "") -> dict:
    if not text.strip():
        return {
            "blocks": [],
            "table": DataTable(("ID", "Тип", "Мітка", "Вміст"), ()),
            "count": 0,
            "summary": "Порожній текст наказу",
        }

    lines = [line.rstrip() for line in text.splitlines()]
    blocks = []

    order_action_keywords = (
        "НАКАЗУЮ", "ПРИЗНАЧИТИ", "ЗВІЛЬНИТИ", "УКЛАСТИ", "ПРОДОВЖИТИ",
        "ПРИСВОЇТИ", "ОГОЛОСИТИ", "ПРИЙНЯТИ", "ЗАРАХУВАТИ", "ВВАЖАТИ",
        "НАПРАВИТИ", "ВІДРЯДИТИ", "ВИКЛЮЧИТИ", "ІМЕНУВАТИ", "СКАСУВАТИ", "ПОНОВИТИ"
    )

    # 1. Знаходимо шапку наказу (до першого § або пункту або дієслова-команди)
    content_start_idx = len(lines)
    header_lines = []
    for idx, line in enumerate(lines):
        clean = line.strip()
        if (
            clean.startswith("§")
            or re.match(r"^\d+[\.\)]", clean)
            or any(kw in clean.upper() for kw in order_action_keywords)
        ):
            content_start_idx = idx
            break
        if clean:
            header_lines.append(line)

    if header_lines:
        blocks.append({
            "id": "block_0",
            "type": "header",
            "label": "Шапка наказу",
            "text": "\n".join(header_lines),
            "lines": header_lines,
        })

    # 2. Розбираємо тіло наказу
    current_parent_heading = ""
    current_block = None
    block_counter = len(blocks)

    for line in lines[content_start_idx:]:
        clean = line.strip()
        if not clean:
            if current_block:
                current_block["lines"].append("")
            continue

        is_section = clean.startswith("§") or (
            ("Відповідно до" in clean or "Згідно з" in clean or clean.endswith(":"))
            and not re.match(r"^\d+[\.\d]*", clean)
            and not clean.startswith("Підстава")
        )
        is_basis = clean.startswith("Підстава") or clean.startswith("Підстави")
        is_item = (
            not is_section
            and not is_basis
            and not ("р.н." in clean or "р. н." in clean)
            and bool(re.match(r"^\d{1,3}(?:\.\d{1,3})*[\.\)]\s+", clean))
        )
        is_footer = clean.startswith("Командир") or clean.startswith("ТВО") or clean.startswith("ТИМЧАСОВО") or clean.startswith("Згідно з оригіналом")

        if is_section:
            if clean.startswith("§") or not current_parent_heading:
                current_parent_heading = line
            else:
                current_parent_heading += "\n\n" + line
            current_block = {
                "id": f"block_{block_counter}",
                "type": "section",
                "label": clean.split()[0] if clean.split() else "Розділ",
                "heading": current_parent_heading,
                "text": line,
                "lines": [line],
            }
            block_counter += 1
            blocks.append(current_block)
        elif is_item:
            match = re.match(r"^(\d+[\.\d]*)", clean)
            label = f"Пункт {match.group(1)}" if match else "Пункт"
            current_block = {
                "id": f"block_{block_counter}",
                "type": "item",
                "label": label,
                "heading": current_parent_heading,
                "text": line,
                "lines": [line],
            }
            block_counter += 1
            blocks.append(current_block)
        elif is_basis:
            current_block = {
                "id": f"block_{block_counter}",
                "type": "basis",
                "label": "Підстава",
                "heading": current_parent_heading,
                "text": line,
                "lines": [line],
            }
            block_counter += 1
            blocks.append(current_block)
        elif is_footer:
            current_block = {
                "id": f"block_{block_counter}",
                "type": "footer",
                "label": "Підпис / Фінал",
                "heading": "",
                "text": line,
                "lines": [line],
            }
            block_counter += 1
            blocks.append(current_block)
        else:
            if current_block:
                current_block["lines"].append(line)
                current_block["text"] = "\n".join(current_block["lines"])
            else:
                current_block = {
                    "id": f"block_{block_counter}",
                    "type": "item",
                    "label": "Основний текст",
                    "heading": current_parent_heading,
                    "text": line,
                    "lines": [line],
                }
                block_counter += 1
                blocks.append(current_block)

    table_rows = [
        (b["id"], b["type"], b["label"], b["text"][:80] + ("..." if len(b["text"]) > 80 else ""))
        for b in blocks
    ]
    table = DataTable(("ID", "Тип", "Мітка", "Вміст (прев'ю)"), tuple(table_rows), "Блоки наказу")

    return {
        "blocks": blocks,
        "table": table,
        "count": len(blocks),
        "summary": f"Розділено наказ на {len(blocks)} блоків-конструкторів",
    }


@node(
    name="Трансформація та фільтрація блоків",
    category="Наказ",
    description=(
        "Фільтрує або змінює окремі блоки наказу: застосовує карту замін (відкриті -> 'військової частини АXXXX' / 'військової частини АXXXX військової частини АYYYY'), "
        "замінює звороти 'цієї самої бригади' на 'цієї самої військової частини', та фільтрує за типом блоків."
    ),
    type_id="builtin.order.filter_transform_blocks",
    execution_inputs=("exec",),
    execution_outputs=("then",),
    outputs={
        "blocks": "List",
        "table": "DataTable",
        "modified_count": "int",
        "summary": "str",
    },
)
def filter_transform_blocks(
    blocks: list | None = None,
    include_types: str = "",
    mapping: dict | None = None,
    replace_unit_phrases: bool = True,
) -> dict:
    if not blocks:
        return {"blocks": [], "table": DataTable(("ID", "Тип", "Мітка", "Вміст"), ()), "modified_count": 0, "summary": "Порожній список блоків"}

    mapping_dict = mapping or {}
    types_set = {t.strip().casefold() for t in include_types.split(",") if t.strip()} if include_types.strip() else None

    result_blocks = []
    modified_count = 0

    for block in blocks:
        if not isinstance(block, dict):
            continue
        b_type = str(block.get("type", "")).casefold()
        if types_set and b_type not in types_set:
            continue

        b_copy = dict(block)
        lines = list(b_copy.get("lines", []))
        text = b_copy.get("text", "\n".join(lines))
        orig_text = text

        # 1. Заміна назв ВЧ та корпусів на форматовані шифри
        for open_name, mapped_val in mapping_dict.items():
            closed_code = _format_full_closed_unit_text(mapped_val, mapping_dict)
            pattern = _build_unit_fuzzy_pattern(open_name)
            if pattern.search(text):
                text = pattern.sub(closed_code, text)

        # 2. Заміна зворотів ("цієї самої бригади" -> "цієї самої військової частини")
        if replace_unit_phrases:
            for pattern, replacer in _UNIT_PHRASE_REPLACEMENTS:
                if pattern.search(text):
                    text = pattern.sub(replacer, text)

        if text != orig_text:
            modified_count += 1
            b_copy["text"] = text
            b_copy["lines"] = text.splitlines()

        result_blocks.append(b_copy)

    table_rows = [
        (b.get("id", ""), b.get("type", ""), b.get("label", ""), b.get("text", "")[:80] + ("..." if len(b.get("text", "")) > 80 else ""))
        for b in result_blocks
    ]
    table = DataTable(("ID", "Тип", "Мітка", "Вміст (прев'ю)"), tuple(table_rows), "Трансформовані блоки")

    summary = f"Опрацьовано блоків: {len(result_blocks)} (змінено: {modified_count})"
    return {"blocks": result_blocks, "table": table, "modified_count": modified_count, "summary": summary}


@node(
    name="Зібрати наказ з блоків",
    category="Наказ",
    description=(
        "Збирає структурований список блоків-конструкторів у єдиний готовий текст наказу "
        "із можливістю заміни шапки чи розділювачів."
    ),
    type_id="builtin.order.assemble_from_blocks",
    execution_inputs=("exec",),
    execution_outputs=("then",),
    outputs={
        "text": "str",
        "count": "int",
        "summary": "str",
    },
)
def assemble_from_blocks(
    blocks: list | None = None,
    new_header: str = "",
    separator: str = "\n\n",
) -> dict:
    if not blocks:
        return {"text": "", "count": 0, "summary": "Порожній список блоків"}

    block_texts = []
    header_inserted = False

    for block in blocks:
        if not isinstance(block, dict):
            continue
        b_type = str(block.get("type", "")).casefold()
        b_text = str(block.get("text", "")).strip()

        if b_type == "header":
            if new_header.strip():
                block_texts.append(new_header.strip())
                header_inserted = True
            elif b_text:
                block_texts.append(b_text)
                header_inserted = True
        elif b_text:
            block_texts.append(b_text)

    # Якщо була вказана нова шапка, але блоку 'header' не було у списку — додаємо її на початок
    if new_header.strip() and not header_inserted:
        block_texts.insert(0, new_header.strip())

    sep = separator if separator else "\n\n"
    final_text = sep.join(block_texts).strip()

    summary = f"Зібрано наказ з {len(block_texts)} блоків"
    return {"text": final_text, "count": len(block_texts), "summary": summary}


@node(
    name="Розрахунок витягів з наказів",
    category="Наказ",
    description=(
        "Автоматично розраховує перелік витягів з наказу для кожної військової частини (ВЧ). "
        "Приймає наказ та таблицю частин, формує закриті коди частин (А1500, А1400) та виводить готовий список пунктів із дефісами (напр. А1500 1,3  А1400 1-6)."
    ),
    type_id="builtin.order.calculate_extracts",
    execution_inputs=("exec",),
    execution_outputs=("then",),
    outputs={
        "summary_text": "str",
        "table": "DataTable",
        "extracts": "dict",
        "count": "int",
        "summary": "str",
    },
)
def calculate_order_extracts(
    text: str = "",
    mapping: dict | None = None,
    header: str = "",
    rules: dict | list | str | None = None,
) -> dict:
    """Розпізнає частини у наказі та формує точний розрахунок витягів у форматі А1500 1,3 А1400 1-6."""
    if not text.strip():
        return {
            "summary_text": "",
            "table": DataTable(("Закрита ВЧ", "Відкрита назва", "Номери пунктів витягу", "Кількість пунктів"), ()),
            "extracts": {},
            "count": 0,
            "summary": "Порожній текст наказу",
        }

    body_text = text
    if rules:
        body_text, _ = _apply_custom_rules(body_text, rules)

    res = map_military_units(text=body_text, mapping=mapping)
    units_map = res.get("unit_paragraphs", {})
    abbr_map = res.get("unit_abbr_map", {})
    extracts_dict: dict[str, str] = {}
    table_rows = []
    summary_lines = []

    for sender, data in units_map.items():
        header_text = header
        if not header_text and isinstance(data, dict):
            header_text = "\n".join(data.get("header_lines", []))

        short_sender = sender
        open_names = ", ".join(sorted(res.get("unit_open_names", {}).get(sender, []))) if "unit_open_names" in res else ""

        items_text = ""
        range_labels = "-"
        count = 0

        if isinstance(data, dict) and "items" in data:
            section_groups: dict[str, list[str]] = {}
            raw_labels = []
            for item in data.get("items", []):
                heading = item.get("parent_heading", "").strip()
                item_text = item.get("text", "").strip()
                label = item.get("label", "")
                if label:
                    raw_labels.append(label)
                if item_text:
                    section_groups.setdefault(heading, []).append(item_text)

            range_labels = _format_item_numbers_range(raw_labels)
            count = len(data.get("items", []))

            body_sections = []
            for heading, item_texts in section_groups.items():
                if heading:
                    body_sections.append(f"{heading}\n\n" + "\n\n".join(item_texts))
                else:
                    body_sections.append("\n\n".join(item_texts))

            items_text = "\n\n".join(body_sections)
        elif isinstance(data, str):
            items_text = data
            count = 1

        full_doc = f"{header_text}\n\n{items_text}".strip() if header_text else items_text.strip()
        extracts_dict[short_sender] = full_doc
        table_rows.append((short_sender, open_names or "—", range_labels, count))
        summary_lines.append(f"{short_sender} {range_labels}")

    table = DataTable(
        ("Закрита ВЧ (Адресат)", "Відкрита назва ВЧ", "Номери пунктів витягу", "Кількість пунктів"),
        tuple(table_rows),
        "Розрахунок витягів з наказів",
    )
    summary_text = "\n".join(summary_lines)
    summary = f"Розраховано витягів: {len(extracts_dict)}"

    return {
        "summary_text": summary_text,
        "table": table,
        "extracts": extracts_dict,
        "count": len(extracts_dict),
        "summary": summary,
    }
