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
    headers = [cell.strip() for cell in rows[0]]
    normalized = {name.casefold(): index for index, name in enumerate(headers)}
    requested = [open_name_column, cipher_column, destination_column]
    missing = [name for name in requested if name.casefold() not in normalized]
    if missing:
        raise ValueError("У таблиці немає колонок: " + ", ".join(missing))
    indexes = [normalized[name.casefold()] for name in requested]
    mapping: dict[str, dict[str, str]] = {}
    table_rows = []
    for row in rows[1:]:
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
