"""Збирає довідники з наказу МО України від 07.09.2020 № 317 (відкриті дані).

Джерела (завантажити вручну із zakon.rada.gov.ua, картка z0927-20):
- «Файли» → сигнальний документ `.docx` — таблиця «Перелік штатних посад рядового,
  сержантського і старшинського складу…» (код посади, найменування, штатне звання,
  код звання, ВОС, тарифний розряд);
- «Текст для друку» (`/laws/show/z0927-20/print`), збережений як `.html` —
  «Перелік військово-облікових спеціальностей» (номер ВОС, найменування, розділ).

Результат — два CSV у `personnel/dictionaries/` (UTF-8 з BOM, роздільник `;`):
- `mo317_positions.csv` — рядок таблиці на посаду/варіант (а, б, в …) з ВОС;
- `mo317_vos.csv` — номер ВОС → найменування, розділ і підрозділ.

    python scripts/order_index/build_mo317_dictionaries.py ПЕРЕЛІК.docx ДРУК.html [--edition "07.04.2026"]
"""

from __future__ import annotations

import argparse
import csv
import re
import sys
from html.parser import HTMLParser
from pathlib import Path

import docx

ROOT = Path(__file__).resolve().parents[2]
OUT_DIR = ROOT / "src" / "nodeautomationtoolkit" / "personnel" / "dictionaries"

POSITION_COLUMNS = [
    "№ з/п", "Код посади", "Найменування посади", "Штатне звання", "Код звання",
    "ВОС (як у переліку)", "Номери ВОС", "Коди ВОС-посада", "Усі ВОС", "Крім ВОС", "Тарифний розряд",
]
VOS_COLUMNS = ["Номер ВОС", "Найменування ВОС", "Розділ", "Підрозділ"]

_RANGE_RE = re.compile(r"(\d{6})\s*[-–—]\s*(\d{6})")
_CODE_RE = re.compile(r"(?<!\d)(\d{6})(?!\d)")


def _clean(text: str) -> str:
    return " ".join(text.replace("\xa0", " ").split())


def _clean_rank(text: str) -> str:
    """«Сержант/ старшина 1 статті» → «Сержант / старшина 1 статті» (армійське / корабельне)."""
    return re.sub(r"\s*/\s*", " / ", _clean(text)).strip(" /")


def _codes(text: str) -> list[str]:
    full_codes: set[str] = set()
    rest = text
    for start, end in _RANGE_RE.findall(text):
        tail = start[3:]
        for vos in range(int(start[:3]), int(end[:3]) + 1):
            full_codes.add(f"{vos:03d}{tail}")
        rest = rest.replace(start, " ").replace(end, " ")
    full_codes.update(_CODE_RE.findall(rest))
    return sorted(full_codes)


def expand_vos(text: str) -> tuple[list[str], list[str], bool, list[str]]:
    """(номери ВОС, коди ВОС-посада, «усі ВОС», виключені номери ВОС) з тексту графи 6.

    «усі ВОС, крім ВОС: 113258, …» — коди після «крім» не дозволені, а виключені.
    """
    all_vos = bool(re.search(r"усі\s+ВОС", text, re.IGNORECASE))
    allowed, excluded = text, ""
    if all_vos:
        parts = re.split(r"\bкрім\b", text, maxsplit=1, flags=re.IGNORECASE)
        if len(parts) == 2:
            allowed, excluded = parts
    codes = _codes(allowed)
    excluded_vos = sorted({code[:3] for code in _codes(excluded)})
    return sorted({code[:3] for code in codes}), codes, all_vos, excluded_vos


def read_positions(path: Path) -> list[dict[str, str]]:
    document = docx.Document(str(path))
    table = max(document.tables, key=lambda t: len(t.rows))
    result: list[dict[str, str]] = []
    previous: dict[str, str] | None = None
    for row in table.rows:
        cells = [cell.text for cell in row.cells]
        if len(cells) < 7 or not re.fullmatch(r"\d{3}", cells[0].strip()):
            continue
        number, code, name, rank, rank_code, vos_text, tariff = (_clean(c) for c in cells[:7])
        rank, rank_code = _clean_rank(cells[3]), _clean_rank(cells[4])
        if not name:
            if previous and previous["№ з/п"] == number:
                # Продовження тієї самої посади: варіант «б)» / «в)» з іншою ВОС чи розрядом.
                name, code, rank, rank_code = (
                    previous["Найменування посади"], previous["Код посади"],
                    previous["Штатне звання"], previous["Код звання"],
                )
            else:
                continue  # код посади виключено з переліку
        vos, codes, all_vos, excluded = expand_vos(vos_text)
        entry = {
            "№ з/п": number,
            "Код посади": code,
            "Найменування посади": name,
            "Штатне звання": rank,
            "Код звання": rank_code,
            "ВОС (як у переліку)": vos_text,
            "Номери ВОС": ", ".join(vos),
            "Коди ВОС-посада": ", ".join(codes),
            "Усі ВОС": "так" if all_vos else "",
            "Крім ВОС": ", ".join(excluded),
            "Тарифний розряд": tariff,
        }
        result.append(entry)
        previous = entry
    return result


class _Rows(HTMLParser):
    def __init__(self):
        super().__init__()
        self.rows: list[list[str]] = []
        self._row: list[str] | None = None
        self._cell: list[str] | None = None

    def handle_starttag(self, tag, attrs):
        if tag == "tr":
            self._row = []
        elif tag in ("td", "th") and self._row is not None:
            self._cell = []
        elif tag == "br" and self._cell is not None:
            self._cell.append(" ")

    def handle_endtag(self, tag):
        if tag in ("td", "th") and self._cell is not None and self._row is not None:
            self._row.append(" ".join("".join(self._cell).split()))
            self._cell = None
        elif tag == "tr" and self._row is not None:
            self.rows.append(self._row)
            self._row = None

    def handle_data(self, data):
        if self._cell is not None:
            self._cell.append(data)


def read_vos(path: Path) -> list[dict[str, str]]:
    parser = _Rows()
    parser.feed(path.read_bytes().decode("utf-8"))
    result: list[dict[str, str]] = []
    section = subsection = ""
    in_list = False
    for row in parser.rows:
        cells = [cell for cell in row]
        if cells[:1] == ["Номер ВОС"]:
            in_list = True
            continue
        if not in_list:
            continue
        if len(cells) == 1:
            title = cells[0]
            if re.match(r"^\d+\.\d+\.", title):
                subsection = title
            elif re.match(r"^\d+\.\s", title):
                section, subsection = title, ""
            continue
        if len(cells) != 2:
            if result:
                break  # далі — тарифний перелік музикантів
            continue
        number, name = cells
        if re.fullmatch(r"\d{3}", number) and name:
            result.append({"Номер ВОС": number, "Найменування ВОС": name, "Розділ": section, "Підрозділ": subsection})
    return result


def _write(path: Path, columns: list[str], rows: list[dict[str, str]], source: str) -> None:
    with open(path, "w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=columns, delimiter=";")
        writer.writeheader()
        writer.writerows(rows)
    print(f"{path.name}: {len(rows)} рядків ← {source}")


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("positions_docx", type=Path)
    parser.add_argument("print_html", type=Path)
    parser.add_argument("--out", type=Path, default=OUT_DIR)
    args = parser.parse_args(argv)
    positions = read_positions(args.positions_docx)
    vos = read_vos(args.print_html)
    if not positions or not vos:
        print("❌ Порожній результат — перевірте файли джерел", file=sys.stderr)
        return 1
    _write(args.out / "mo317_positions.csv", POSITION_COLUMNS, positions, args.positions_docx.name)
    _write(args.out / "mo317_vos.csv", VOS_COLUMNS, vos, args.print_html.name)
    known_vos = {row["Номер ВОС"] for row in vos}
    missing = sorted({v for row in positions for v in row["Номери ВОС"].split(", ") if v} - known_vos)
    print(f"Посад: {len({row['Найменування посади'].casefold() for row in positions})} різних, "
          f"кодів посад: {len({row['Код посади'] for row in positions})}; "
          f"ВОС у посадах без назви в переліку ВОС: {len(missing)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
