"""Збирає довідники з наказу МО України від 01.08.2023 № 444 (відкриті дані).

Джерело: «Текст для друку» картки z1562-23 на zakon.rada.gov.ua
(`/laws/show/z1562-23/print`), збережений як `.html`.

Результат — CSV у `personnel/dictionaries/` (UTF-8 з BOM, роздільник `;`):
- `mo444_officer_vos.csv` — Перелік ВОС осіб офіцерського складу: група, код ВОС
  (6 цифр), найменування, посади, на яких застосовується, допустиме призначення
  офіцерів без додаткової підготовки та з нею;
- `mo444_vos_replacement.csv` — додаток 2: код ВОС, що підлягає заміні → новий код.

На відміну від наказу № 317, тут ВОС прив'язані не до назв посад, а до категорій
посад («на командних і штабних посадах…»). Перелік ВОС, за якими присвоюють
первинне звання молодшого лейтенанта запасу, у тексті сторінки не опубліковано.

    python scripts/order_index/build_mo444_dictionaries.py ДРУК.html
"""

from __future__ import annotations

import argparse
import csv
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_mo317_dictionaries import OUT_DIR, _Rows  # noqa: E402

VOS_COLUMNS = [
    "Група", "Назва групи", "Код ВОС", "Найменування ВОС", "Посади, на яких застосовується",
    "Допустимо без додаткової підготовки", "Допустимо з додатковою підготовкою", "Примітка",
]
REPLACEMENT_COLUMNS = ["Код ВОС, що підлягає заміні", "Код ВОС, на який здійснюється заміна"]

_CODE = re.compile(r"\d{6}")


def parse(html_text: str) -> tuple[list[dict[str, str]], list[dict[str, str]]]:
    parser = _Rows()
    parser.feed(html_text)
    vos_rows: list[dict[str, str]] = []
    replacements: list[dict[str, str]] = []
    group = group_name = ""
    part = ""
    for cells in parser.rows:
        first = cells[0] if cells else ""
        if cells[:3] == ["ВОС", "Посади, на яких застосовується", "Допустиме призначення офіцерів"]:
            part = "vos"
            continue
        if len(cells) == 6 and cells[0] == "яка підлягає заміні":
            part = "replacement"
            continue
        if any("Додаток 1" in cell for cell in cells):
            part = ""
            continue
        if part == "vos":
            if len(cells) == 2 and first.startswith("Група"):
                group, group_name = first.replace("Група", "").strip(), cells[1]
                continue
            excluded = next((cell for cell in cells if "виключено" in cell), "")
            match = re.match(r"^\{?Рядок\s+(\d{6})", first)
            if excluded and (match or _CODE.fullmatch(first)):
                code = match.group(1) if match else first
                vos_rows.append(_vos(group, group_name, code, "", "", "", "", excluded.strip("{} ")))
                continue
            if _CODE.fullmatch(first) and len(cells) >= 2:
                padded = cells + [""] * (5 - len(cells))
                vos_rows.append(_vos(group, group_name, *padded[:5], ""))
        elif part == "replacement":
            if all(not cell or _CODE.fullmatch(cell) for cell in cells) and any(cells):
                for old, new in zip(cells[0::2], cells[1::2], strict=False):
                    if old and new:
                        replacements.append({REPLACEMENT_COLUMNS[0]: old, REPLACEMENT_COLUMNS[1]: new})
    # «Рядок 211000 виключено» поруч із новим рядком того самого коду — лишається чинний.
    active_codes = {row["Код ВОС"] for row in vos_rows if not row["Примітка"]}
    vos_rows = [row for row in vos_rows if not (row["Примітка"] and row["Код ВОС"] in active_codes)]
    return vos_rows, replacements


def _vos(group, group_name, code, name, positions, without, with_training, note) -> dict[str, str]:
    return dict(zip(VOS_COLUMNS, (group, group_name, code, name, positions, _codes(without),
                                  _codes(with_training), note), strict=True))


def _codes(text: str) -> str:
    """«012000 060100» → «012000, 060100»; словесний опис лишається як є."""
    codes = _CODE.findall(text)
    rest = _CODE.sub("", text).strip(" ,;")
    return ", ".join(codes + ([rest] if rest else []))


def _write(path: Path, columns: list[str], rows: list[dict[str, str]]) -> None:
    with open(path, "w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=columns, delimiter=";")
        writer.writeheader()
        writer.writerows(rows)
    print(f"{path.name}: {len(rows)} рядків")


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("print_html", type=Path)
    parser.add_argument("--out", type=Path, default=OUT_DIR)
    args = parser.parse_args(argv)
    vos_rows, replacements = parse(args.print_html.read_bytes().decode("utf-8"))
    if not vos_rows:
        print("❌ Порожній результат — перевірте файл джерела", file=sys.stderr)
        return 1
    _write(args.out / "mo444_officer_vos.csv", VOS_COLUMNS, vos_rows)
    _write(args.out / "mo444_vos_replacement.csv", REPLACEMENT_COLUMNS, replacements)
    active = [row for row in vos_rows if not row["Примітка"]]
    print(f"Груп: {len({row['Група'] for row in vos_rows})}, чинних ВОС: {len(active)}, "
          f"виключених: {len(vos_rows) - len(active)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
