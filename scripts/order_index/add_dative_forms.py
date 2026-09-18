"""Дописує стовпець «Давальний» у довідники звань і посад.

Наказ про присвоєння військового звання пише особу в давальному:
«Капітану НАЗАРЕНКУ Олександру Васильовичу, старшому офіцеру відділу …»
(зразки додатка 53). У довідниках `personnel/dictionaries/ranks.csv` і
`positions.csv` були лише знахідний та орудний, тож давального взяти не було
звідки.

Звідки беруться форми:
1. посади — зі зведеного довідника `order_index/positions_all.csv`, де стовпець
   «Давальний» уже є (він зібраний з Excel-генератора й наказу МО № 317);
2. чого там немає — за правилом: прикметник на -ий/-ій → -ому/-ьому,
   іменник на приголосний → +у, на -ь/-й → +ю, на -а → -і, на -я → -ї.

Правило навмисне просте: рядок, який воно зіпсувало, користувач виправляє
прямо в CSV, а генератор попереджає, коли посади немає в довіднику.

Запуск: python scripts/order_index/add_dative_forms.py
"""

from __future__ import annotations

import csv
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
DICTIONARIES = ROOT / "src" / "nodeautomationtoolkit" / "personnel" / "dictionaries"
ALL_POSITIONS = ROOT / "src" / "nodeautomationtoolkit" / "order_index" / "positions_all.csv"
DATIVE = "Давальний"

_ADJECTIVE = {"ий": "ому", "ій": "ьому", "а": "ій", "я": "ій"}


def dative_word(word: str, adjective: bool = False) -> str:
    """Давальний одного слова; `adjective` — коли слово стоїть перед іменником."""
    low = word.casefold()
    if adjective:
        for ending, form in _ADJECTIVE.items():
            if low.endswith(ending) and len(low) > len(ending) + 1:
                return low[: -len(ending)] + form
    if low.endswith(("ий", "ій")):
        return low[:-2] + ("ому" if low.endswith("ий") else "ьому")
    if low.endswith("ь"):
        return low[:-1] + "ю"
    if low.endswith("й"):
        return low[:-1] + "ю"
    if low.endswith("а"):
        return low[:-1] + "і"
    if low.endswith("я"):
        return low[:-1] + "ї"
    if low.endswith("о"):
        return low[:-1] + "у"
    if low[-1:].isalpha():
        return low + "у"
    return low


def dative_phrase(phrase: str) -> str:
    """Давальний словосполучення: змінюються всі слова до першого родового означення.

    «старший офіцер відділу» → «старшому офіцеру відділу»: перші два слова
    (прикметник + іменник) стають у давальний, решта лишається як є.
    """
    words = phrase.split()
    if not words:
        return ""
    result = []
    changed = 0
    for index, word in enumerate(words):
        if changed >= 2 or "-" in word and index:
            result.append(word)
            continue
        low = word.casefold()
        if index == 0 and low.endswith(("ий", "ій")) and len(words) > 1:
            result.append(dative_word(word, adjective=True))
            continue
        if changed == 0 or (changed == 1 and index == 1 and result and result[0] != words[0]):
            result.append(dative_word(word))
            changed = 2 if index else 1
            continue
        result.append(word)
    return " ".join(result)


def known_datives() -> dict[str, str]:
    """Називний → давальний зі зведеного довідника посад."""
    if not ALL_POSITIONS.is_file():
        return {}
    rows = list(csv.DictReader(ALL_POSITIONS.read_text(encoding="utf-8-sig").splitlines(), delimiter=";"))
    table = {}
    for row in rows:
        nominative = (row.get("Називний") or "").strip().casefold()
        dative = (row.get("Давальний") or "").strip()
        if nominative and dative:
            table.setdefault(nominative, dative)
    return table


def add_column(path: Path, ready: dict[str, str]) -> int:
    lines = path.read_text(encoding="utf-8-sig").splitlines()
    rows = list(csv.reader(lines, delimiter=";"))
    if not rows:
        return 0
    header = [cell.strip() for cell in rows[0]]
    if DATIVE in header:
        print(f"{path.name}: стовпець «{DATIVE}» уже є")
        return 0
    insert_at = len(header) - 1 if header and header[-1] == "Примітка" else len(header)
    header.insert(insert_at, DATIVE)

    filled = 0
    out = [header]
    for row in rows[1:]:
        if not row or not row[0].strip():
            out.append(row)
            continue
        row = [*row, *[""] * (len(header) - 1 - len(row))]
        nominative = row[0].strip()
        form = ready.get(nominative.casefold()) or dative_phrase(nominative)
        row.insert(insert_at, form)
        filled += 1
        out.append(row)

    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        csv.writer(handle, delimiter=";", lineterminator="\n").writerows(out)
    return filled


def main() -> int:
    ready = known_datives()
    print(f"Готових форм зі зведеного довідника: {len(ready)}")
    for name in ("ranks.csv", "positions.csv"):
        path = DICTIONARIES / name
        if not path.is_file():
            print(f"{name}: файла немає")
            continue
        print(f"{name}: дописано рядків {add_column(path, ready)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
