"""Збирає довідник відмінків посад і звань для генератора наказів.

Форми беруться з трьох джерел (у порядку довіри):

1. **Excel-генератор «Нормалізатор на війська»** — аркуші «посада кандидата»
   (називний → знахідний), «вакантна посада» (називний → орудний) і «звання
   кандидата» (називний → знахідний). Це форми, вивірені людьми на реальних
   наказах, тому вони перемагають.
2. **`order_index/positions_all.csv`** — зведений довідник (родовий, давальний,
   орудний), зібраний із того ж Excel і наказу МО № 317.
3. **Те, що вже лежить у `personnel/dictionaries/`** — зокрема рядки, які
   користувач виправив руками: вони не затираються.

Знахідний для назв посад збігається з родовим (істоти чоловічого роду:
«командир» → «командира»), тому порожній знахідний заповнюється родовим.

Файл Excel у git не кладеться (там реальні дані) — шлях передається аргументом
або береться з робочого столу.

Запуск:
    python scripts/order_index/build_position_forms.py [шлях до .xlsm]
"""

from __future__ import annotations

import csv
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
DICTIONARIES = ROOT / "src" / "nodeautomationtoolkit" / "personnel" / "dictionaries"
ALL_POSITIONS = ROOT / "src" / "nodeautomationtoolkit" / "order_index" / "positions_all.csv"
DEFAULT_EXCEL = Path.home() / "Desktop" / "Нормалізатор на війська.xlsm"

POSITION_COLUMNS = ["Називний", "Родовий", "Знахідний", "Давальний", "Орудний", "Джерело", "Примітка"]
RANK_COLUMNS = ["Називний", "Родовий", "Знахідний", "Давальний", "Джерело", "Примітка"]
MANUAL = "вручну"


#: Латинські літери, схожі на кириличні: в Excel трапляється «гpанатометник»
#: із латинською «p». Такі рядки ставали окремими посадами без відмінків.
_LATIN_LOOKALIKES = str.maketrans("ABCEHIKMOPTXacehiopxy", "АВСЕНІКМОРТХасеніорху")


def _clean(value) -> str:
    return " ".join(str(value or "").translate(_LATIN_LOOKALIKES).split())


def instrumental_form(nominative: str) -> str:
    """Орудний за правилом — коли готової форми немає ніде.

    «гранатометник» → «гранатометником», «стрілець» → «стрільцем»,
    «медична сестра» → «медичною сестрою».
    """
    words = _clean(nominative).split()
    if not words:
        return ""
    result = []
    for index, word in enumerate(words):
        low = word.casefold()
        following = words[index + 1].casefold() if index + 1 < len(words) else ""
        adjective = low.endswith(("ий", "ій")) or (
            low.endswith(("а", "я")) and following.endswith(("а", "я"))
        )
        if adjective:
            if low.endswith("ий"):
                result.append(word[:-2] + "им")
            elif low.endswith("ій"):
                result.append(word[:-2] + "ім")
            else:
                result.append(word[:-1] + "ою")
            continue
        if low.endswith("ець"):
            result.append(word[:-3] + "цем")
        elif low.endswith(("а",)):
            result.append(word[:-1] + "ою")
        elif low.endswith(("я",)):
            result.append(word[:-1] + "ею")
        elif low.endswith(("ь", "й")):
            result.append(word[:-1] + "ем")
        elif low.endswith("о"):
            result.append(word[:-1] + "ом")
        elif low[-1:].isalpha():
            result.append(word + "ом")
        else:
            result.append(word)
        result.extend(words[index + 1 :])
        break
    return " ".join(result)


def _leading_first_declension(nominative: str) -> int:
    """Скільки слів на початку назви закінчуються на -а/-я.

    «старшина батальйону» → 1, «старша медична сестра» → 3, «командир
    відділення» → 0. Саме ці слова мають окремий знахідний («старшину»,
    «медичну сестру»); решта назв має знахідний, однаковий із родовим
    («командира відділення»). Останнє слово перевіряти не можна: «відділення»
    там — залежне слово в родовому, а не головне.
    """
    count = 0
    for word in _clean(nominative).casefold().split():
        if word.endswith(("а", "я")):
            count += 1
        else:
            break
    return count


def _same_shape(nominative: str, form: str) -> bool:
    """Чи схожа форма на відмінок ЦІЄЇ назви: відмінок не змінює кількості слів.

    В Excel трапляються рядки, де навпроти короткої назви стоїть форма довшої
    посади («командир відділення» → «командиром відділення - інструктором з
    фізичної підготовки»). Такі форми до довідника не беремо.
    """
    left, right = _clean(nominative).split(), _clean(form).split()
    return bool(right) and len(left) == len(right)


def accusative_form(nominative: str, genitive: str) -> str:
    """Знахідний назви посади.

    Для більшості назв він збігається з родовим («командира роти»). Для назв
    на -а/-я виводиться за правилом: кожне слово до останнього іменника
    змінює -а на -у, -я на -ю («старша медична сестра» → «старшу медичну
    сестру»).
    """
    nominative = _clean(nominative)
    if not nominative:
        return ""
    leading = _leading_first_declension(nominative)
    if not leading:
        return _clean(genitive)
    words = []
    for index, word in enumerate(nominative.split()):
        low = word.casefold()
        if index >= leading:
            words.append(word)
        elif low.endswith("а"):
            words.append(word[:-1] + "у")
        elif low.endswith("я"):
            words.append(word[:-1] + "ю")
        else:
            words.append(word)
    return " ".join(words)


#: Закінчення прикметника в називному → давальний.
_ADJECTIVE_DATIVE = {"ий": "ому", "ій": "ьому", "а": "ій", "я": "ій"}


def dative_form(nominative: str) -> str:
    """Давальний назви посади за правилом — коли готової форми немає ніде.

    «старша медична сестра» → «старшій медичній сестрі», «командир роти» →
    «командиру роти»: змінюються слова на початку (прикметники й головний
    іменник), залежні слова в родовому лишаються.
    """
    words = _clean(nominative).split()
    if not words:
        return ""
    result = []
    for index, word in enumerate(words):
        low = word.casefold()
        following = words[index + 1].casefold() if index + 1 < len(words) else ""
        adjective = low.endswith(("ий", "ій")) or (
            low.endswith(("а", "я")) and following.endswith(("а", "я"))
        )
        if index > 0 and not result[-1].endswith(("ому", "ьому", "ій", "у", "ю", "і")):
            result.append(word)
            continue
        if adjective:
            for ending, form in _ADJECTIVE_DATIVE.items():
                if low.endswith(ending) and len(low) > len(ending) + 1:
                    result.append(word[: -len(ending)] + form)
                    break
            else:
                result.append(word)
        elif low.endswith(("а", "я")):
            result.append(word[:-1] + "і")
        elif low.endswith(("ь", "й")):
            result.append(word[:-1] + "ю")
        elif low.endswith("о"):
            result.append(word[:-1] + "у")
        elif low[-1:].isalpha():
            result.append(word + "у")
        else:
            result.append(word)
        if not adjective:
            result.extend(words[index + 1 :])
            break
    return " ".join(result)


#: Перевернуті рядки з каталогу МО 317: назва посади так не пишеться, а модуль
#: відмінювання вчить із них хибні форми слів («старша» → «старші»). Список
#: явний, бо правилом їх від законних назв на кшталт «асистент фармацевта» чи
#: «командир відділення» не відрізнити: там останнє слово теж на -а/-я.
CATALOGUE_ARTIFACTS = {
    "сестра медична",
    "сестра медична операційна",
    "старша медична",
    "головна медична",
    "молодша медична",
    "старша сестра медична",
    "старша сестра медична операційна",
    "головна сестра медична",
    "молодша сестра медична",
}


def _is_catalogue_form(nominative: str) -> bool:
    return _clean(nominative).casefold() in CATALOGUE_ARTIFACTS


def _read_csv(path: Path) -> list[dict[str, str]]:
    if not path.is_file():
        return []
    return list(csv.DictReader(path.read_text(encoding="utf-8-sig").splitlines(), delimiter=";"))


def _excel_pairs(sheet, first: int, second: int) -> dict[str, tuple[str, str]]:
    """Пари «називний → форма» з двох стовпців аркуша (нумерація з нуля)."""
    table: dict[str, tuple[str, str]] = {}
    for row in sheet.iter_rows(min_row=2, values_only=True):
        left = _clean(row[first]) if first < len(row) else ""
        right = _clean(row[second]) if second < len(row) else ""
        if not left or not right or left.startswith("#") or right.startswith("#"):
            continue
        if "(" in left:
            # «головна медична сестра (головний медичний брат)» — підпис обох
            # родів в одній клітинці, а не назва посади.
            continue
        table.setdefault(left.casefold(), (left, right))
    return table


def read_excel(path: Path) -> dict[str, dict[str, dict[str, str]]]:
    """Форми з Excel: {«посади»: {ключ: {…}}, «звання»: {…}}."""
    from openpyxl import load_workbook

    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        positions: dict[str, dict[str, str]] = {}
        # «посада кандидата»: F, G — називний і знахідний.
        for key, (nominative, form) in _excel_pairs(workbook["посада кандидата"], 5, 6).items():
            positions.setdefault(key, {"Називний": nominative})["Знахідний"] = form
        # «вакантна посада»: E, F — називний і орудний.
        for key, (nominative, form) in _excel_pairs(workbook["вакантна посада"], 4, 5).items():
            positions.setdefault(key, {"Називний": nominative})["Орудний"] = form

        ranks: dict[str, dict[str, str]] = {}
        # «звання кандидата»: H, I — називний і знахідний.
        for key, (nominative, form) in _excel_pairs(workbook["звання кандидата"], 7, 8).items():
            ranks.setdefault(key, {"Називний": nominative})["Знахідний"] = form
        return {"посади": positions, "звання": ranks}
    finally:
        workbook.close()


def merge_positions(excel: dict[str, dict[str, str]]) -> list[dict[str, str]]:
    merged: dict[str, dict[str, str]] = {}

    skipped: list[str] = []

    catalogue: list[str] = []

    def put(key: str, values: dict[str, str], source: str) -> None:
        nominative = _clean(values.get("Називний"))
        if nominative and key not in merged and _is_catalogue_form(nominative):
            catalogue.append(nominative)
            return
        row = merged.setdefault(key, {name: "" for name in POSITION_COLUMNS})
        nominative = nominative or row.get("Називний", "")
        for column, value in values.items():
            value = _clean(value)
            if not value or row.get(column):
                continue
            if column != "Називний" and column != "Примітка" and not _same_shape(nominative, value):
                skipped.append(f"{nominative} -> {value}")
                continue
            row[column] = value
        sources = [part for part in (row.get("Джерело") or "").split(", ") if part]
        if source not in sources:
            sources.append(source)
        row["Джерело"] = ", ".join(sources)

    # 1. Що вже є (разом із ручними правками користувача).
    for row in _read_csv(DICTIONARIES / "positions.csv"):
        nominative = _clean(row.get("Називний"))
        if not nominative:
            continue
        put(
            nominative.casefold(),
            {
                "Називний": nominative,
                "Родовий": row.get("Родовий", ""),
                "Знахідний": row.get("Знахідний", ""),
                "Давальний": row.get("Давальний", ""),
                "Орудний": row.get("Орудний", ""),
                "Примітка": row.get("Примітка", ""),
            },
            _clean(row.get("Джерело")) or "довідник",
        )

    # 2. Excel — вивірені людьми форми.
    for key, values in excel.items():
        put(key, values, "Excel")

    # 3. Зведений довідник: родовий, давальний, орудний.
    for row in _read_csv(ALL_POSITIONS):
        nominative = _clean(row.get("Називний"))
        if not nominative:
            continue
        put(
            nominative.casefold(),
            {
                "Називний": nominative,
                "Родовий": row.get("Родовий", ""),
                "Давальний": row.get("Давальний", ""),
                "Орудний": row.get("Орудний", ""),
            },
            "зведений",
        )

    if catalogue:
        print(f"Пропущено перевернутих назв із каталогу: {len(catalogue)}")
        for line in catalogue[:5]:
            print(f"  {line}")
    if skipped:
        print(f"Пропущено форм, не схожих на свою назву: {len(skipped)}")
        for line in skipped[:5]:
            print(f"  {line}")
    # Давальний: спершу готова форма зі зведеного довідника, інакше правило.
    # Старі форми, виведені грубішим правилом, тут перезаписуються.
    ready_datives = {
        _clean(row.get("Називний")).casefold(): _clean(row.get("Давальний"))
        for row in _read_csv(ALL_POSITIONS)
        if _clean(row.get("Давальний"))
    }
    for key, row in merged.items():
        ready = ready_datives.get(key)
        row["Давальний"] = ready if ready and _same_shape(row["Називний"], ready) else dative_form(
            row["Називний"]
        )

    for row in merged.values():
        if not row["Родовий"] and row["Знахідний"] and not _leading_first_declension(row["Називний"]):
            row["Родовий"] = row["Знахідний"]
        if not row["Родовий"]:
            # Родовий = давальний без закінчення? Ні: виводимо зі знахідного або
            # лишаємо порожнім лише тоді, коли нема з чого — таких рядків немає.
            row["Родовий"] = accusative_form(row["Називний"], "") or row["Давальний"]
        if not row["Знахідний"]:
            row["Знахідний"] = accusative_form(row["Називний"], row["Родовий"])
        if not row["Орудний"]:
            row["Орудний"] = instrumental_form(row["Називний"])
    return [merged[key] for key in sorted(merged, key=lambda key: merged[key]["Називний"].casefold())]


def merge_ranks(excel: dict[str, dict[str, str]]) -> list[dict[str, str]]:
    merged: dict[str, dict[str, str]] = {}
    for row in _read_csv(DICTIONARIES / "ranks.csv"):
        nominative = _clean(row.get("Називний"))
        if not nominative:
            continue
        merged[nominative.casefold()] = {
            "Називний": nominative,
            "Родовий": _clean(row.get("Родовий")),
            "Знахідний": _clean(row.get("Знахідний")),
            "Давальний": _clean(row.get("Давальний")),
            "Джерело": _clean(row.get("Джерело")) or "довідник",
            "Примітка": _clean(row.get("Примітка")),
        }
    for key, values in excel.items():
        row = merged.setdefault(
            key,
            {name: "" for name in RANK_COLUMNS} | {"Називний": values["Називний"], "Джерело": "Excel"},
        )
        for column, value in values.items():
            if _clean(value) and not row.get(column):
                row[column] = _clean(value)
        if "Excel" not in row["Джерело"]:
            row["Джерело"] = ", ".join(part for part in (row["Джерело"], "Excel") if part)
    for row in merged.values():
        if not row["Знахідний"] and row["Родовий"]:
            row["Знахідний"] = row["Родовий"]
    return [merged[key] for key in merged]


def write_csv(path: Path, columns: list[str], rows: list[dict[str, str]]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle, delimiter=";", lineterminator="\n")
        writer.writerow(columns)
        for row in rows:
            writer.writerow([row.get(column, "") for column in columns])


def main(argv: list[str]) -> int:
    excel_path = Path(argv[1]) if len(argv) > 1 else DEFAULT_EXCEL
    if not excel_path.is_file():
        print(f"Не знайдено файл Excel: {excel_path}")
        print("Передайте шлях аргументом: python scripts/order_index/build_position_forms.py <файл.xlsm>")
        return 1
    excel = read_excel(excel_path)
    print(f"З Excel: посад {len(excel['посади'])}, звань {len(excel['звання'])}")

    positions = merge_positions(excel["посади"])
    write_csv(DICTIONARIES / "positions.csv", POSITION_COLUMNS, positions)
    filled = sum(1 for row in positions if row["Давальний"])
    print(f"positions.csv: {len(positions)} посад, з них із давальним {filled}")

    ranks = merge_ranks(excel["звання"])
    write_csv(DICTIONARIES / "ranks.csv", RANK_COLUMNS, ranks)
    print(f"ranks.csv: {len(ranks)} звань")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
