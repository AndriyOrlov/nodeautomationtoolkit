"""Збирає єдиний словник посад індексатора: `order_index/positions_all.csv`.

Джерела:
- `personnel/dictionaries/positions.csv` — посади з Excel-генератора (з їхніми формами);
- `personnel/dictionaries/mo317_positions.csv` — перелік штатних посад наказу МО № 317
  (коди посад, ВОС, тарифні розряди);
- посади офіцерського складу, складені тут: «посада × підрозділ»
  (командир роти, начальник відділу, старший офіцер управління …);
- рядки самого `positions_all.csv` з джерелом «вручну» — переносяться без змін.

Форми (родовий, давальний, орудний) будує `order_index/position_forms.py`; форми з
Excel, що відрізняються, лишаються в «Інші форми» — щоб розпізнавати й такі написання.

    python scripts/order_index/build_all_positions.py
"""

from __future__ import annotations

import csv
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "src"))

from nodeautomationtoolkit.order_index.position_dictionary import (  # noqa: E402
    ALL_POSITIONS_FILE,
    BASE_POSITIONS_FILE,
    MO317_POSITIONS_FILE,
    strip_brackets,
)
from nodeautomationtoolkit.order_index.position_forms import decline  # noqa: E402

COLUMNS = [
    "Називний", "Родовий", "Давальний", "Орудний", "Інші форми", "Склад",
    "Код посади (МО 317)", "Номери ВОС (МО 317)", "Усі ВОС", "Назва в переліку МО 317", "Джерело",
]
MANUAL = "вручну"
RANK_AND_FILE = "рядовий, сержантський, старшинський"
OFFICERS = "офіцерський"

_UNITS = ["відділення", "взводу", "роти", "батареї", "батальйону", "дивізіону", "полку", "бригади", "групи",
          "загону", "ескадрильї", "ланки", "екіпажу", "військової частини", "корпусу", "командування"]
_BODIES = ["відділу", "відділення", "управління", "служби", "штабу", "центру", "бази", "пункту", "кафедри",
           "факультету", "курсу", "лабораторії", "госпіталю", "клініки", "групи", "секції", "сектору", "складу",
           "майстерні", "станції", "вузла", "штабу батальйону", "штабу полку", "штабу бригади", "апарату",
           "департаменту", "інспекції", "комендатури", "навчального центру", "територіального центру комплектування та соціальної підтримки"]
_STAFF = ["відділу", "відділення", "управління", "групи", "штабу", "служби", "сектору", "центру", "напряму", "секції"]
_BRANCHES = ["озброєння", "тилу", "логістики", "морально-психологічного забезпечення", "роботи з особовим складом",
             "персоналу", "бойової підготовки", "повітряно-десантної підготовки", "технічного забезпечення",
             "психологічної підтримки персоналу", "цивільно-військового співробітництва", "зв'язку",
             "інженерного забезпечення", "ракетно-артилерійського озброєння", "безпеки військової служби"]
_BRANCH_UNITS = ["роти", "батареї", "батальйону", "дивізіону", "полку", "бригади", "військової частини"]
_OFFICER_SINGLE = [
    "командир", "командувач", "заступник командувача", "начальник штабу", "офіцер", "старший офіцер",
    "головний офіцер", "офіцер-психолог", "психолог", "юрисконсульт", "старший юрисконсульт", "перекладач",
    "старший перекладач", "військовий капелан", "військовий комендант", "помічник командира",
    "інженер", "старший інженер", "провідний інженер", "технік", "старший технік", "лікар", "старший лікар",
    "лікар-ординатор", "ординатор", "старший ординатор", "викладач", "старший викладач", "доцент", "професор",
    "науковий співробітник", "старший науковий співробітник", "провідний науковий співробітник",
    "головний науковий співробітник", "докторант", "ад'юнкт", "слухач", "курсант", "льотчик", "старший льотчик",
    "льотчик-оператор", "штурман", "старший штурман", "інспектор", "старший інспектор", "головний інспектор",
    "старший інспектор-льотчик", "старший інспектор-штурман", "військовий спостерігач", "оперуповноважений",
    "старший оперуповноважений", "черговий", "оперативний черговий", "помічник оперативного чергового",
    "військовослужбовець", "відповідальний виконавець", "ад'ютант", "комендант", "кореспондент",
    "старший помічник", "помічник начальника", "заступник начальника", "начальник",
    "старший авіаційний технік", "авіаційний технік", "старший бортовий технік", "бортовий технік",
    "головний сержант", "старшина", "перший заступник командувача", "помічник командувача",
    "начальник медичної служби", "начальник фінансово-економічної служби", "начальник служби пального",
    "начальник продовольчої служби", "начальник речової служби", "начальник автомобільної служби",
    "начальник бронетанкової служби", "начальник інженерної служби", "начальник хімічної служби",
    "начальник служби радіаційного, хімічного, біологічного захисту", "начальник зв'язку",
    "начальник розвідки", "начальник артилерії", "начальник протиповітряної оборони",
    "начальник фізичної підготовки", "начальник клубу", "начальник оркестру", "начальник кафедри",
    "начальник курсу", "начальник факультету", "командир навчальної групи",
]


def officer_positions() -> list[str]:
    names = list(_OFFICER_SINGLE)
    for unit in _UNITS:
        names += [f"командир {unit}", f"заступник командира {unit}", f"перший заступник командира {unit}",
                  f"помічник командира {unit}", f"начальник штабу {unit}", f"заступник начальника штабу {unit}",
                  f"старший офіцер {unit}", f"офіцер {unit}", f"головний сержант {unit}", f"старшина {unit}"]
    for body in _BODIES:
        names += [f"начальник {body}", f"заступник начальника {body}", f"помічник начальника {body}"]
    for staff in _STAFF:
        names += [f"старший офіцер {staff}", f"офіцер {staff}", f"головний офіцер {staff}"]
    for unit in _BRANCH_UNITS:
        for branch in _BRANCHES:
            names.append(f"заступник командира {unit} з {branch}")
        names.append(f"начальник штабу - перший заступник командира {unit}")
    return names


def _read(path: Path) -> list[dict[str, str]]:
    if not path.is_file():
        return []
    with open(path, encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle, delimiter=";"))


def display_name(name: str) -> str:
    """Назва для словника: без дужок, малими, але тире між двома посадами лишається « - »."""
    text = " ".join(str(name or "").split()).casefold()
    parts = [strip_brackets(part) for part in re.split(r"\s+[-–—]\s+", text)]
    return " - ".join(part for part in parts if part)


def _norm(text: str) -> str:
    return " ".join(str(text or "").split()).casefold()


def build() -> list[dict[str, str]]:
    entries: dict[str, dict[str, str]] = {}

    def entry(nominative: str, source: str, staff: str) -> dict[str, str]:
        key = strip_brackets(nominative)
        row = entries.get(key)
        if row is None:
            row = dict.fromkeys(COLUMNS, "")
            row["Називний"] = display_name(nominative)
            for column, case in (("Родовий", "Р"), ("Давальний", "Д"), ("Орудний", "О")):
                row[column] = decline(row["Називний"], case)
            entries[key] = row
        sources = [s for s in row["Джерело"].split(", ") if s]
        if source not in sources:
            row["Джерело"] = ", ".join([*sources, source])
        staffs = [s for s in row["Склад"].split("; ") if s]
        if staff and staff not in staffs:
            row["Склад"] = "; ".join([*staffs, staff])
        return row

    for manual in _read(ALL_POSITIONS_FILE):
        if MANUAL in manual.get("Джерело", ""):
            entries[strip_brackets(manual["Називний"])] = {column: manual.get(column, "") for column in COLUMNS}

    for row in _read(BASE_POSITIONS_FILE):
        nominative = _norm(row.get("Називний"))
        if not nominative or nominative.startswith("#"):
            continue
        target = entry(nominative, "Excel", RANK_AND_FILE)
        generated = {target["Родовий"], target["Орудний"]}
        others = [f for f in target["Інші форми"].split(" | ") if f]
        for column in ("Знахідний", "Орудний"):
            form = _norm(row.get(column))
            if form and form not in generated and form not in others:
                others.append(form)
        target["Інші форми"] = " | ".join(others)

    for row in _read(MO317_POSITIONS_FILE):
        name = row.get("Найменування посади", "")
        if not name:
            continue
        target = entry(name, "МО 317", RANK_AND_FILE)
        for column, value in (("Код посади (МО 317)", row.get("Код посади", "")),
                              ("Назва в переліку МО 317", " ".join(name.split()))):
            values = [v for v in target[column].split(", " if column.startswith("Код") else " | ") if v]
            if value and value not in values:
                target[column] = (", " if column.startswith("Код") else " | ").join([*values, value])
        vos = {v for v in target["Номери ВОС (МО 317)"].split(", ") if v}
        vos.update(v for v in row.get("Номери ВОС", "").split(", ") if v)
        target["Номери ВОС (МО 317)"] = ", ".join(sorted(vos))
        if row.get("Усі ВОС"):
            target["Усі ВОС"] = "так"

    for name in officer_positions():
        staff = RANK_AND_FILE if name.startswith(("головний сержант", "старшина")) else OFFICERS
        entry(name, "складено", staff)

    return sorted(entries.values(), key=lambda row: row["Називний"])


def main() -> int:
    rows = build()
    with open(ALL_POSITIONS_FILE, "w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=COLUMNS, delimiter=";")
        writer.writeheader()
        writer.writerows(rows)
    by_source: dict[str, int] = {}
    for row in rows:
        for source in row["Джерело"].split(", "):
            by_source[source] = by_source.get(source, 0) + 1
    print(f"{ALL_POSITIONS_FILE.name}: {len(rows)} посад; за джерелами: {by_source}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
