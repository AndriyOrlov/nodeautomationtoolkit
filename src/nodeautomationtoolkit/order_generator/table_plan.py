"""Таблиця з даними про осіб (план звільнення, список до звання) → записи.

План переміщення має сталу форму (додаток 16), і його розбирає `plan.py`.
План звільнення (додаток 21) і списки до присвоєння звань у кожній частині
роблять по-своєму: назви граф однакові за змістом, а порядок і формулювання —
ні. Тому тут колонки впізнаються **за заголовком**, а не за номером, а перелік
назв лежить окремим файлом `personnel/dictionaries/table_columns.csv`, який
користувач доповнює сам — без зміни коду.

Що не впізналось — не губиться мовчки: `TablePlan.unknown_columns` повертає
заголовки, яких немає в довіднику, і вікно показує їх у журналі.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from openpyxl import load_workbook

from ..personnel.dictionaries import load_keyed_table, normalize_key
from .record import PLAN, PersonRecord, Position, Value

COLUMNS_FILE = "table_columns.csv"

#: Поля запису, які можна заповнити з таблиці (решта — з індексу чи вручну).
KNOWN_FIELDS = {
    "rank",
    "full_name",
    "surname",
    "name",
    "patronymic",
    "ipn",
    "birth",
    "education",
    "service_since",
    "position",
    "shpk",
    "vos",
    "tariff",
    "new_position",
    "new_shpk",
    "basis",
    "dismissal",
    "destination",
    "service_calendar",
    "service_privileged",
    "registration",
    "uniform",
    "dismissal_note",
    "new_rank",
    "rank_seniority",
    "rank_since",
    "rank_note",
}


@dataclass
class TablePlan:
    """Прочитана таблиця: записи про осіб і те, чого програма не зрозуміла."""

    path: Path
    records: list[PersonRecord] = field(default_factory=list)
    header_row: int = 0
    columns: dict[str, str] = field(default_factory=dict)  # заголовок → поле
    unknown_columns: list[str] = field(default_factory=list)
    problems: list[str] = field(default_factory=list)


def column_titles(user_directory: str | Path | None = None) -> dict[str, str]:
    """Заголовок графи (нормалізований) → назва поля запису."""
    table = {}
    for row in load_keyed_table(COLUMNS_FILE, user_directory).values():
        title = (row.get("Заголовок") or "").strip()
        field_name = (row.get("Поле") or "").strip()
        if title and field_name in KNOWN_FIELDS:
            table[normalize_key(title)] = field_name
    return table


def _cell(value) -> str:
    if value is None:
        return ""
    if hasattr(value, "strftime"):
        return value.strftime("%d.%m.%Y")
    return " ".join(str(value).split())


def _find_header(sheet, titles: dict[str, str]) -> tuple[int, dict[int, str], list[str]]:
    """Рядок заголовків — той, де впізналось найбільше граф (щонайменше три)."""
    best: tuple[int, dict[int, str], list[str]] = (0, {}, [])
    for row in sheet.iter_rows(min_row=1, max_row=min(sheet.max_row, 30)):
        mapping: dict[int, str] = {}
        unknown: list[str] = []
        for cell in row:
            text = _cell(cell.value)
            if not text:
                continue
            field_name = titles.get(normalize_key(text))
            if field_name:
                mapping.setdefault(cell.column, field_name)
            elif len(text) > 2 and not text.isdigit():
                unknown.append(text)
        if len(mapping) > len(best[1]):
            best = (row[0].row, mapping, unknown)
    return best if len(best[1]) >= 3 else (0, {}, [])


def _split_name(record: PersonRecord, text: str) -> None:
    """«ІВАНЕНКО Олексій Вікторович» в одній графі — розкладаємо на три поля."""
    parts = text.replace(",", " ").split()
    if not parts:
        return
    record.surname = Value(parts[0], PLAN)
    if len(parts) > 1:
        record.name = Value(parts[1], PLAN)
    if len(parts) > 2:
        record.patronymic = Value(" ".join(parts[2:]), PLAN)


def _apply(record: PersonRecord, field_name: str, text: str) -> None:
    if not text:
        return
    if field_name == "full_name":
        _split_name(record, text)
    elif field_name == "position":
        record.current = Position(text=Value(text, PLAN), **_position_extra(record.current))
    elif field_name in ("shpk", "vos", "tariff"):
        setattr(record.current, field_name, Value(text, PLAN))
    elif field_name == "new_position":
        record.target = Position(text=Value(text, PLAN), **_position_extra(record.target))
    elif field_name == "new_shpk":
        record.target.shpk = Value(text, PLAN)
    elif hasattr(record, field_name):
        setattr(record, field_name, Value(text, PLAN))


def _position_extra(position: Position) -> dict:
    return {name: getattr(position, name) for name in ("shpk", "vos", "tariff")}


def read_table(path: str | Path, user_directory: str | Path | None = None) -> TablePlan:
    """Читає таблицю й повертає записи про осіб (без звірки з індексом)."""
    path = Path(path)
    plan = TablePlan(path=path)
    titles = column_titles(user_directory)
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = workbook.worksheets[0]
        header_row, mapping, unknown = _find_header(sheet, titles)
        if not mapping:
            plan.problems.append(
                "не вдалося впізнати жодної графи таблиці — додайте заголовки "
                "у довідник граф (table_columns.csv)"
            )
            return plan
        plan.header_row = header_row
        plan.columns = {str(column): field_name for column, field_name in mapping.items()}
        plan.unknown_columns = unknown

        for row in sheet.iter_rows(min_row=header_row + 1, max_row=sheet.max_row):
            values = {cell.column: _cell(cell.value) for cell in row}
            if not any(values.get(column) for column in mapping):
                continue
            record = PersonRecord()
            for column, field_name in mapping.items():
                _apply(record, field_name, values.get(column, ""))
            if not record.surname:
                continue
            plan.records.append(record)
    finally:
        workbook.close()
    return plan
