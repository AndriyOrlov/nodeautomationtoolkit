"""Вища, нижча чи рівнозначна посада — перенесено з аркуша «на посаду» Excel-генератора.

Правило з Excel:
- шпк однієї з посад невідомий (зокрема запасна рота) → просто «Призначається»;
- шпк однаковий → порівнюється тарифний розряд: однаковий — рівнозначна,
  у старої посади більший — нижча, менший — вища;
- шпк різний → у старої посади вищий шпк — нижча, нижчий — вища.

Формулювання висновку залежить ще й від висновку ВЛК у підставах. Коефіцієнти
шпк (`shpk_levels.csv`) і тексти (`appointment_wording.csv`) — окремі довідники,
їх правлять без зміни коду. Важливий лише порядок коефіцієнтів, а не самі числа.

Модуль ще НЕ підключено до генераторів — це заготовка для генератора наказів.
"""

from __future__ import annotations

import re
from pathlib import Path

from nodeautomationtoolkit.personnel.dictionaries import load_keyed_table, normalize_key

SHPK_LEVELS_FILE = "shpk_levels.csv"
WORDING_FILE = "appointment_wording.csv"

HIGHER = "вища"
LOWER = "нижча"
EQUAL = "рівнозначна"
UNKNOWN = "невідомо"

_QUOTES_RE = re.compile("[\"“”„«»'’‘]")
_TARIFF_RE = re.compile(r"(\d+)\s*(?:т\.?\s?р\.?|тарифн)|т\.?\s?р\.?\s*(\d+)", re.IGNORECASE)
_VLK_RE = re.compile(r"\bВЛК\b|військово-лікарськ\w*\s+комісі|обмежено\s+придатн", re.IGNORECASE)


def shpk_key(text: str) -> str:
    """«шпк “старший солдат”», «шпк - солдат», «ШПК, сержант» → «старший солдат»."""
    value = _QUOTES_RE.sub(" ", str(text or ""))
    value = re.sub(r"^\s*шпк\b[\s,\-–—:]*", "", value, flags=re.IGNORECASE)
    return normalize_key(value.strip(" ,.;"))


def shpk_level(text: str, user_directory: str | Path | None = None) -> int | None:
    row = load_keyed_table(SHPK_LEVELS_FILE, user_directory).get(shpk_key(text))
    if not row:
        return None
    try:
        return int(row.get("Коефіцієнт", ""))
    except ValueError:
        return None


def tariff_grade(text: str) -> int | None:
    """«12 т.р.», «т.р. 12», «12 тарифний розряд» → 12."""
    match = _TARIFF_RE.search(str(text or ""))
    if not match:
        return None
    return int(match.group(1) or match.group(2))


def compare_positions(
    from_shpk: str,
    to_shpk: str,
    from_tariff: str = "",
    to_tariff: str = "",
    user_directory: str | Path | None = None,
) -> str:
    """Порівнює стару й нову посаду: HIGHER / LOWER / EQUAL / UNKNOWN."""
    from_level = shpk_level(from_shpk, user_directory)
    to_level = shpk_level(to_shpk, user_directory)
    if from_level is None or to_level is None:
        return UNKNOWN
    if from_level != to_level:
        return LOWER if from_level > to_level else HIGHER
    from_grade, to_grade = tariff_grade(from_tariff), tariff_grade(to_tariff)
    if from_grade is None or to_grade is None or from_grade == to_grade:
        return EQUAL
    return LOWER if from_grade > to_grade else HIGHER


def mentions_vlk(basis_text: str) -> bool:
    """Чи є в підставах висновок ВЛК.

    Ознаки: «ВЛК», «військово-лікарської комісії», «обмежено придатний». Передавати
    лише текст підстав: у самому пункті «ВЛК» трапляється й поза висновком.
    """
    return bool(_VLK_RE.search(str(basis_text or "")))


def appointment_wording(
    level: str, vlk: bool = False, user_directory: str | Path | None = None
) -> str:
    """Текст висновку «Призначається на … посаду …» з довідника формулювань."""
    table = load_keyed_table(WORDING_FILE, user_directory)
    keys = [f"{level}_влк", level] if vlk else [level]
    for key in keys:
        row = table.get(normalize_key(key))
        if row and row.get("Текст"):
            return row["Текст"]
    return ""
