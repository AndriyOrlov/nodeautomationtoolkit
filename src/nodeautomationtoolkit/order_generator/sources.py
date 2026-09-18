"""Перехідники: план переміщення, пункт попереднього наказу, ручний ввід → `PersonRecord`.

Кожне джерело заповнює лише те, що в ньому справді є. Поєднує їх `merge.py`.
"""

from __future__ import annotations

import re

from ..order_index.store import PersonItem
from .plan import PlanEntry
from .record import MANUAL, ORDER, PLAN, PersonRecord, Position, Value

_BIRTH_RE = re.compile(r"(\d{4})\s*р\.\s*н\.")
_EDUCATION_RE = re.compile(r"освіта:\s*(.+?)(?=(?:,\s*)?у\s+ЗС\b|$)", re.IGNORECASE | re.DOTALL)
_SERVICE_RE = re.compile(r"у\s+ЗС\s*[-–]?\s*(?:із|з)\s+(\d{2}\.\d{4})", re.IGNORECASE)
_APPOINTED_RE = re.compile(r"Призначається[^.]*\.", re.IGNORECASE)
_EXTRA_RE = re.compile(r"(Кандидат[^.]*наук\.|Доктор[^.]*наук\.)", re.IGNORECASE)


def from_plan(entry: PlanEntry) -> PersonRecord:
    """Рядок плану переміщення: особа, займана посада, вакансія, підстава."""
    candidate, vacancy, bio = entry.candidate, entry.vacancy, entry.biography
    record = PersonRecord(
        rank=Value(candidate.rank, PLAN),
        surname=Value(candidate.surname, PLAN),
        name=Value(candidate.name, PLAN),
        patronymic=Value(candidate.patronymic, PLAN),
        ipn=Value(candidate.ipn, PLAN),
        birth=Value(bio.birth_year, PLAN),
        education=Value(bio.education, PLAN),
        service_since=Value(bio.service_since, PLAN),
        current=Position(
            text=Value(candidate.former_position or candidate.position, PLAN),
            shpk=Value(candidate.params.shpk, PLAN),
            vos=Value(candidate.params.vos, PLAN),
            tariff=Value(candidate.params.tariff, PLAN),
        ),
        target=Position(
            text=Value(vacancy.position, PLAN),
            shpk=Value(vacancy.params.shpk, PLAN),
            vos=Value(vacancy.params.vos, PLAN),
            tariff=Value(vacancy.params.tariff, PLAN),
        ),
        basis=Value(entry.basis, PLAN),
        evaluation=Value(entry.evaluation, PLAN),
    )
    record.problems.extend(entry.problems)
    return record


def from_order_item(item: PersonItem) -> PersonRecord:
    """Пункт попереднього наказу: звірена біографія і посада, на яку призначали."""
    text = item.text
    birth = _BIRTH_RE.search(text)
    education = _EDUCATION_RE.search(text)
    service = _SERVICE_RE.search(text)
    extra = _EXTRA_RE.search(text)
    record = PersonRecord(
        rank=Value(item.rank, ORDER),
        surname=Value(item.surname, ORDER),
        name=Value(item.name, ORDER),
        patronymic=Value(item.patronymic, ORDER),
        ipn=Value(item.ipn, ORDER),
        birth=Value(birth.group(1) if birth else "", ORDER),
        education=Value(_clean(education.group(1)) if education else "", ORDER),
        service_since=Value(service.group(1) if service else "", ORDER),
        extra_bio=Value(extra.group(1) if extra else "", ORDER),
        # Посада, НА ЯКУ призначили цим наказом, — це посада, З ЯКОЇ призначають тепер.
        current=Position(text=Value(item.target_position or item.current_position, ORDER)),
        previous_item=Value(text, ORDER),
    )
    where = f"наказ № {item.order_number}" if item.order_number else item.path
    if item.order_date:
        where += f" від {item.order_date[8:10]}.{item.order_date[5:7]}.{item.order_date[:4]}"
    record.notes.append(f"Пункт узято з: {where}")
    return record


def from_manual(values: dict[str, str]) -> PersonRecord:
    """Ручний ввід: ті самі імена полів, що в `PersonRecord` («target.text», «rank» …)."""
    record = PersonRecord()
    for key, raw in values.items():
        text = " ".join(str(raw or "").split())
        if not text:
            continue
        if "." in key:
            group, _, field_name = key.partition(".")
            position = getattr(record, group, None)
            if isinstance(position, Position) and hasattr(position, field_name):
                setattr(position, field_name, Value(text, MANUAL))
        elif hasattr(record, key):
            setattr(record, key, Value(text, MANUAL))
    return record


def _clean(text: str) -> str:
    return " ".join(str(text or "").split()).strip(" ,")
