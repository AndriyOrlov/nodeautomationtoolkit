"""Від кинутого файла до готового запису про особу.

Кроки:
1. прочитати документ і витягти з нього все, що є (`documents.py`);
2. для кожної знайденої особи взяти з індексу її найсвіжіший пункт наказу
   (за РНОКПП, інакше за ПІБ) — звідти вивірена біографія і посада, на яку
   її призначали минулого разу;
3. поєднати джерела (`merge.py`) і **м'яко звірити**: збіги йдуть у примітки,
   розбіжності — у проблеми, але запис усе одно будується з того, що є.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

from ..order_index.store import PersonItem, find_person_items
from . import merge, sources
from .documents import DocumentFacts, PersonMention, read_document
from .record import DOCUMENT, PersonRecord, Position, Value


@dataclass
class ResolvedPerson:
    record: PersonRecord
    mention: PersonMention
    previous: list[PersonItem] = field(default_factory=list)


@dataclass
class ResolvedDocument:
    facts: DocumentFacts
    people: list[ResolvedPerson] = field(default_factory=list)

    @property
    def kind(self) -> str:
        return self.facts.kind


def record_from_document(facts: DocumentFacts, person: PersonMention) -> PersonRecord:
    """Запис із самого документа: що знайшлося поруч із цією особою."""
    record = PersonRecord(
        rank=Value(person.rank, DOCUMENT),
        surname=Value(person.surname, DOCUMENT),
        name=Value(person.name, DOCUMENT),
        patronymic=Value(person.patronymic, DOCUMENT),
        ipn=Value(person.ipn or facts.first("ipn"), DOCUMENT),
        birth=Value(_year(facts.first("birth")), DOCUMENT),
        education=Value(facts.first("education"), DOCUMENT),
        service_since=Value(facts.first("service_since"), DOCUMENT),
        basis=Value(facts.first("basis") or facts.first("law_reference"), DOCUMENT),
    )
    # Посада «на яку» береться лише за явною ознакою («до призначення на посаду»).
    # Друга згадана посада нею не є: у рапорті це адресат, а не нове призначення.
    target = facts.first("target_position")
    positions = [position for position in facts.positions if position != target]
    if positions:
        record.current = Position(text=Value(positions[0], DOCUMENT))
    if target:
        record.target = Position(
            text=Value(target, DOCUMENT),
            shpk=Value(facts.first("shpk"), DOCUMENT),
            vos=Value(facts.first("vos"), DOCUMENT),
            tariff=Value(facts.first("tariff"), DOCUMENT),
        )
    if facts.first("consent"):
        record.notes.append(f"Згода в документі: «{facts.first('consent')}»")
    if facts.first("discharge"):
        record.notes.append(f"Ознака звільнення: «{facts.first('discharge')}»")
    record.notes.append(f"Документ: {facts.kind} ({Path(facts.path).name or 'без файла'})")
    return record


def resolve_document(path: str | Path, index_folder: str | Path | None = None, limit: int = 5) -> ResolvedDocument:
    facts = read_document(path)
    return resolve_facts(facts, index_folder, limit)


def resolve_facts(
    facts: DocumentFacts, index_folder: str | Path | None = None, limit: int = 5
) -> ResolvedDocument:
    result = ResolvedDocument(facts=facts)
    for person in facts.people:
        document_record = record_from_document(facts, person)
        previous: list[PersonItem] = []
        if index_folder:
            previous = find_person_items(
                index_folder, ipn=person.ipn or facts.first("ipn"), full_name=person.full_name, limit=limit
            )
        order_record = sources.from_order_item(previous[0]) if previous else None
        record = merge.build_record(document=document_record, order=order_record)
        if order_record is None and index_folder:
            record.notes.append("У наказах особу не знайдено — усе взято з документа")
        _verify(record, document_record, previous)
        result.people.append(ResolvedPerson(record=record, mention=person, previous=previous))
    return result


def _verify(record: PersonRecord, document: PersonRecord, previous: list[PersonItem]) -> None:
    """М'яка звірка документа з наказами: збіги — в примітки, різниця — в проблеми."""
    if not previous:
        return
    item = previous[0]
    checks = (
        ("рік народження", str(document.birth), _year(item.text)),
        ("РНОКПП", str(document.ipn), item.ipn),
    )
    for title, from_document, from_order in checks:
        if not from_document or not from_order:
            continue
        if from_document == from_order:
            record.notes.append(f"{title} збігається з наказом № {item.order_number or '—'}")
        else:
            record.problems.append(
                f"{title} різний: у документі «{from_document}», у наказі № {item.order_number or '—'} "
                f"«{from_order}»"
            )
    if document.current and item.target_position:
        if merge._same_position(str(document.current.text), item.target_position):
            record.notes.append("Посада в документі збігається з посадою з останнього наказу")
        else:
            record.problems.append(
                f"посада в документі «{document.current.text}» не збігається з останнім наказом "
                f"«{item.target_position}»"
            )
    if len(previous) > 1:
        record.notes.append(f"Знайдено пунктів про особу: {len(previous)}")


def _year(text: str) -> str:
    match = re.search(r"(?:19|20)\d{2}", str(text or ""))
    return match.group(0) if match else ""
