"""Складання тексту наказу з готових записів про осіб (етап «генерація»).

Формулювання взяті зі зразків додатка 53 до Інструкції з діловодства (наказ
МО № 170), а не вигадані. Груповий пункт про призначення виглядає так:

    Відповідно до пункту ___ Положення про проходження громадянами України
    військової служби у Збройних Силах України нижчепойменованих осіб
    офіцерського складу … ЗВІЛЬНИТИ з займаних посад і ПРИЗНАЧИТИ:

    1. Полковника ЛІВІЦЬКОГО Олександра Станіславовича, начальника управління
    особового складу штабу … - НАЧАЛЬНИКОМ УПРАВЛІННЯ ПЕРСОНАЛУ ШТАБУ …,
    ВОС - 2905001.
    1965 р.н., освіта: НАОУ (оср) у 2003 р., у ЗС - із 09.1980.
    2649423014.
    Призначається на вищу посаду у порядку просування по службі
    з шпк «полковник» на шпк «полковник».

Звання й посада «з якої» — у знахідному, посада «на яку» — в орудному ВЕЛИКИМИ,
ПІБ — у знахідному (`names.py`). Відмінки беруться з довідників пакета
`personnel`, які користувач править сам; чого немає в довіднику — лишається як
є, а перевірка (`check.py`) про це попереджає.

Порядок осіб і нумерація — рішення користувача 18.09.2026: сортування за
прізвищем, наскрізна нумерація пунктів.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from ..builtin_nodes.typography import apply_ukrainian_typography, clean_duplicated_units
from ..personnel import appointment, declension
from .names import FullName, accusative
from .record import PersonRecord

#: Стала частина шапки — з усіх зразків додатка 53.
REGULATION = (
    "Положення про проходження громадянами України військової служби "
    "у Збройних Силах України"
)
DEFAULT_KIND = "осіб офіцерського складу"


@dataclass
class OrderParams:
    """Реквізити наказу, які вводить користувач (в Excel це були колонки Z–AJ)."""

    number: str = ""
    date: str = ""
    section: str = ""  # «§ 1»; порожньо — наказ без розділів
    points: str = "пункту ___"  # пункти Положення: «пунктів 45, 50»
    kind: str = DEFAULT_KIND  # тип особового складу
    unit: str = ""  # чиїх осіб звільняють — у родовому, як стоїть у наказі
    target_unit: str = ""  # «ПРИЗНАЧИТИ ДО …:», коли всі йдуть в одну частину
    bases: list[str] = field(default_factory=list)  # директиви, план переміщення
    footer_bases: list[str] = field(default_factory=list)  # «Підстава: …» в кінці
    numbering_start: int = 1
    sort_by_surname: bool = True
    dictionary_folder: str = ""  # тека користувацьких довідників відмінків


@dataclass
class OrderItem:
    """Один пункт наказу: номер, рядки тексту й запис, з якого його склали."""

    number: int
    lines: list[str]
    record: PersonRecord
    notes: list[str] = field(default_factory=list)

    @property
    def text(self) -> str:
        return "\n".join(self.lines)


@dataclass
class OrderDraft:
    """Проєкт наказу: шапка, пункти, підстави в кінці."""

    params: OrderParams
    heading: list[str] = field(default_factory=list)
    items: list[OrderItem] = field(default_factory=list)
    footer: list[str] = field(default_factory=list)

    @property
    def text(self) -> str:
        return "\n".join(numbered_lines(self))


def _clean(text) -> str:
    return " ".join(str(text or "").split())


def rank_accusative(rank: str, folder: str = "") -> tuple[str, bool]:
    """«полковник» → «полковника»; другим значенням — чи знайдено в довіднику."""
    result = declension.decline_rank(_clean(rank), "З", folder or None)
    return result.text, result.found


def position_accusative(position: str, folder: str = "") -> tuple[str, bool]:
    """Посада, з якої звільняють: «командир взводу» → «командира взводу»."""
    result = declension.decline_position(_clean(position), "З", folder or None)
    return result.text, result.found


def position_instrumental(position: str, folder: str = "") -> tuple[str, bool]:
    """Посада, на яку призначають: «командир взводу» → «КОМАНДИРОМ ВЗВОДУ»."""
    result = declension.decline_position(_clean(position), "О", folder or None)
    return result.text.upper(), result.found


def appointment_line(record: PersonRecord, folder: str = "") -> str:
    """«Призначається на вищу посаду … з шпк «майор» на шпк «підполковник».»"""
    from_shpk, to_shpk = _clean(record.current.shpk), _clean(record.target.shpk)
    level = appointment.compare_positions(
        from_shpk,
        to_shpk,
        _clean(record.current.tariff),
        _clean(record.target.tariff),
        folder or None,
    )
    wording = appointment.appointment_wording(
        level, appointment.mentions_vlk(str(record.basis)), folder or None
    )
    if not wording:
        return ""
    if from_shpk and to_shpk:
        wording += (
            f" з шпк «{appointment.shpk_key(from_shpk)}» "
            f"на шпк «{appointment.shpk_key(to_shpk)}»"
        )
    return wording.rstrip(".") + "."


def biography_line(record: PersonRecord) -> str:
    """«1965 р.н., освіта: НАОУ (оср) у 2003 р., у ЗС - із 09.1980.»"""
    parts = []
    if record.birth:
        parts.append(f"{_clean(record.birth)} р.н.")
    if record.education:
        parts.append(f"освіта: {_clean(record.education)}")
    if record.service_since:
        parts.append(f"у ЗС - із {_clean(record.service_since)}")
    return ", ".join(parts) + "." if parts else ""


def main_line(record: PersonRecord, folder: str = "") -> str:
    """Перший рядок пункту: звання, ПІБ, посада «з якої» - ПОСАДА «НА ЯКУ», ВОС."""
    rank, _ = rank_accusative(str(record.rank), folder)
    full_name = accusative(
        FullName(str(record.surname), str(record.name), str(record.patronymic))
    )
    head = " ".join(
        part for part in (rank[:1].upper() + rank[1:] if rank else "", full_name) if part
    )
    line = head
    current, _ = position_accusative(str(record.current.text), folder)
    if current:
        line += f", {current}"
    target, _ = position_instrumental(str(record.target.text), folder)
    if target:
        line += f" - {target}"
    vos = _clean(record.target.vos) or _clean(record.current.vos)
    if vos:
        line += f", ВОС - {vos}"
    return line.rstrip(" ,") + "."


def item_lines(record: PersonRecord, params: OrderParams) -> list[str]:
    """Усі рядки одного пункту, без номера."""
    folder = params.dictionary_folder
    lines = [main_line(record, folder)]
    biography = biography_line(record)
    if biography:
        lines.append(biography)
    if record.ipn:
        lines.append(f"{_clean(record.ipn)}.")
    if record.extra_bio:
        lines.append(_clean(record.extra_bio).rstrip(".") + ".")
    conclusion = appointment_line(record, folder)
    if conclusion:
        lines.append(conclusion)
    return [apply_ukrainian_typography(clean_duplicated_units(line)) for line in lines]


def heading_lines(params: OrderParams) -> list[str]:
    """Шапка розділу: «§ …» окремим рядком і «Відповідно до …:»."""
    lines = []
    if _clean(params.section):
        lines.append(_clean(params.section))
    bases = "".join(f", {_clean(base)}" for base in params.bases if _clean(base))
    who = _clean(params.kind) or DEFAULT_KIND
    unit = _clean(params.unit)
    appoint = "ПРИЗНАЧИТИ"
    if _clean(params.target_unit):
        appoint += f" ДО {_clean(params.target_unit).upper()}"
    text = (
        f"Відповідно до {_clean(params.points)} {REGULATION}{bases} "
        f"нижчепойменованих {who}"
        + (f" {unit}" if unit else "")
        + f" ЗВІЛЬНИТИ з займаних посад і {appoint}:"
    )
    lines.append(apply_ukrainian_typography(clean_duplicated_units(text.strip())))
    return lines


def single_item_lines(record: PersonRecord, params: OrderParams) -> list[str]:
    """Пункт на одну особу: шапка, особа й дія — одним реченням (зразок додатка 53)."""
    folder = params.dictionary_folder
    rank, _ = rank_accusative(str(record.rank), folder)
    full_name = accusative(
        FullName(str(record.surname), str(record.name), str(record.patronymic))
    )
    current, _ = position_accusative(str(record.current.text), folder)
    target, _ = position_instrumental(str(record.target.text), folder)
    bases = "".join(f", {_clean(base)}" for base in params.bases if _clean(base))

    sentence = f"Відповідно до {_clean(params.points)} {REGULATION}{bases} {rank} {full_name}"
    if current:
        sentence += f", {current},"
    sentence += " ЗВІЛЬНИТИ з займаної посади і ПРИЗНАЧИТИ"
    if target:
        sentence += f" {target}"
    vos = _clean(record.target.vos) or _clean(record.current.vos)
    if vos:
        sentence += f", ВОС - {vos}"
    lines = [apply_ukrainian_typography(clean_duplicated_units(sentence.rstrip(" ,") + "."))]
    lines += item_lines(record, params)[1:]
    return lines


def compose_order(records: list[PersonRecord], params: OrderParams | None = None) -> OrderDraft:
    """Записи про осіб → проєкт наказу (шапка, пронумеровані пункти, підстави)."""
    params = params or OrderParams()
    people = list(records)
    if params.sort_by_surname:
        people.sort(key=lambda record: record.sort_key())

    draft = OrderDraft(params=params)
    if not people:
        return draft

    if len(people) == 1 and not _clean(params.section):
        record = people[0]
        draft.items.append(
            OrderItem(
                number=params.numbering_start,
                lines=single_item_lines(record, params),
                record=record,
            )
        )
    else:
        draft.heading = heading_lines(params)
        for offset, record in enumerate(people):
            draft.items.append(
                OrderItem(
                    number=params.numbering_start + offset,
                    lines=item_lines(record, params),
                    record=record,
                )
            )
    draft.footer = [
        apply_ukrainian_typography(f"Підстава: {_clean(base)}")
        for base in params.footer_bases
        if _clean(base)
    ]
    return draft


def numbered_lines(draft: OrderDraft) -> list[str]:
    """Текст наказу рядками разом із номерами пунктів — для показу й для збірки.

    Порожній рядок перед кожним пунктом — правило PROJECT_RULES 5.3; тут воно
    закладається в сам текст, а верстку довершує `build.py`.
    """
    lines: list[str] = []
    lines.extend(draft.heading)
    for item in draft.items:
        if lines:
            lines.append("")
        first, *rest = item.lines
        lines.append(f"{item.number}. {first}")
        lines.extend(rest)
    if draft.footer:
        lines.append("")
        lines.extend(draft.footer)
    return lines
