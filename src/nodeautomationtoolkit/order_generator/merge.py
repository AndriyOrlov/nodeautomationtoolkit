"""Поєднання джерел в один запис про особу.

Джерела приходять у будь-якому наборі: лише план; план і пункт попереднього
наказу; лише наказ і ручний ввід; лише ручний ввід. Правила:

- **ручний ввід перемагає завжди** — це рішення користувача;
- **біографія (рік народження, освіта, у ЗС, РНОКПП)** береться з наказу:
  там вона вже вивірена; з плану — лише те, чого в наказі немає;
- **посада «з якої»** — це посада, НА ЯКУ особу призначили попереднім наказом;
  якщо наказу немає, береться займана посада з плану;
- **посада «на яку»** — з плану (або з ручного вводу);
- **звання, підстава, оцінювання** — з плану, бо вони свіжіші за наказ.

Розбіжності не «виправляються» мовчки: якщо план каже, що людина обіймає одну
посаду, а в наказі вона інша, це лишається в `problems` для перевірки.
"""

from __future__ import annotations

from .record import DOCUMENT, MANUAL, ORDER, PLAN, PersonRecord, Position, Value

#: Пріоритет джерел для кожного поля: перше знайдене непорожнє й перемагає.
_FIELD_PRIORITY = {
    "rank": (MANUAL, PLAN, DOCUMENT, ORDER),
    # ПІБ у пункті наказу стоїть у знахідному («ІВАНЕНКА Олексія Вікторовича»),
    # а генератор відмінює сам і чекає називний. Тому ПІБ беремо з плану чи
    # документа, а з наказу — лише коли більше нізвідки (тоді про це пишемо).
    "surname": (MANUAL, PLAN, DOCUMENT, ORDER),
    "name": (MANUAL, PLAN, DOCUMENT, ORDER),
    "patronymic": (MANUAL, PLAN, DOCUMENT, ORDER),
    "ipn": (MANUAL, ORDER, PLAN, DOCUMENT),
    "birth": (MANUAL, ORDER, PLAN, DOCUMENT),
    "education": (MANUAL, ORDER, PLAN, DOCUMENT),
    "service_since": (MANUAL, ORDER, PLAN, DOCUMENT),
    "extra_bio": (MANUAL, ORDER, PLAN, DOCUMENT),
    "basis": (MANUAL, DOCUMENT, PLAN, ORDER),
    "evaluation": (MANUAL, PLAN, DOCUMENT, ORDER),
    "previous_item": (ORDER,),
}


def _pick(field_name: str, records: dict[str, PersonRecord]) -> Value:
    for source in _FIELD_PRIORITY[field_name]:
        record = records.get(source)
        value = getattr(record, field_name, None) if record else None
        if value:
            return value
    return Value()


def _merge_position(primary: Position | None, fallback: Position | None) -> Position:
    """Реквізити (шпк, ВОС, розряд) добираються з запасного джерела, якщо їх немає."""
    primary = primary or Position()
    fallback = fallback or Position()
    if not primary.text:
        return Position(**{name: getattr(fallback, name) for name in ("text", "shpk", "vos", "tariff")})
    merged = Position(text=primary.text)
    for name in ("shpk", "vos", "tariff"):
        merged_value = getattr(primary, name) or getattr(fallback, name)
        setattr(merged, name, merged_value)
    return merged


def build_record(
    plan: PersonRecord | None = None,
    order: PersonRecord | None = None,
    manual: PersonRecord | None = None,
    document: PersonRecord | None = None,
) -> PersonRecord:
    """Один запис про особу з тих джерел, які є (будь-який їх набір)."""
    records = {PLAN: plan, ORDER: order, MANUAL: manual, DOCUMENT: document}
    record = PersonRecord(**{name: _pick(name, records) for name in _FIELD_PRIORITY})

    manual_current = manual.current if manual and manual.current else None
    order_current = order.current if order and order.current else None
    plan_current = plan.current if plan and plan.current else None
    document_current = document.current if document and document.current else None
    # Посада «з якої»: ручна → з наказу (та, на яку призначали) → з плану → з документа.
    record.current = _merge_position(
        manual_current or order_current, plan_current or document_current
    )

    manual_target = manual.target if manual and manual.target else None
    plan_target = plan.target if plan and plan.target else None
    document_target = document.target if document and document.target else None
    record.target = _merge_position(manual_target or plan_target or document_target, None)

    for source in (plan, order, manual, document):
        if source:
            record.problems.extend(source.problems)
            record.notes.extend(source.notes)

    _check(record, plan or document, order)
    return record


def _check(record: PersonRecord, plan: PersonRecord | None, order: PersonRecord | None) -> None:
    if record.surname.source == ORDER:
        record.problems.append(
            "ПІБ узято з пункту наказу, а там воно у знахідному відмінку "
            f"(«{record.full_name}») — звірте називний"
        )
    if not record.target:
        record.problems.append("немає посади, на яку призначається")
    if not record.current:
        record.problems.append("немає посади, з якої призначається")
    if not record.ipn:
        record.problems.append("немає РНОКПП")
    if not record.birth:
        record.problems.append("немає року народження")
    if plan and order and plan.current and order.current:
        if not _same_position(str(plan.current.text), str(order.current.text)):
            record.problems.append(
                "посада в плані не збігається з посадою з наказу: "
                f"план — «{plan.current.text}», наказ — «{order.current.text}»"
            )
    if order and plan and order.ipn and plan.ipn and str(order.ipn) != str(plan.ipn):
        record.problems.append("РНОКПП у плані й у наказі різні")


def _same_position(first: str, second: str) -> bool:
    from ..order_index.position_dictionary import stem

    def key(text: str) -> tuple[str, ...]:
        words = [word for word in text.casefold().replace("-", " ").split() if word.isalpha()]
        return tuple(stem(word) for word in words)

    left, right = key(first), key(second)
    if not left or not right:
        return True
    shorter, longer = sorted((left, right), key=len)
    return shorter == longer[: len(shorter)] or set(shorter) <= set(longer)
