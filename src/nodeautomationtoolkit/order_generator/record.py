"""Спільний запис про особу для генератора наказів.

Джерела бувають різні й у різних поєднаннях: план переміщення, попередній наказ
із індексу, ручний ввід, а згодом — подання, витяг чи розпізнаний скан. Тому
кожне поле зберігає не лише значення, а й джерело, з якого воно взялося:
видно, що взято з наказу, що з плану, а що дописав користувач.

Правило переходу посад (як працює колега): посада, НА ЯКУ особу призначили
попереднім наказом, стає посадою, З ЯКОЇ її призначають тепер; нова посада
приходить із плану (або вводиться вручну).
"""

from __future__ import annotations

from dataclasses import dataclass, field, fields

#: Назви джерел у журналі та в підказках.
PLAN = "план"
ORDER = "наказ"
DOCUMENT = "документ"  # подання, рапорт, аркуш бесіди — будь-який кадровий документ
MANUAL = "вручну"
DICTIONARY = "довідник"


@dataclass
class Value:
    """Значення поля разом із джерелом."""

    text: str = ""
    source: str = ""

    def __bool__(self) -> bool:
        return bool(self.text)

    def __str__(self) -> str:
        return self.text


@dataclass
class Position:
    """Посада з реквізитами, як вона стоїть у наказі чи в плані."""

    text: Value = field(default_factory=Value)  # повна назва з частиною
    shpk: Value = field(default_factory=Value)
    vos: Value = field(default_factory=Value)
    tariff: Value = field(default_factory=Value)

    def __bool__(self) -> bool:
        return bool(self.text)

    def filled(self) -> dict[str, Value]:
        return {item.name: getattr(self, item.name) for item in fields(self) if getattr(self, item.name)}


@dataclass
class PersonRecord:
    """Усе, що потрібно для пункту наказу про одну особу."""

    rank: Value = field(default_factory=Value)
    surname: Value = field(default_factory=Value)
    name: Value = field(default_factory=Value)
    patronymic: Value = field(default_factory=Value)
    ipn: Value = field(default_factory=Value)
    birth: Value = field(default_factory=Value)
    education: Value = field(default_factory=Value)
    service_since: Value = field(default_factory=Value)
    extra_bio: Value = field(default_factory=Value)  # «Кандидат технічних наук.» тощо
    current: Position = field(default_factory=Position)
    target: Position = field(default_factory=Position)
    basis: Value = field(default_factory=Value)
    evaluation: Value = field(default_factory=Value)
    # Звільнення з військової служби (стаття 26 Закону; зразки додатка 53).
    dismissal: Value = field(default_factory=Value)  # ключ підстави: «1.а», «б»
    destination: Value = field(default_factory=Value)  # «у запас» / «у відставку»
    service_calendar: Value = field(default_factory=Value)  # «29 років 3 місяці»
    service_privileged: Value = field(default_factory=Value)  # «29 років 11 місяців» / «немає»
    registration: Value = field(default_factory=Value)  # ТЦК, куди стає на облік
    uniform: Value = field(default_factory=Value)  # «так» / «ні» — право носіння форми
    dismissal_note: Value = field(default_factory=Value)  # «Чинність контракту припиняється …»
    # Присвоєння військового звання (зразки додатка 53).
    new_rank: Value = field(default_factory=Value)  # звання, яке присвоюють
    rank_seniority: Value = field(default_factory=Value)  # «вислуга у званні - 11 років»
    rank_since: Value = field(default_factory=Value)  # «строк рахувати з 04.12.2013»
    rank_note: Value = field(default_factory=Value)  # «достроково на 6 місяців»
    previous_item: Value = field(default_factory=Value)  # текст пункту, з якого переносили
    problems: list[str] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)

    @property
    def full_name(self) -> str:
        return " ".join(str(part) for part in (self.surname, self.name, self.patronymic) if part)

    def sort_key(self) -> tuple[str, str, str]:
        return (str(self.surname).casefold(), str(self.name).casefold(), str(self.patronymic).casefold())

    def sources(self) -> dict[str, str]:
        """Поле → джерело; для журналу й підказок у вікні."""
        result: dict[str, str] = {}
        for item in fields(self):
            value = getattr(self, item.name)
            if isinstance(value, Value) and value:
                result[item.name] = value.source
            elif isinstance(value, Position):
                for sub_name, sub_value in value.filled().items():
                    result[f"{item.name}.{sub_name}"] = sub_value.source
        return result
