"""Перевірка проєкту наказу перед збіркою (етап «перевірка»).

Нічого не виправляє мовчки: кожна знахідка — це рядок для людини, написаний
простою мовою (що не так, у кого, що зробити). Виправляє тільки людина, бо
наказ підписує вона.

Рівні:
- `ERROR` — так наказ друкувати не можна (немає посади, збився номер, у двох
  пунктів однакова особа);
- `WARNING` — варто глянути (РНОКПП не сходиться, посади немає в довіднику,
  тож відмінок може бути хибним, немає освіти чи року народження).
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from ..personnel import declension
from ..personnel.ipn import check_ipn
from .compose import OrderDraft, OrderItem

ERROR = "помилка"
WARNING = "увага"


@dataclass
class Problem:
    """Одна знахідка перевірки."""

    level: str
    where: str  # «Пункт 3» або «Наказ»
    person: str
    what: str  # що не так
    how: str = ""  # що зробити

    def line(self) -> str:
        mark = "✖" if self.level == ERROR else "⚠"
        who = f" — {self.person}" if self.person else ""
        tail = f" {self.how}" if self.how else ""
        return f"{mark} {self.where}{who}: {self.what}.{tail}"


@dataclass
class CheckResult:
    problems: list[Problem] = field(default_factory=list)

    @property
    def errors(self) -> list[Problem]:
        return [problem for problem in self.problems if problem.level == ERROR]

    @property
    def warnings(self) -> list[Problem]:
        return [problem for problem in self.problems if problem.level == WARNING]

    @property
    def ready(self) -> bool:
        """Чи можна збирати документ: попередження не заважають, помилки — так."""
        return not self.errors

    def lines(self) -> list[str]:
        return [problem.line() for problem in self.problems]


def _add(result: CheckResult, level: str, where: str, person: str, what: str, how: str = "") -> None:
    result.problems.append(Problem(level, where, person, what, how))


def _check_person(result: CheckResult, item: OrderItem, folder: str) -> None:
    record = item.record
    where = f"Пункт {item.number}"
    person = record.full_name or "особа без ПІБ"

    if not record.surname or not record.name or not record.patronymic:
        _add(result, ERROR, where, person, "неповне ПІБ", "Допишіть прізвище, ім'я та по батькові.")
    if not record.rank:
        _add(result, ERROR, where, person, "немає військового звання", "Впишіть звання.")
    if not record.current.text:
        _add(
            result, ERROR, where, person,
            "немає посади, з якої призначається",
            "Візьміть її з попереднього наказу або впишіть вручну.",
        )
    if not record.target.text:
        _add(
            result, ERROR, where, person,
            "немає посади, на яку призначається",
            "Візьміть її з плану переміщення або впишіть вручну.",
        )

    check = check_ipn(str(record.ipn))
    if check.kind == "missing":
        _add(result, WARNING, where, person, "немає РНОКПП", "Без нього пункт неповний.")
    elif check.kind == "invalid" or (check.kind == "ipn" and not check.valid):
        _add(
            result, WARNING, where, person,
            f"РНОКПП «{record.ipn}» не сходиться за контрольною цифрою",
            "Звірте з обліковими документами: найчастіше це описка в одній цифрі.",
        )
    elif check.kind == "ipn" and check.birth_date and record.birth:
        year = re.search(r"(?:19|20)\d{2}", str(record.birth))
        if year and int(year.group(0)) != check.birth_date.year:
            _add(
                result, WARNING, where, person,
                f"рік народження {year.group(0)} не збігається з роком у РНОКПП "
                f"({check.birth_date.year})",
                "Звірте рік народження й номер.",
            )

    if not record.birth:
        _add(result, WARNING, where, person, "немає року народження", "Додайте «19__ р.н.».")
    if not record.education:
        _add(result, WARNING, where, person, "немає освіти", "Додайте рядок «освіта: …».")
    if not record.service_since:
        _add(result, WARNING, where, person, "немає дати вступу на службу", "Додайте «у ЗС - із __.____».")
    if not record.target.vos and not record.current.vos:
        _add(result, WARNING, where, person, "немає ВОС", "Додайте код ВОС нової посади.")

    for text, title in ((str(record.current.text), "посаду, з якої"), (str(record.target.text), "посаду, на яку")):
        if not text:
            continue
        case = "З" if title.endswith("з якої") else "О"
        if not declension.decline_position(text, case, folder or None).found:
            _add(
                result, WARNING, where, person,
                f"{title} призначається, не знайдено в довіднику посад — відмінок може бути хибним",
                "Допишіть посаду в довідник посад або виправте відмінок у тексті.",
            )

    for problem in record.problems:
        _add(result, WARNING, where, person, problem, "Звірте джерела.")


def _check_text(result: CheckResult, item: OrderItem) -> None:
    where = f"Пункт {item.number}"
    person = item.record.full_name
    for line in item.lines:
        plain = line.replace(" ", " ")
        if "  " in plain:
            _add(result, WARNING, where, person, "у тексті подвійний пробіл", "Приберіть зайвий пробіл.")
        if plain.count("«") != plain.count("»"):
            _add(result, WARNING, where, person, "незакрита лапка", "Перевірте лапки в рядку.")
        if not plain.rstrip().endswith((".", ":", ";")):
            _add(
                result, WARNING, where, person,
                "рядок не закінчується крапкою",
                "Кожен рядок пункту закінчується крапкою.",
            )


def check_draft(draft: OrderDraft) -> CheckResult:
    """Перевіряє весь проєкт наказу: реквізити, кожен пункт, повтори, нумерацію."""
    result = CheckResult()
    folder = draft.params.dictionary_folder

    if not draft.items:
        _add(result, ERROR, "Наказ", "", "немає жодного пункту", "Додайте план переміщення або особу вручну.")
        return result
    if "___" in draft.params.points:
        _add(
            result, WARNING, "Наказ", "",
            "у шапці не вказані пункти Положення",
            "Впишіть їх у полі «Пункти Положення».",
        )
    if not draft.params.number or not draft.params.date:
        _add(
            result, WARNING, "Наказ", "",
            "немає номера або дати наказу",
            "Їх можна дописати перед друком.",
        )

    seen_ipn: dict[str, int] = {}
    seen_name: dict[str, int] = {}
    expected = draft.params.numbering_start
    for item in draft.items:
        if item.number != expected:
            _add(
                result, ERROR, f"Пункт {item.number}", "",
                f"збилася нумерація: очікувався пункт {expected}",
                "Перенумеруйте пункти.",
            )
        expected = item.number + 1

        _check_person(result, item, folder)
        _check_text(result, item)

        ipn = "".join(character for character in str(item.record.ipn) if character.isdigit())
        if ipn:
            if ipn in seen_ipn:
                _add(
                    result, ERROR, f"Пункт {item.number}", item.record.full_name,
                    f"той самий РНОКПП уже є в пункті {seen_ipn[ipn]}",
                    "Одна особа — один пункт.",
                )
            else:
                seen_ipn[ipn] = item.number
        key = item.record.full_name.casefold()
        if key:
            if key in seen_name:
                _add(
                    result, ERROR, f"Пункт {item.number}", item.record.full_name,
                    f"та сама особа вже є в пункті {seen_name[key]}",
                    "Приберіть повтор або звірте однофамільців.",
                )
            else:
                seen_name[key] = item.number
    return result
