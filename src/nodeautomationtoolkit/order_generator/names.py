"""Прізвище, ім'я та по батькові у відмінках для пункту наказу.

Знахідний — для призначення й звільнення:
«КОЛМИК Олексій Вікторович» → «КОЛМИКА Олексія Вікторовича»,
«ЗІНЧЕНКО Світлана Петрівна» → «ЗІНЧЕНКО Світлану Петрівну».

Давальний — для присвоєння звання (зразки додатка 53):
«НАЗАРЕНКО Олександр Васильович» → «НАЗАРЕНКУ Олександру Васильовичу»,
«ДУБ Романна Іванівна» → «ДУБ Романні Іванівні».

Стать — за по батькові (-ич / -на). Прізвище лишається ВЕЛИКИМИ.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

_VOWELS = "аеєиіїоуюя"
_MALE_NAMES = {"петро": "петра", "павло": "павла", "дмитро": "дмитра", "ігор": "ігоря", "лазар": "лазаря",
               "ілля": "іллю", "лев": "лева", "олег": "олега"}


@dataclass(frozen=True)
class FullName:
    surname: str
    name: str
    patronymic: str

    @property
    def feminine(self) -> bool:
        return self.patronymic.casefold().endswith("на")

    def short(self) -> str:
        initials = "".join(f"{part[0]}." for part in (self.name, self.patronymic) if part)
        return f"{self.surname.title()} {initials}".strip()


def split_full_name(text: str) -> FullName:
    parts = " ".join(str(text or "").split()).split(" ")
    parts += [""] * (3 - len(parts))
    return FullName(parts[0], parts[1], " ".join(parts[2:]).strip())


def _restore_case(source: str, form: str) -> str:
    if source.isupper():
        return form.upper()
    if source[:1].isupper():
        return form[:1].upper() + form[1:]
    return form


def _masculine_surname(low: str) -> str:
    if low.endswith(("ський", "цький", "зький")) or (low.endswith("ий") and len(low) > 4):
        return low[:-2] + "ого"
    if low.endswith("ній") and len(low) > 5:
        return low[:-2] + "ього"
    if low.endswith("ець") and len(low) > 5:
        stem = low[:-3]
        return (stem[:-1] + "ль" if stem.endswith("л") else stem) + "ця"
    if low.endswith("ій") or low.endswith("й"):
        return low[:-1] + "я"
    if low.endswith("ь"):
        return low[:-1] + "я"
    if low.endswith("ів"):
        return low[:-2] + "ова"
    if low.endswith("їв"):
        return low[:-2] + "єва"
    if low.endswith("о"):
        return low[:-1] + "а"
    if low.endswith("а"):
        return low[:-1] + "у"
    if low.endswith("я"):
        return low[:-1] + "ю"
    if low[-1:] not in _VOWELS:
        return low + "а"
    return low


def _feminine_surname(low: str) -> str:
    if low.endswith(("ська", "цька", "зька")) or low.endswith(("ова", "єва", "іна", "їна", "ина")):
        return low[:-1] + "у"
    if low.endswith("а") and not low.endswith("енка"):
        return low[:-1] + "у"
    if low.endswith("я"):
        return low[:-1] + "ю"
    return low  # «КОВАЛЬЧУК», «ПЕТРЕНКО» — не змінюються


def _masculine_name(low: str) -> str:
    if low in _MALE_NAMES:
        return _MALE_NAMES[low]
    if low.endswith(("ій", "й", "ь")):
        return low[:-1] + "я"
    if low.endswith("о"):
        return low[:-1] + "а"
    if low.endswith("а"):
        return low[:-1] + "у"
    if low.endswith("я"):
        return low[:-1] + "ю"
    return low + "а"


def _feminine_name(low: str) -> str:
    if low.endswith("а"):
        return low[:-1] + "у"
    if low.endswith("я"):
        return low[:-1] + "ю"
    return low  # «Любов»


def _masculine_patronymic_accusative(low: str) -> str:
    """«Васильович» → «Васильовича», «Ілліч» → «Ілліча»."""
    return low + "а" if low.endswith(("ич", "іч")) else low


def accusative(full_name: FullName) -> str:
    """Знахідний відмінок ПІБ; для чоловіків збігається з родовим."""
    surname, name, patronymic = full_name.surname, full_name.name, full_name.patronymic
    if full_name.feminine:
        forms = (_feminine_surname(surname.casefold()), _feminine_name(name.casefold()),
                 re.sub(r"на$", "ну", patronymic.casefold()))
    else:
        # «Ілліч», «Ілліч» і «Іванович» — по батькові на -ич/-іч.
        forms = (_masculine_surname(surname.casefold()), _masculine_name(name.casefold()),
                 _masculine_patronymic_accusative(patronymic.casefold()))
    sources = (surname, name, patronymic)
    return " ".join(_restore_case(source, form) for source, form in zip(sources, forms, strict=True) if source)

def _masculine_surname_dative(low: str) -> str:
    if low.endswith(("ський", "цький", "зький")) or (low.endswith("ий") and len(low) > 4):
        return low[:-2] + "ому"
    if low.endswith("ій") and len(low) > 4:
        return low[:-2] + "ьому"
    if low.endswith("ь"):
        return low[:-1] + "ю"
    if low.endswith(("й", "о")):
        return low[:-1] + "у"
    if low.endswith("а"):
        return low[:-1] + "і"
    if low.endswith("я"):
        return low[:-1] + "ї"
    if low[-1:] not in _VOWELS:
        return low + "у"
    return low


def _feminine_dative(low: str) -> str:
    if low.endswith(("ська", "цька", "зька")) or low.endswith(("ова", "єва", "іна", "їна", "ина")):
        return low[:-1] + "ій"
    # «Марія» → «Марії», але «Леся» → «Лесі»: подвоєння «ї» лише після голосної.
    if low.endswith(("ія", "їя", "ея", "оя", "ая", "уя")):
        return low[:-1] + "ї"
    if low.endswith("я"):
        return low[:-1] + "і"
    if low.endswith("а"):
        return low[:-1] + "і"
    if low.endswith("ов"):  # «Любов»
        return low + "і"
    return low  # «ДУБ», «КОВАЛЬЧУК» — не змінюються


def _masculine_name_dative(low: str) -> str:
    if low in _MALE_NAMES:  # «Петро» → «Петра» у знахідному, «Петру» — в давальному
        return _MALE_NAMES[low][:-1] + "у" if _MALE_NAMES[low].endswith("а") else _MALE_NAMES[low]
    if low.endswith(("ій", "й")):
        return low[:-1] + "ю"
    if low.endswith("ь"):
        return low[:-1] + "ю"
    if low.endswith("о"):
        return low[:-1] + "у"
    if low.endswith("а"):
        return low[:-1] + "і"
    if low.endswith("я"):
        return low[:-1] + "ї"
    return low + "у"


def _masculine_patronymic_dative(low: str) -> str:
    """«Васильович» → «Васильовичу», «Ілліч» → «Іллічу»."""
    return low + "у" if low.endswith(("ич", "іч")) else low


def dative(full_name: FullName) -> str:
    """Давальний відмінок ПІБ — для наказу про присвоєння звання."""
    surname, name, patronymic = full_name.surname, full_name.name, full_name.patronymic
    if full_name.feminine:
        forms = (
            _feminine_dative(surname.casefold()),
            _feminine_dative(name.casefold()),
            re.sub(r"на$", "ні", patronymic.casefold()),
        )
    else:
        forms = (
            _masculine_surname_dative(surname.casefold()),
            _masculine_name_dative(name.casefold()),
            _masculine_patronymic_dative(patronymic.casefold()),
        )
    sources = (surname, name, patronymic)
    return " ".join(
        _restore_case(source, form) for source, form in zip(sources, forms, strict=True) if source
    )
