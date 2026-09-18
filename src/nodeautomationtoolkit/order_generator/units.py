"""Повна назва військової частини з таблиці відповідностей.

В Excel-генераторі назву частини підтягував XLOOKUP із зовнішнього файла на
мережевому диску. Мережевих файлів у нас нема (PROJECT_RULES 1.1), зате є та
сама таблиця, з якої працюють витяги й повідомлення: стовпець A — відкрита
повна назва, B — шифр, C — скорочення.

Користувач пише в реквізитах наказу коротко («А1234», «72 омбр», «окрема
механізована»), а в шапку має стати повна відкрита назва з таблиці, ще й у
родовому («нижчепойменованих осіб офіцерського складу 72 окремої …»).

Знайти нічого не вдалося — не біда: що ввели, те й лишається, а перевірка
скаже, що в таблиці такої частини немає. Мовчки нічого не підставляється.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path

from ..builtin_nodes.recipient_mapping import read_recipient_mapping

#: Слова, які нічого не додають до пошуку («військова частина А1234»).
_NOISE = {"військова", "військової", "військовій", "частина", "частини", "частині", "в/ч", "вч"}
_CIPHER_RE = re.compile(r"^[АA]\s?\d{3,5}$", re.IGNORECASE)


@dataclass(frozen=True)
class UnitMatch:
    """Знайдений рядок таблиці."""

    open_name: str  # стовпець A — повна відкрита назва
    cipher: str  # стовпець B
    abbreviation: str  # стовпець C
    how: str  # як знайшли: «шифр», «скорочення», «назва»

    def __bool__(self) -> bool:
        return bool(self.open_name)


def _key(text: str) -> str:
    value = re.sub(r"[«»\"'’.,]", " ", str(text or "")).casefold()
    words = [word for word in value.split() if word not in _NOISE]
    return " ".join(words)


@lru_cache(maxsize=8)
def _entries(path: str, modified: float) -> tuple[tuple[str, str, str], ...]:
    """Рядки таблиці (назва, шифр, скорочення). Кеш скидається зі зміною файла."""
    result = read_recipient_mapping(path)
    mapping = result["mapping"]
    rows = []
    for open_name, entry in mapping.items():
        rows.append((open_name, entry.get("cipher", ""), entry.get("abbreviation", "")))
    return tuple(rows)


def load_units(table_path: str | Path) -> list[tuple[str, str, str]]:
    """Читає таблицю; порожній або відсутній файл дає порожній список."""
    path = Path(table_path or "")
    if not path.is_file():
        return []
    return list(_entries(str(path), path.stat().st_mtime))


def find_unit(query: str, table_path: str | Path) -> UnitMatch | None:
    """Шукає частину за шифром, скороченням або словами назви.

    Порядок: точний шифр («А1234») → точне скорочення («72 омбр») → назва,
    яка містить усі слова запиту. Якщо підходить кілька назв — береться
    найкоротша: у таблиці довші назви зазвичай є підрозділами.
    """
    rows = load_units(table_path)
    if not rows or not str(query or "").strip():
        return None
    wanted = _key(query)
    compact = wanted.replace(" ", "")

    for open_name, cipher, abbreviation in rows:
        if cipher and _key(cipher).replace(" ", "") == compact:
            return UnitMatch(open_name, cipher, abbreviation, "шифр")
    for open_name, cipher, abbreviation in rows:
        if abbreviation and _key(abbreviation) == wanted:
            return UnitMatch(open_name, cipher, abbreviation, "скорочення")

    words = wanted.split()
    candidates = [
        (open_name, cipher, abbreviation)
        for open_name, cipher, abbreviation in rows
        if words and all(word in _key(open_name) for word in words)
    ]
    if not candidates:
        return None
    open_name, cipher, abbreviation = min(candidates, key=lambda row: len(row[0]))
    return UnitMatch(open_name, cipher, abbreviation, "назва")


#: Закінчення прикметника в називному → родовий.
_ADJECTIVE_GENITIVE = {"ий": "ого", "ій": "ього", "а": "ої", "я": "ьої", "е": "ого"}
#: Роди частин, де родовий не виводиться загальним правилом.
_KIND_GENITIVE = {
    "полк": "полку",
    "батальйон": "батальйону",
    "дивізіон": "дивізіону",
    "корпус": "корпусу",
    "центр": "центру",
    "загін": "загону",
    "взвод": "взводу",
    "склад": "складу",
    "вузол": "вузла",
    "пункт": "пункту",
    "госпіталь": "госпіталю",
    "комплекс": "комплексу",
    "відділ": "відділу",
    "інститут": "інституту",
    "факультет": "факультету",
    "полігон": "полігону",
    "арсенал": "арсеналу",
    "штаб": "штабу",
    "підрозділ": "підрозділу",
}
#: Загальні закінчення іменника: називний → родовий.
_NOUN_GENITIVE = (
    ("ння", "ння"),  # «командування», «управління» — не змінюються
    ("ія", "ії"),
    ("а", "и"),
    ("я", "ї"),
    ("о", "а"),
    ("ь", "я"),
    ("й", "ю"),
)


def _restore_case(source: str, form: str) -> str:
    if source[:1].isupper():
        return form[:1].upper() + form[1:]
    return form


def _adjective_genitive(word: str) -> str:
    low = word.casefold()
    for ending, form in _ADJECTIVE_GENITIVE.items():
        if low.endswith(ending) and len(low) > len(ending) + 1:
            return _restore_case(word, low[: -len(ending)] + form)
    return word


def _noun_genitive(word: str) -> str:
    low = word.casefold()
    if low in _KIND_GENITIVE:
        return _restore_case(word, _KIND_GENITIVE[low])
    for ending, form in _NOUN_GENITIVE:
        if low.endswith(ending):
            return _restore_case(word, low[: -len(ending)] + form)
    return _restore_case(word, low + "а")


def unit_genitive(name: str) -> str:
    """«72 окрема механізована бригада …» → «72 окремої механізованої бригади …».

    Змінюється лише перша група слів — номер, прикметники й сам рід частини.
    Хвіст підпорядкування («оперативного командування «Захід» Сухопутних військ
    Збройних Сил України») у назвах із таблиці вже стоїть у родовому й
    лишається недоторканим.

    Слово вважається прикметником, якщо воно на -ий/-ій або на -а/-я/-е і
    наступне слово теж жіночого чи середнього роду («окрема механізована
    бригада», «Головне управління»).
    """
    words = " ".join(str(name or "").split()).split()
    result: list[str] = []
    for index, word in enumerate(words):
        low = word.casefold()
        if not any(character.isalpha() for character in low):
            result.append(word)  # номер частини, лапки
            continue
        following = words[index + 1].casefold() if index + 1 < len(words) else ""
        adjective = (
            low.endswith(("ий", "ій"))
            # Середній рід: «Головне управління», «Центральне управління».
            or (low.endswith("е") and bool(following))
            # Жіночий рід: прикметник лише перед таким самим словом на -а/-я,
            # інакше це сама назва («рота забезпечення», «бригада зв'язку»).
            or (
                low.endswith(("а", "я"))
                and following.endswith(("а", "я"))
                # «рота забезпечення», «бригада управління»: наступне слово на
                # -ння — це вже залежний іменник, а не другий прикметник.
                and not following.endswith("ння")
            )
        )
        if adjective:
            result.append(_adjective_genitive(word))
            continue
        result.append(_noun_genitive(word))
        result.extend(words[index + 1 :])
        break
    return " ".join(result)


def unit_in_case(name: str, case_label: str = "Р") -> str:
    """Назва частини у відмінку. Поки що вміємо родовий — він і потрібен у шапці."""
    text = " ".join(str(name or "").split())
    if not text or case_label in ("", "Н"):
        return text
    if case_label == "Р":
        return unit_genitive(text)
    return text


def full_unit_name(query: str, table_path: str | Path, case_label: str = "") -> tuple[str, str]:
    """Повна назва частини за коротким записом: (назва, як знайшли).

    Порожній другий елемент означає, що в таблиці такої частини немає і назву
    лишили такою, як її ввели.
    """
    match = find_unit(query, table_path)
    if not match:
        return " ".join(str(query or "").split()), ""
    name = unit_in_case(match.open_name, case_label) if case_label else match.open_name
    return name, match.how
