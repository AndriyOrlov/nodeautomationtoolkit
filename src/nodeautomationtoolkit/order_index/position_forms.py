"""Відмінкові форми назви посади: родовий, давальний, орудний.

Змінюється лише «ядро» назви — прикметники й іменник на початку
(«старший бойовий медик», «стрілець-помічник», «головна медична сестра»,
«диспетчер поїзний»); усе після ядра стоїть у родовому й не змінюється
(«командир ВЗВОДУ», «заступник КОМАНДИРА БАТАЛЬЙОНУ»). Правила — для назв осіб:
чоловічий рід на приголосний, -ець, -ар/-яр, -й, жіночий на -а/-я.

Звірено з формами довідника Excel (`personnel/dictionaries/positions.csv`) —
розбіжності там переважно описки самого довідника.
"""

from __future__ import annotations

import re

CASES = ("Р", "Д", "О")

_MASCULINE_ADJECTIVE = {"ий": ("ого", "ому", "им"), "ій": ("ього", "ьому", "ім")}
_FEMININE_ADJECTIVE = {"а": ("ої", "ій", "ою"), "я": ("ьої", "ій", "ьою")}
_SIBILANTS = ("ж", "ш", "щ", "ч")
#: Іменники на -ар/-яр/-ер із м'якою основою: лікар → лікаря, кухар → кухарем.
_SOFT_R = {"лікар", "кухар", "писар", "секретар", "пекар", "токар", "слюсар", "бондар", "бібліотекар",
           "фельд'єгер", "шахтар", "вівчар", "свинар"}
#: Змінна основа: майстер → майстра, швець → шевця.
_STEM_CHANGE = {"майстер": ("майстр", ("а", "у", "ом")), "швець": ("шевц", ("я", "ю", "ем"))}
#: Перша частина складеного слова, що не змінюється: «штаб-сержант» → «штаб-сержанта».
_FIXED_PREFIXES = {"штаб", "обер", "унтер", "віце"}
#: Прикметники, що самі є назвою посади: після них ядро закінчується.
_ADJECTIVE_NOUNS = {"черговий", "вожатий", "вартовий", "рульовий", "оперуповноважений", "уповноважений",
                    "заряджаючий", "пожежний", "підносчий", "навідний", "стройовий"}
_OBLIQUE_ENDINGS = ("их", "іх", "ого", "ього", "ої", "ьої", "ів", "ам", "ям", "ою", "ею", "ому", "ьому", "ах", "ях")


def _noun(word: str, index: int) -> str:
    """Форма іменника: index 0 — родовий, 1 — давальний, 2 — орудний."""
    low = word
    if low in _STEM_CHANGE:
        stem, endings = _STEM_CHANGE[low]
        return stem + endings[index]
    if low.endswith("ець") and len(low) > 4:
        stem = low[:-3]
        if stem.endswith("л"):
            stem = stem[:-1] + "ль"  # стрілець → стрільця
        return stem + ("ця", "цю", "цем")[index]
    if low.endswith("ій"):
        return low[:-2] + ("ія", "ію", "ієм")[index]
    if low.endswith("й"):
        return low[:-1] + ("я", "ю", "єм")[index]
    if low.endswith("ь"):
        return low[:-1] + ("я", "ю", "ем")[index]
    if low.endswith("а"):
        return low[:-1] + ("и", "і", "ою")[index]
    if low.endswith("я"):
        return low[:-1] + ("і", "і", "ею")[index]
    if low in _SOFT_R:
        return low + ("я", "ю", "ем")[index]
    if low.endswith(_SIBILANTS):
        return low + ("а", "у", "ем")[index]
    if re.search("[бвгґджзклмнпрстфхц]$", low):
        return low + ("а", "у", "ом")[index]
    return low


def _is_adjective(word: str) -> bool:
    return len(word) > 4 and (word.endswith("ий") or word.endswith("ній"))


def _adjective(word: str, index: int, feminine: bool = False) -> str:
    table = _FEMININE_ADJECTIVE if feminine else _MASCULINE_ADJECTIVE
    for ending, forms in table.items():
        if word.endswith(ending) and len(word) > len(ending) + 2:
            return word[: -len(ending)] + forms[index]
    return word


def _feminine_core(words: list[str]) -> int:
    """Скільки слів займає жіноче ядро «головна медична сестра» (0 — немає)."""
    count = 0
    while count < len(words) and words[count].endswith(("на", "ча", "ша", "ова", "ська", "цька", "ла", "ша")):
        count += 1
    if count and count < len(words) and words[count].endswith(("а", "я")) and not words[count].endswith(_OBLIQUE_ENDINGS):
        return count + 1
    if len(words) == 1 and words[0] in {"сестра", "старшина"}:
        return 1
    return 0


def _compound(word: str, index: int) -> str:
    """«стрілець-помічник», «водій-заряджаючий», «машиніст-екскаватора» — частинами."""
    parts = word.split("-")
    result = []
    for position, part in enumerate(parts):
        if position and part.endswith(("а", "я", "у", "ю", "ів")) and not _is_adjective(part):
            result.append(part)  # залежне слово в родовому: «машиніст-екскаватора»
        elif not position and part in _FIXED_PREFIXES and len(parts) > 1:
            result.append(part)
        elif _is_adjective(part):
            result.append(_adjective(part, index))
        else:
            result.append(_noun(part, index))
    return "-".join(result)


def _nominative_looking(word: str) -> bool:
    return bool(word) and word[0].isalpha() and not word.endswith(_OBLIQUE_ENDINGS)


def decline(nominative: str, case: str) -> str:
    text = " ".join(nominative.split()).casefold()
    # «начальник штабу - перший заступник командира» — дві посади, відмінюється кожна
    if " - " in text:
        return " - ".join(_decline_one(part, case) for part in text.split(" - "))
    return _decline_one(text, case)


def _masculine_noun_looking(word: str) -> bool:
    """Іменник у називному чоловічого роду: закінчується на приголосний, -й, -ь."""
    last = word.split("-")[-1]
    return bool(re.search("[бвгґджзйклмнпрстфхцчшщь']$", last))


def _decline_one(text: str, case: str) -> str:
    words = text.split(" ")
    index = CASES.index(case)

    feminine = _feminine_core(words)
    if feminine:
        core = [_adjective(w, index, True) for w in words[: feminine - 1]]
        return " ".join([*core, _noun(words[feminine - 1], index), *words[feminine:]])

    result: list[str] = []
    for position, word in enumerate(words):
        rest = words[position + 1:]
        result.append(_compound(word, index))
        last = word.split("-")[-1]
        next_word = rest[0] if rest else ""
        if _is_adjective(last) and _nominative_looking(next_word) and (
            last not in _ADJECTIVE_NOUNS or _masculine_noun_looking(next_word) or _is_adjective(next_word)
        ):
            continue  # прикметник перед іменником ядра («черговий помічник», але «черговий частини»)
        # ядро закінчилось; прикметник після іменника узгоджується: «диспетчер поїзний»
        tail = 0
        while tail < len(rest) and _is_adjective(rest[tail]) and not _is_adjective(last):
            result.append(_adjective(rest[tail], index))
            tail += 1
        return " ".join([*result, *rest[tail:]])
    return " ".join(result)


def all_forms(nominative: str) -> dict[str, str]:
    return {case: decline(nominative, case) for case in CASES}
