"""Відмінювання звань і посад за довідниками (перенесено з Excel-генератора).

В Excel посада відмінювалась так: з перших 4–5 слів бралось «ядро» (1 слово, або
2, якщо перше — прикметник зі списку «старший», «головний» …), ядро замінювалось
за довідником, решта лишалась. Тут те саме, але без окремого списку прикметників:

1. Шукається найдовший ПОЧАТОК назви, який є в довіднику цілком
   («старший бойовий медик взводу» → «старший бойовий медик» + «взводу»).
2. Якщо цілої фрази немає — відмінюються слова з початку по одному: слово з
   довідника (форми слів виводяться з фраз довідника), прикметник на -ий/-ій/-а
   перед таким словом — за правилом, складне слово через дефіс — частинами.
   Перше слово, якого немає в довіднику, зупиняє відмінювання: далі йдуть родові
   означення («командира ВЗВОДУ»), вони не змінюються.
3. Якщо не змінилося жодне слово з довідника, назва повертається як є з
   `found=False` — щоб викликач показав це користувачу, а не вгадував.
"""

from __future__ import annotations

from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path

from nodeautomationtoolkit.personnel.dictionaries import (
    CaseDictionary,
    load_case_dictionary,
    normalize_key,
)

RANKS_FILE = "ranks.csv"
POSITIONS_FILE = "positions.csv"
_MAX_PHRASE_WORDS = 6

# Закінчення прикметника в називному → форма у відмінку.
_MASCULINE_ADJECTIVE = {
    "ий": {"Р": "ого", "Д": "ому", "З": "ого", "О": "им", "М": "ому"},
    "ій": {"Р": "ього", "Д": "ьому", "З": "ього", "О": "ім", "М": "ьому"},
}
_FEMININE_ADJECTIVE = {
    "а": {"Р": "ої", "Д": "ій", "З": "у", "О": "ою", "М": "ій"},
}


@dataclass(frozen=True)
class DeclensionResult:
    text: str
    found: bool


def _restore_letter_case(source: str, form: str) -> str:
    letters = [char for char in source if char.isalpha()]
    if letters and all(char.isupper() for char in letters) and len(letters) > 1:
        return form.upper()
    if source[:1].isupper():
        return form[:1].upper() + form[1:]
    return form


def _adjective_form(word_key: str, case: str) -> str | None:
    for ending, forms in (*_MASCULINE_ADJECTIVE.items(), *_FEMININE_ADJECTIVE.items()):
        if word_key.endswith(ending) and len(word_key) > len(ending) + 2 and case in forms:
            return word_key[: -len(ending)] + forms[case]
    return None


def _decline_word(word_key: str, word_forms: dict[str, str]) -> str | None:
    if word_key in word_forms:
        return word_forms[word_key]
    if "-" in word_key:
        parts = word_key.split("-")
        declined = [word_forms.get(part) for part in parts]
        # «офіцер-психолог» при відомому лише «офіцер» дав би «офіцера-психолог».
        # Частину слова не вгадуємо: або відомі всі частини, або жодна.
        if all(declined):
            return "-".join(declined)
    return None


def _is_adjective(word_key: str) -> bool:
    endings = (*_MASCULINE_ADJECTIVE, *_FEMININE_ADJECTIVE)
    return any(word_key.endswith(ending) for ending in endings)


# Іменник жіночого роду на -а після жіночого прикметника («старша медична сестра»):
# форма за правилом, бо в довіднику є лише фрази з дужками, які не перенесено.
_FEMININE_NOUN = {"Р": "и", "Д": "і", "З": "у", "О": "ою", "М": "і"}


def _decline(text: str, dictionary: CaseDictionary, case: str) -> DeclensionResult:
    source = " ".join(str(text or "").split())
    if not source:
        return DeclensionResult("", False)
    words = source.split(" ")
    phrases = dictionary.phrases(case)

    for length in range(min(len(words), _MAX_PHRASE_WORDS), 0, -1):
        head = " ".join(words[:length])
        form = phrases.get(normalize_key(head))
        rest = words[length:]
        if form and rest and _is_adjective(normalize_key(words[length - 1])):
            # У довіднику є «старша медична» без іменника; з іменником після неї
            # виходило б «старшу медичну сестра» — відмінюємо по словах.
            continue
        if form:
            return DeclensionResult(" ".join([_restore_letter_case(head, form), *rest]), True)

    word_forms = dictionary.word_forms(case)
    declined: list[str] = []
    pending_adjectives: list[str] = []
    noun_found = False
    for index, word in enumerate(words):
        key = normalize_key(word)
        form = _decline_word(key, word_forms)
        has_next = index + 1 < len(words)
        if form and _is_adjective(key) and has_next:
            # «старший» теж є в довіднику, але це прикметник: без іменника після
            # нього виходило б «старшого штурман».
            pending_adjectives.append(_restore_letter_case(word, form))
            continue
        if (
            not form
            and pending_adjectives
            and all(
                normalize_key(adjective).endswith(_FEMININE_ADJECTIVE["а"][case])
                for adjective in pending_adjectives
            )
            and key.endswith("а")
            and case in _FEMININE_NOUN
        ):
            form = key[:-1] + _FEMININE_NOUN[case]
        if form:
            declined.extend(pending_adjectives)
            pending_adjectives = []
            declined.append(_restore_letter_case(word, form))
            noun_found = True
            continue
        adjective = _adjective_form(key, case)
        if adjective and has_next:
            pending_adjectives.append(_restore_letter_case(word, adjective))
            continue
        break

    if not noun_found:
        return DeclensionResult(source, False)
    # Прикметники без іменника після них (кінець відмінюваної частини) не змінюються.
    consumed = len(declined)
    return DeclensionResult(" ".join([*declined, *words[consumed:]]), True)


@lru_cache(maxsize=8)
def _dictionary(name: str, user_directory: str) -> CaseDictionary:
    return load_case_dictionary(name, user_directory or None)


def decline_rank(
    rank: str, case: str = "З", user_directory: str | Path | None = None
) -> DeclensionResult:
    """Військове звання у відмінку: «капітан медичної служби» → «капітана медичної служби»."""
    return _decline(rank, _dictionary(RANKS_FILE, str(user_directory or "")), case)


def decline_position(
    position: str, case: str = "З", user_directory: str | Path | None = None
) -> DeclensionResult:
    """Посада у відмінку: «командир взводу» → «командира взводу» (З), «командиром взводу» (О)."""
    return _decline(position, _dictionary(POSITIONS_FILE, str(user_directory or "")), case)


def clear_dictionary_cache() -> None:
    """Скидає прочитані довідники — після того, як користувач змінив файл."""
    _dictionary.cache_clear()
