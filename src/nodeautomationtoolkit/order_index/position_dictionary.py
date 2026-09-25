"""Розпізнавання посади за довідником — у будь-якому відмінку й регістрі.

Словник один — `order_index/positions_all.csv` (CSV, користувач править сам):
усі посади в називному, родовому, давальному й орудному, з кодами посад і ВОС
за наказом МО № 317. Збирає його `scripts/order_index/build_all_positions.py`
з таблиці Excel-генератора, переліку МО № 317 і посад офіцерського складу;
рядки з джерелом «вручну» при перезбиранні зберігаються.

Кожне слово зводиться до основи: відкидається відмінкове закінчення
(«командира», «командиром», «КОМАНДИРУ» → «командир»). Фраза довідника
шукається як послідовність основ на початку тексту — найдовша перемагає.
Для відомих форм зі зміною основи («стрілець» → «стрільця») беруться форми
зі стовпців довідника.

Загальна посада без підрозділу («командир», «начальник», «старший офіцер»)
доповнюється першим іменником після неї: «командира механізованої роти» →
«командир роти», «заступника командира 72 окремої бригади» →
«заступник командира бригади». Так та сама посада в різних частинах і
підрозділах зводиться в один рядок, а повний текст лишається в згадці.
"""

from __future__ import annotations

import csv
import re
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path

from ..personnel.dictionaries import DEFAULT_DICTIONARY_DIR

ALL_POSITIONS_FILE = Path(__file__).with_name("positions_all.csv")
BASE_POSITIONS_FILE = DEFAULT_DICTIONARY_DIR / "positions.csv"
#: Перелік штатних посад рядового, сержантського і старшинського складу (наказ МО № 317)
#: з кодами посад і ВОС; збирається `scripts/order_index/build_mo317_dictionaries.py`.
MO317_POSITIONS_FILE = DEFAULT_DICTIONARY_DIR / "mo317_positions.csv"
MO317_VOS_FILE = DEFAULT_DICTIONARY_DIR / "mo317_vos.csv"
MO317_NAME_COLUMN = "найменування посади"
#: ВОС осіб офіцерського складу (наказ МО № 444): прив'язані до категорій посад, не до назв;
#: збирає `scripts/order_index/build_mo444_dictionaries.py`.
MO444_OFFICER_VOS_FILE = DEFAULT_DICTIONARY_DIR / "mo444_officer_vos.csv"
MO444_VOS_REPLACEMENT_FILE = DEFAULT_DICTIONARY_DIR / "mo444_vos_replacement.csv"

_MAX_WORDS = 14
_WORD_RE = re.compile(r"[^\W\d_](?:[\w'’ʼ`-]*[^\W\d_])?|\d+|_{2,}|\S")
_ENDINGS = sorted(
    {
        "ього", "ьому", "ого", "ому", "ими", "ієм", "ією", "ієві", "еєю", "єві", "ові", "еві",
        "ія", "ію", "ії", "ій", "еї", "ея", "ею", "єю", "ою", "им", "ім", "ий", "ої", "ів",
        "ом", "ем", "єм", "их", "ах", "ям", "ами", "ями",
        "а", "я", "у", "ю", "і", "и", "ї", "е", "є", "о", "й", "ь",
    },
    key=len,
    reverse=True,
)
_ADJECTIVE_ENDINGS = ("ого", "ього", "ої", "ьої", "ому", "ьому", "ій", "их", "ий", "ім", "им", "ою", "ої")
#: Посади, які без підрозділу нічого не кажуть — до них додається іменник.
_GENERIC_HEADS = {"командир", "начальник", "заступник", "помічник", "офіцер"}
#: Слова перед посадою, які до неї не належать: «колишньому старшому помічнику».
_SKIPPED_PREFIXES = ("колишн", "тимчасово", "виконуюч")
_BRANCH_PREPOSITIONS = {"з", "із", "зі", "по"}
_PREPOSITIONS = {"у", "в", "на", "для", "до", "за"}
#: Скільки іменників можна додати («заступник командира роти» — два).
_MAX_EXTENSION = 3


#: Стовпці словника з формами назви (решта — коди, ВОС, джерело).
_FORM_COLUMNS = {"називний", "родовий", "давальний", "знахідний", "орудний", "місцевий", "інші форми"}
_DASHES = {"-", "–", "—"}


def stem(word: str) -> str:
    if word in _DASHES:
        return "-"
    word = re.sub("[’ʼ`]", "'", word.casefold())
    parts = []
    for part in word.split("-"):
        for ending in _ENDINGS:
            if part.endswith(ending) and len(part) - len(ending) >= 2:
                part = part[: -len(ending)]
                break
        parts.append(part)
    return "-".join(parts)


@dataclass(frozen=True)
class PositionMatch:
    nominative: str
    start: int  # символ, з якого починається посада в тексті
    end: int  # символ, де закінчилась розпізнана частина (з доданими іменниками)


class PositionDictionary:
    def __init__(self, phrases: dict[tuple[str, ...], str]):
        self.phrases = phrases
        self.max_words = max((len(key) for key in phrases), default=1)

    @classmethod
    def from_files(cls, *paths: Path) -> PositionDictionary:
        phrases: dict[tuple[str, ...], str] = {}
        for path in paths:
            if not path.is_file():
                continue
            rows = list(csv.reader(path.read_text(encoding="utf-8-sig").splitlines(), delimiter=";"))
            header = [cell.strip().casefold() for cell in rows[0]] if rows else []
            if MO317_NAME_COLUMN in header:
                # Перелік МО № 317: лише назви в називному; слова в дужках можна не писати (прим. 6).
                name_index = header.index(MO317_NAME_COLUMN)
                for row in rows[1:]:
                    if name_index < len(row) and row[name_index].strip():
                        for variant in mo317_name_variants(row[name_index]):
                            key = tuple(stem(word) for word in variant.split())
                            if 0 < len(key) <= _MAX_WORDS:
                                phrases.setdefault(key, strip_brackets(row[name_index]))
                continue
            form_columns = [index for index, name in enumerate(header) if name in _FORM_COLUMNS]
            for row in rows[1:]:
                if not row or not row[0].strip() or row[0].lstrip().startswith("#"):
                    continue
                nominative = " ".join(row[0].split()).casefold()
                for index in form_columns:
                    if index >= len(row):
                        continue
                    for form in row[index].split("|"):
                        key = tuple(stem(word) for word in form.split())
                        if 0 < len(key) <= _MAX_WORDS:
                            phrases.setdefault(key, nominative)
        return cls(phrases)

    def match_at_start(self, text: str) -> PositionMatch | None:
        """Посада, з якої починається `text` (перед нею можуть бути лише пробіли).

        Два проходи: слова як є і слова без номерів та означень підрозділу
        («командира 2 механізованої роти» → «командира роти»); перемагає довша посада.
        """
        tokens = [(m.group(0), m.start(), m.end()) for m in _WORD_RE.finditer(text)]
        while tokens and tokens[0][0].casefold().startswith(_SKIPPED_PREFIXES):
            tokens = tokens[1:]
        plain = self._match_words(text, tokens, _leading_words(tokens, self.max_words, skip_dependents=False))
        condensed = self._match_words(text, tokens, _leading_words(tokens, self.max_words, skip_dependents=True))
        candidates = [match for match in (plain, condensed) if match]
        if not candidates:
            return None
        return max(candidates, key=lambda match: (len(match.nominative.split()), match is plain))

    def _match_words(self, text: str, tokens, words) -> PositionMatch | None:
        while words and words[-1][0] in _DASHES:
            words.pop()
        for length in range(len(words), 0, -1):
            if words[length - 1][0] in _DASHES:
                continue
            key = tuple(stem(word) for word, _start, _end in words[:length])
            nominative = self.phrases.get(key)
            if nominative and _is_truncated_phrase(nominative) and len(tokens) > length:
                # «командир механізованого» у довіднику Excel — обрізок; далі йде іменник,
                # тож беремо коротшу фразу й доповнюємо її як загальну посаду.
                continue
            if nominative:
                end = words[length - 1][2]
                consumed = next(i for i, token in enumerate(tokens) if token[2] == end) + 1
                nominative, end = self._extend(nominative, text, tokens, consumed, end)
                return PositionMatch(nominative, words[0][1], end)
        return None

    def _extend(self, nominative: str, text: str, tokens, consumed: int, end: int) -> tuple[str, int]:
        added = 0
        index = consumed
        while added < _MAX_EXTENSION and stem(nominative.split()[-1]).split("-")[-1] in {
            stem(head) for head in _GENERIC_HEADS
        }:
            noun = None
            while index < len(tokens):
                word, _start, word_end = tokens[index]
                index += 1
                if word.isdigit() or word.startswith("__"):
                    continue
                if not word[0].isalpha():
                    break  # кома, тире, дужка — посада закінчилась
                lower = word.casefold()
                if lower in _BRANCH_PREPOSITIONS:
                    # «заступник командира з морально-психологічного забезпечення» — окрема посада
                    phrase = [lower]
                    while index < len(tokens) and tokens[index][0][0].isalpha():
                        phrase.append(tokens[index][0].casefold())
                        word_end = tokens[index][2]
                        index += 1
                        if not phrase[-1].endswith(_ADJECTIVE_ENDINGS):
                            break
                    return f"{nominative} {' '.join(phrase)}", word_end
                if lower in _PREPOSITIONS:
                    break
                if lower.endswith(_ADJECTIVE_ENDINGS):
                    continue
                noun = (lower, word_end)
                break
            if noun is None:
                break
            nominative = f"{nominative} {noun[0]}"
            end = noun[1]
            added += 1
        return nominative, end


def strip_brackets(name: str) -> str:
    """Ключ назви посади: «Льотчик (літака)» → «льотчик», «Стрілець - радист» → «стрілець-радист».

    Вкладені й незакриті дужки теж прибираються; апострофи й тире зводяться до одного вигляду.
    """
    text = " ".join(name.split()).casefold()
    previous = None
    while previous != text:
        previous = text
        text = re.sub(r"\s*\([^()]*\)?", "", text)
    text = re.sub(r"\s*[-–—]\s*", "-", re.sub("[’ʼ`]", "'", text))
    return " ".join(text.replace(")", " ").split()).strip(" ,-")


def mo317_name_variants(name: str) -> set[str]:
    """Назва без дужок і з їхнім вмістом без самих дужок («льотчик літака»)."""
    clean = " ".join(name.split()).casefold()
    return {strip_brackets(clean), " ".join(re.sub(r"[()]", " ", clean).split())} - {""}


@dataclass
class Mo317Info:
    codes: set[str]  # коди посад
    vos: set[str]  # номери ВОС, для яких посада передбачена
    all_vos: bool  # «Усі ВОС» хоча б в одному варіанті
    ranks: set[str]
    rows: list[dict[str, str]]  # рядки переліку як є (варіанти а/б/в)


def load_mo317(path: Path = MO317_POSITIONS_FILE) -> dict[str, Mo317Info]:
    """Назва посади без дужок (малими) → коди посад, ВОС, штатні звання з наказу № 317."""
    if not path.is_file():
        return {}
    signature = (str(path), path.stat().st_mtime)
    return _load_mo317_cached(signature)


@lru_cache(maxsize=2)
def _load_mo317_cached(signature: tuple[str, float]) -> dict[str, Mo317Info]:
    result: dict[str, Mo317Info] = {}
    with open(signature[0], encoding="utf-8-sig", newline="") as handle:
        for row in csv.DictReader(handle, delimiter=";"):
            key = strip_brackets(row.get("Найменування посади", ""))
            if not key:
                continue
            info = result.setdefault(key, Mo317Info(set(), set(), False, set(), []))
            info.codes.add(row.get("Код посади", ""))
            info.vos.update(v for v in row.get("Номери ВОС", "").split(", ") if v)
            info.all_vos = info.all_vos or bool(row.get("Усі ВОС"))
            if row.get("Штатне звання"):
                info.ranks.add(row["Штатне звання"])
            info.rows.append(row)
    return result


def load_officer_vos(path: Path = MO444_OFFICER_VOS_FILE) -> dict[str, dict[str, str]]:
    """Код ВОС офіцерського складу (6 цифр) → рядок переліку наказу МО № 444."""
    if not path.is_file():
        return {}
    with open(path, encoding="utf-8-sig", newline="") as handle:
        return {row["Код ВОС"]: row for row in csv.DictReader(handle, delimiter=";")}


def load_vos_names(path: Path = MO317_VOS_FILE) -> dict[str, str]:
    if not path.is_file():
        return {}
    with open(path, encoding="utf-8-sig", newline="") as handle:
        return {row["Номер ВОС"]: row["Найменування ВОС"] for row in csv.DictReader(handle, delimiter=";")}


def _leading_words(tokens, limit: int, skip_dependents: bool) -> list[tuple[str, int, int]]:
    """Слова з початку тексту до першого розділового знака.

    `skip_dependents`: пропустити номери й прикметники-означення, що стоять ПІСЛЯ
    іменника («командира 2 механізованої роти»), але не перед ним («старшого офіцера»).
    """
    words: list[tuple[str, int, int]] = []
    for token in tokens:
        if len(words) >= limit:
            break
        word = token[0]
        if skip_dependents and (word.isdigit() or word.startswith("__")):
            continue
        if not (word[0].isalpha() or word in _DASHES):
            break  # тире всередині лишається: «начальник штабу - перший заступник командира»
        if (
            skip_dependents
            and words
            and words[-1][0] not in _DASHES
            and word.casefold().endswith(_ADJECTIVE_ENDINGS)
            and not words[-1][0].casefold().endswith(_ADJECTIVE_ENDINGS)
        ):
            continue
        words.append(token)
    return words


def _is_truncated_phrase(nominative: str) -> bool:
    words = nominative.split()
    return len(words) > 1 and words[-1].endswith(("ого", "ього"))


@lru_cache(maxsize=4)
def _cached(signature: tuple[tuple[str, float], ...]) -> PositionDictionary:
    return PositionDictionary.from_files(*(Path(path) for path, _mtime in signature))


def load_position_dictionary(*paths: Path) -> PositionDictionary:
    """Словник посад; якщо `positions_all.csv` ще не зібрано — вихідні довідники напряму."""
    if not paths:
        paths = (ALL_POSITIONS_FILE,) if ALL_POSITIONS_FILE.is_file() else (BASE_POSITIONS_FILE, MO317_POSITIONS_FILE)
    signature = tuple((str(path), path.stat().st_mtime if path.is_file() else 0.0) for path in paths)
    return _cached(signature)
