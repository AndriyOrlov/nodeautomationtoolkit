"""Поділ знайденого тексту на посаду і частину.

«командира механізованої роти військової частини А1111» →
(«командира механізованої роти», «військової частини А1111»).

Межа — найперша з ознак: запис зі списку `unit_markers.txt` (його править
користувач), номер частини або слово з великої літери в тексті звичайного
регістру. Якщо від посади лишилося саме «командира»/«начальника», до неї
додається вид частини: «командира 72 окремої бригади» → «командира бригади».
"""

from __future__ import annotations

import re
from functools import lru_cache
from pathlib import Path

MARKERS_FILE = Path(__file__).with_name("unit_markers.txt")

_UP = "А-ЯІЇЄҐ"
_LO = "а-яіїєґ"
_APOS = "'’ʼ`"

# Номер частини: «72 окремої», «А1111», «__ бригади», «00 ОКРЕМОГО».
# Не номер: «2 категорії», «(на 200 ліжок)», «(S-4)».
_NUMBER_RE = re.compile(
    rf"(?<![\w{_APOS}(/-])(?:[АAВB]\s?\d{{4}}\b|\d{{1,4}}(?:-?[{_LO}{_UP}]{{1,3}})?(?=\s)|_{{2,}})"
)
_NOT_UNIT_AFTER_NUMBER_RE = re.compile(
    rf"\s*(?:категорі|розряд|клас|ліжок|тарифн|рок|р\.|ступен)", re.IGNORECASE
)
_CAPITAL_WORD_RE = re.compile(rf"(?<=\s)[{_UP}][{_LO}{_APOS}-]")

#: Посада без підрозділу — треба додати вид частини.
_LEADERS = re.compile(
    rf"^(?:(?:перш|старш|тимчасово)[{_LO}]*\s+)?(?:(?:командир|заступник|начальник|помічник)[{_LO}]*\s*)+$",
    re.IGNORECASE,
)
_L = f"[{_LO}]"
_UNIT_KIND_RE = re.compile(
    rf"(?<![\w/])(?:військов{_L}*\s+частин{_L}*|в/ч|бригад(?:а|и|і|у|ою)|полк(?:у|ом|і|ові)?"
    rf"|батальйон(?:у|і|ом|ові)?|дивізіон(?:у|і|ом|ові)?|загон(?:у|і|ом)|загін|центр(?:у|і|ом|ові)?"
    rf"|баз(?:а|и|і|у|ою)|госпітал(?:ь|ю|і|ем)|корпус(?:у|і|ом)?|вуз(?:ол|ла|лу|лі|лом)"
    rf"|арсенал(?:у|і|ом)?|академі(?:я|ї|ю|єю)|університет(?:у|і|ом)?|інститут(?:у|і|ом)?"
    rf"|комендатур(?:а|и|і|у|ою)|станці(?:я|ї|ю|єю)|майстерн(?:я|і|ю|ею))(?![\w])",
    re.IGNORECASE,
)
#: Номер перед підрозділом усередині частини («1 механізованого батальйону», «2 роти»)
#: не відділяє частину, а просто прибирається з назви посади.
_SUBDIVISION_AFTER_NUMBER_RE = re.compile(
    rf"\s+(?:(?!окрем)[{_LO}{_APOS}-]+\s+){{0,3}}?(?:рот(?:а|и|і|у|ою)|взвод(?:у|і|ом)?|батаре(?:я|ї|ю|єю)"
    rf"|відділенн(?:я|і|ю|ям)|груп(?:а|и|і|у|ою)|батальйон(?:у|і|ом)?|дивізіон(?:у|і|ом)?"
    rf"|ескадрил(?:ья|ьї|ью)|ланк(?:а|и|і|у|ою)|пункт(?:у|і|ом)?|відділ(?:у|і|ом)?|служб(?:а|и|і|у|ою)"
    rf"|розрахунк(?:у|і|ом)|екіпаж(?:у|і|ем)?|позиці(?:я|ї|ю))(?![\w])",
    re.IGNORECASE,
)


@lru_cache(maxsize=4)
def _markers(path: str, _mtime: float) -> tuple[re.Pattern, ...]:
    patterns = []
    for line in Path(path).read_text(encoding="utf-8").splitlines():
        phrase = line.strip().casefold()
        if not phrase or phrase.startswith("#"):
            continue
        words = [re.escape(word) for word in phrase.split()]
        patterns.append(re.compile(r"(?<![\w/])" + r"\s+".join(words)))
    return tuple(patterns)


def load_markers(path: Path = MARKERS_FILE) -> tuple[re.Pattern, ...]:
    if not path.exists():
        return ()
    return _markers(str(path), path.stat().st_mtime)


def _paren_depth(text: str, end: int) -> int:
    return text.count("(", 0, end) - text.count(")", 0, end)


def _is_subdivision_number(text: str, match: re.Match) -> bool:
    if not match.group(0)[0].isdigit() or _NOT_UNIT_AFTER_NUMBER_RE.match(text, match.end()):
        return False
    return bool(_SUBDIVISION_AFTER_NUMBER_RE.match(text, match.end()))


def _unit_start(text: str, markers: tuple[re.Pattern, ...]) -> int:
    lower = text.lower()  # для кирилиці довжина не змінюється
    candidates = []
    for pattern in markers:
        match = pattern.search(lower, 1)
        if match:
            candidates.append(match.start())
    for match in _NUMBER_RE.finditer(text):
        if match.start() == 0 or _paren_depth(text, match.start()):
            continue
        if _NOT_UNIT_AFTER_NUMBER_RE.match(text, match.end()) or _is_subdivision_number(text, match):
            continue
        candidates.append(match.start())
        break
    if re.search(f"[{_LO}]", text):
        for match in _CAPITAL_WORD_RE.finditer(text):
            if not _paren_depth(text, match.start()) and text[match.start() - 2 : match.start()] != "« ":
                candidates.append(match.start())
                break
    return min(candidates) if candidates else len(text)


def split_position_unit(text: str, markers: tuple[re.Pattern, ...] | None = None) -> tuple[str, str]:
    if markers is None:
        markers = load_markers()
    start = _unit_start(text, markers)
    position = text[:start]
    # Номери підрозділів («1 механізованої роти») однакову посаду не розділяють.
    position = _NUMBER_RE.sub(lambda m: "" if _is_subdivision_number(position, m) else m.group(0), position)
    position = re.sub(r"\s+", " ", position).strip(" ,;:-–—")
    unit = text[start:].strip(" ,;:-–—")
    if not position:
        return text.strip(), ""
    if unit and _LEADERS.match(position):
        kind = _UNIT_KIND_RE.search(unit)
        if kind:
            position = f"{position} {kind.group(0)}"
    return position, unit
