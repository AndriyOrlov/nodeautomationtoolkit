"""Ланцюг підрозділів у назві посади.

«командир ударного взводу безпілотних систем роти безпілотних систем командного
пункту безпілотних систем окремого батальйону безпілотних систем полку
безпілотних систем» →

    ударного взводу безпілотних систем
    роти безпілотних систем
    командного пункту безпілотних систем
    окремого батальйону безпілотних систем
    полку безпілотних систем

Ланка починається з виду підрозділу (взвод, рота, пункт, батальйон, полк …)
разом із прикметниками перед ним і триває до наступного виду. Так у індексі
шукається не лише посада, а й кожен підрозділ окремо.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

_LO = "а-яіїєґ"
_APOS = "'’ʼ`"

#: Види підрозділів і частин у будь-якому відмінку — початок нової ланки.
_KINDS = (
    "відділенн", "взвод", "рот", "батаре", "батальйон", "дивізіон", "полк", "бригад", "корпус",
    "командуванн", "загон", "загін", "ескадрил", "ескадр", "екіпаж", "ланк", "груп", "розрахунк",
    "пункт", "відділ", "управлінн", "служб", "штаб", "центр", "баз", "склад", "майстерн",
    "станці", "вуз", "арсенал", "комендатур", "госпітал", "клінік", "поліклінік", "кафедр",
    "факультет", "курс", "лаборатор", "інститут", "університет", "академі", "училищ", "ліце",
    "частин", "дивізі", "ескадрон", "секці", "сектор", "напрям", "апарат", "департамент",
)
_KIND_RE = re.compile(rf"^(?:{'|'.join(_KINDS)})[{_LO}]*$")
#: Лише прикметникові закінчення: «начальника» — іменник посади, а не означення.
_ADJECTIVE_ENDINGS = ("ого", "ього", "ої", "ьої", "их", "іх", "ому", "ьому", "ій", "им", "ім", "ий", "ьої")
_WORD_RE = re.compile(rf"[^\W\d_][\w{_APOS}-]*|\d+|_{{2,}}")
#: Слова посади, схожі на вид підрозділу: «командир роти» — це посада, а не ланка.
_NOT_KIND = ("командир", "командувач", "командирськ")
#: «військової частини», «командного пункту» — вид із двох слів: перше не окрема ланка.
_PAIRED_FIRST = ("військов", "командн", "командно-", "навчальн", "тиловог", "об'єднан")


@dataclass(frozen=True)
class Subdivision:
    text: str  # ланка так, як у наказі
    kind: str  # вид підрозділу («взводу», «роти», «полку»)
    level: int  # 0 — найменший підрозділ, далі вгору за текстом


def _is_kind(word: str) -> bool:
    low = word.casefold()
    return bool(_KIND_RE.match(low)) and not low.startswith(_NOT_KIND)


def _is_adjective(word: str) -> bool:
    low = word.casefold()
    return low.endswith(_ADJECTIVE_ENDINGS) and len(low) > 4 and not _is_kind(low)


def split_chain(text: str) -> list[Subdivision]:
    """Ланцюг підрозділів у назві посади (від найменшого до найбільшого)."""
    tokens = [(match.group(0), match.start(), match.end()) for match in _WORD_RE.finditer(text or "")]
    if not tokens:
        return []
    starts: list[int] = []
    for index, (word, _start, _end) in enumerate(tokens):
        if not _is_kind(word):
            continue
        begin = index
        while begin - 1 >= 0 and (not starts or begin - 1 > starts[-1]):
            previous = tokens[begin - 1][0]
            # номер частини й підрозділу належить до своєї ланки: «5 окремої бригади»
            if previous.isdigit() or previous.startswith("__"):
                begin -= 1
                continue
            if not _is_adjective(previous):
                break
            begin -= 1
            if previous.casefold().startswith(_PAIRED_FIRST):
                break
        if starts and begin <= starts[-1]:
            continue  # той самий вид уже відкрив ланку («військової частини»)
        starts.append(begin)

    chain: list[Subdivision] = []
    for level, begin in enumerate(starts):
        end_token = starts[level + 1] - 1 if level + 1 < len(starts) else len(tokens) - 1
        piece = text[tokens[begin][1]:tokens[end_token][2]].strip(" ,;")
        kind = next((word for word, _s, _e in tokens[begin:end_token + 1] if _is_kind(word)), "")
        if piece:
            chain.append(Subdivision(text=" ".join(piece.split()), kind=kind.casefold(), level=level))
    return chain


def chain_key(text: str) -> str:
    """Ключ групування ланки: без регістру, апострофів і номерів («1 роти» = «роти»)."""
    from .position_dictionary import stem

    words = [word for word in _WORD_RE.findall(text or "") if not word.isdigit() and not word.startswith("__")]
    return " ".join(stem(word) for word in words)
