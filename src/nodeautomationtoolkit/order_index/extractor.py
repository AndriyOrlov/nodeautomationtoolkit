"""Витяг посад із тексту наказу по о/с (без Word, без мережі).

Пункт наказу за додатком 53 має вигляд:

    Капітана ПЕТРЕНКА Петра Петровича, командира роти … - КОМАНДИРОМ БАТАЛЬЙОНУ …, ВОС - 0211003.

- **займана посада** — текст після «ПРІЗВИЩЕ Ім'я По батькові,» до ВЕЛИКИХ
  літер, «, ВОС» або кінця речення (родовий, у контрактах — орудний);
- **посада призначення** — перша група слів ВЕЛИКИМИ після займаної, яка не
  є дієсловом наказу (ЗВІЛЬНИТИ, ЗАРАХУВАТИ …); після «ПРИЗНАЧИТИ» береться
  те, що стоїть за ним («ПРИЗНАЧИТИ ЗАСТУПНИКОМ КОМАНДИРА …»).

Особисті дані (ПІБ, ІПН) не зберігаються: індексуються лише посади.
"""

from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree

ROLE_CURRENT = "займана"
ROLE_TARGET = "призначення"

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

_UP = "А-ЯІЇЄҐ"
_LO = "а-яіїєґ"
_APOS = "'’ʼ`"

# «ПРІЗВИЩЕ Ім'я По-батькові,» — прізвище лише великими, ім'я й по батькові з великої.
_PERSON_RE = re.compile(
    rf"(?<![\w{_APOS}-])[{_UP}][{_UP}{_APOS}-]+\s+"
    rf"[{_UP}][{_LO}{_APOS}-]+\s+"
    rf"[{_UP}][{_LO}{_APOS}-]+\s*,\s*"
)
# Слово з п'яти й більше великих літер: початок ВЕЛИКОЇ групи.
_CAPS_WORD_RE = re.compile(rf"(?<![\w{_APOS}-])[{_UP}][{_UP}{_APOS}-]{{4,}}(?![\w{_APOS}-])")
_HAS_LOWER_RE = re.compile(r"[a-zа-яіїєґ]")
_VOS_RE = re.compile(r",?\s*ВОС\b")
_SENTENCE_END_RE = re.compile(r"[.;](?=\s|$)")
_SECTION_RE = re.compile(r"^\s*§\s*\d+")
# «ПІБ, 1946 р.н., посада» — біографія одразу після ПІБ посадою не є.
_LEADING_BIO_RE = re.compile(
    r"^(?:(?:\d{2}\.\d{2}\.)?\d{4}\s*(?:р\.?\s*н\.?|року\s+народження)|РНОКПП[\s:-]*\d+|\d{10})\s*[,.;]?\s*",
    re.IGNORECASE,
)

#: Скорочення, після яких крапка не завершує речення («м. Києва», «ім. Івана»).
_ABBREVIATIONS = {"м", "ім", "вул", "р", "с", "смт", "обл", "т", "д", "св", "ген", "півн", "півд", "зах", "сх"}

#: Дієслова наказу: самі по собі посадою не є.
ORDER_VERBS = {
    "ЗВІЛЬНИТИ", "ПРИЗНАЧИТИ", "ЗАРАХУВАТИ", "ВИКЛЮЧИТИ", "ПЕРЕВЕСТИ", "НАДАТИ",
    "ВВАЖАТИ", "ПРИЙНЯТИ", "УКЛАСТИ", "ПРОДОВЖИТИ", "ПРИСВОЇТИ", "ПОЗБАВИТИ",
    "ПОНИЗИТИ", "ВІДКЛИКАТИ", "НАПРАВИТИ", "ПОВЕРНУТИ", "ДОПУСТИТИ", "ВІДНОВИТИ",
    "СКАСУВАТИ", "ВНЕСТИ", "ПОЛАГАТИ", "ЗАЛИШИТИ", "ВІДРАХУВАТИ", "ТИМЧАСОВО",
}
#: «ПРИЗНАЧИТИ ДО ВІЙСЬКОВОГО ІНСТИТУТУ», «У РОЗПОРЯДЖЕННЯ …» — це частина, а не посада.
_NOT_POSITION_START = {"ДО", "У", "В", "НА", "ІЗ", "З", "ЗА", "НАЧАЛЬНИКУ"}

_INSTRUMENTAL_ENDINGS = ("ом", "ем", "єм", "им", "ім", "ою", "ею", "єю")
_DATIVE_ENDINGS = ("ому", "ьому", "у", "ю", "ові", "еві", "єві")
_GENITIVE_ENDINGS = ("а", "я", "ого", "ього", "ої")
#: Посада призначення починається з орудного відмінка: КОМАНДИРОМ, СТАРШИМ, ВОДІЄМ.
_TARGET_HEAD_RE = re.compile(r"(?:ОМ|ЕМ|ЄМ|ИМ|ІМ)(?:-[А-ЯІЇЄҐ'’ʼ]+)?$")
_NOT_TARGET_HEADS = {"ТАКИМ", "ЯКИМ", "НИМ", "ВИКЛЮЧЕННЯМ", "ЗАПАСОМ"}
# Займана посада закінчується перед підрядним зворотом: «, який …», «, надалі …».
_CURRENT_TAIL_RE = re.compile(
    rf",\s+(?:як(?:ий|а|е|і)|надалі|із?\s+\d|у\s+порядку|у\s+по|звільнен|призначен|зарахован)(?![{_LO}])"
)


#: М'який перенос, нульовий пробіл, BOM; нерозривний пробіл окремо.
_INVISIBLE = (chr(0xAD), chr(0x200B), chr(0xFEFF))
_NBSP = chr(0xA0)


@dataclass(frozen=True)
class PositionHit:
    text: str  # повний текст із наказу: посада + підрозділ + частина
    role: str
    section: str
    paragraph: int
    position: str = ""  # назва для групування: з довідника в називному або вирізана евристикою
    known: bool = False  # чи знайдена посада в довіднику
    position_end: int = 0  # де в `text` закінчилась назва посади (далі — підрозділи й частина)


def normalize_space(text: str) -> str:
    for invisible in _INVISIBLE:
        text = (text or "").replace(invisible, "")
    return re.sub(r"[\s" + _NBSP + "]+", " ", text or "").strip()


def position_key(text: str) -> str:
    """Ключ групування: регістр, лапки, апострофи й тире не розрізняються."""
    key = normalize_space(text).casefold()
    key = re.sub(r"[«»„“”\"]", '"', key)
    key = re.sub(f"[{_APOS}]", "'", key)
    key = re.sub(r"\s*[–—-]\s*", "-", key)
    return key.strip(" ,.;:-")


def guess_case(text: str) -> str:
    first = normalize_space(text).split(" ", 1)[0].casefold() if text else ""
    first = first.split("-", 1)[0]
    if first.endswith(_INSTRUMENTAL_ENDINGS):
        return "орудний"
    if first.endswith(_DATIVE_ENDINGS):
        return "давальний"
    if first.endswith(_GENITIVE_ENDINGS):
        return "родовий"
    return "—"


# ─────────────────────────────────────────────────────────────────────────────
# Метадані наказу — з назви файлу (PROJECT_RULES 3.1), запасно — із шапки.
# ─────────────────────────────────────────────────────────────────────────────
_MONTHS = {
    "січня": 1, "лютого": 2, "березня": 3, "квітня": 4, "травня": 5, "червня": 6,
    "липня": 7, "серпня": 8, "вересня": 9, "жовтня": 10, "листопада": 11, "грудня": 12,
}


def order_metadata(filename: str, header_text: str = "") -> tuple[str, str]:
    """(номер, дата ISO yyyy-mm-dd або "")."""
    number, date = "", ""
    for source in (filename, (header_text or "")[:3000]):
        if not number:
            m = re.search(r"№\s*([A-Za-zА-Яа-яІіЇїЄєҐґ0-9\-/]+)", source)
            if m:
                number = m.group(1).strip()
        if not date:
            m = re.search(r"\b([0-9]{2})\.([0-9]{2})\.([0-9]{4})\b", source)
            if m:
                date = _iso(int(m.group(3)), int(m.group(2)), int(m.group(1)))
            else:
                m = re.search(r"(\d{1,2})\s*[»”’\"]?\s+(" + "|".join(_MONTHS) + r")\s+(\d{4})", source, re.I)
                if m:
                    date = _iso(int(m.group(3)), _MONTHS[m.group(2).lower()], int(m.group(1)))
    return number, date


def _iso(year: int, month: int, day: int) -> str:
    try:
        return datetime(year, month, day).strftime("%Y-%m-%d")
    except ValueError:
        return ""


# ─────────────────────────────────────────────────────────────────────────────
# Читання .docx
# ─────────────────────────────────────────────────────────────────────────────
def read_docx_paragraphs(path: str | Path) -> list[str]:
    """Абзаци тіла документа разом із таблицями; файл лише читається."""
    with zipfile.ZipFile(path) as archive:
        data = archive.read("word/document.xml")
    root = ElementTree.fromstring(data)
    paragraphs: list[str] = []
    for paragraph in root.iter(f"{_W}p"):
        parts: list[str] = []
        for node in paragraph.iter():
            if node.tag == f"{_W}t" and node.text:
                parts.append(node.text)
            elif node.tag in (f"{_W}tab", f"{_W}br", f"{_W}cr"):
                parts.append(" ")
        paragraphs.append("".join(parts))
    return paragraphs


# ─────────────────────────────────────────────────────────────────────────────
# Розбір
# ─────────────────────────────────────────────────────────────────────────────
def iter_blocks(paragraphs: list[str]):
    """Склеює абзаци, розірвані посеред речення: (текст, №_першого_абзацу, §)."""
    section = ""
    buffer: list[str] = []
    start = 0
    for index, raw in enumerate(paragraphs):
        text = normalize_space(raw)
        if not text:
            continue
        if _SECTION_RE.match(text):
            if buffer:
                yield " ".join(buffer), start, section
                buffer = []
            section = text[:12]
            continue
        if not buffer:
            start = index
        buffer.append(text)
        if text.endswith((".", ":", ";")):
            yield " ".join(buffer), start, section
            buffer = []
    if buffer:
        yield " ".join(buffer), start, section


def _sentence_end(text: str) -> int:
    for match in _SENTENCE_END_RE.finditer(text):
        before = re.search(rf"([{_UP}{_LO}]+)$", text[: match.start()])
        if match.group() == "." and before and before.group(1).casefold() in _ABBREVIATIONS:
            continue
        return match.start()
    return len(text)


def _caps_run(text: str) -> str:
    """Слова без малих літер від початку `text` до «ВОС», малого слова чи крапки."""
    words: list[str] = []
    for token in text.split(" "):
        if not token:
            continue
        if token.startswith("ВОС") or _HAS_LOWER_RE.search(token):
            break
        words.append(token)
        if token.endswith((".", ";", ":")):
            break
    return " ".join(words).strip(" ,.;:-–—")


def _clean_position(text: str) -> str:
    text = normalize_space(text).strip(" ,.;:-–—")
    if len(text) < 4 or len(text) > 400 or not re.search(f"[{_LO}{_UP}]", text):
        return ""
    return text


def _target_after(rest: str) -> str:
    skip_until = 0
    for match in _CAPS_WORD_RE.finditer(rest):
        if match.start() < skip_until:
            continue  # слово всередині вже відкинутої групи («ІМЕНУВАТИ ЗАЛІЗНЯКОМ»)
        run = _caps_run(rest[match.start():])
        if not run:
            continue
        skip_until = match.start() + len(run)
        # «СКАСУВАТИ, ПОНОВИТИ … І ПРИЗНАЧИТИ НАЧАЛЬНИКОМ …» — беремо після ПРИЗНАЧИТИ.
        appoint = re.search(r"(?:^|\s)ПРИЗНАЧИТИ\s+(.+)$", run)
        if appoint:
            run = appoint.group(1)
        head = run.split(" ", 1)[0].strip(",")
        if head in _NOT_POSITION_START:
            return ""
        if head in ORDER_VERBS or head.endswith("ТИ") or head in _NOT_TARGET_HEADS:
            continue
        if not _TARGET_HEAD_RE.search(head):
            return ""  # «ЗАПАС», звання в лапках, «ОСОБИСТИЙ НОМЕР» — не посада
        # «… УКРАЇНИ І НАПРАВИТИ ДО СКЛАДУ МІСІЇ …» — друга дія вже не посада.
        run = re.sub(rf",?\s+(?:І|ТА)\s+[{_UP}]+ТИ\b.*$", "", run)
        return _clean_position(run)
    return ""


def _hit(text: str, role: str, section: str, paragraph: int, dictionary) -> PositionHit | None:
    if not text or not re.match(rf"[{_UP}{_LO}]", text):
        return None  # «1946 р.н», номер, лапки — не посада
    if dictionary is not None:
        match = dictionary.match_at_start(text)
        if match:
            return PositionHit(text, role, section, paragraph, match.nominative, True, match.end)
    from .unit_split import split_position_unit

    position = split_position_unit(text)[0]
    if not re.match(rf"[{_UP}{_LO}]", position):
        return None
    return PositionHit(text, role, section, paragraph, position.casefold(), False, len(position))


def extract_positions_from_block(
    text: str, paragraph: int = 0, section: str = "", dictionary=None
) -> list[PositionHit]:
    hits: list[PositionHit] = []
    persons = list(_PERSON_RE.finditer(text))
    for number, person in enumerate(persons):
        stop = persons[number + 1].start() if number + 1 < len(persons) else len(text)
        rest = text[person.end():stop]
        while True:
            bio = _LEADING_BIO_RE.match(rest)
            if not bio or not bio.group(0):
                break
            rest = rest[bio.end():]

        ends = [_sentence_end(rest)]
        for pattern in (_VOS_RE, _CURRENT_TAIL_RE):
            match = pattern.search(rest)
            if match:
                ends.append(match.start())
        caps = _CAPS_WORD_RE.search(rest)
        if caps:
            ends.append(caps.start())
        current_end = min(ends)

        current = _clean_position(rest[:current_end])
        if current and not _HAS_LOWER_RE.search(current):
            current = ""  # уся група великими — це вже не займана посада
        if re.match(rf"як(?:ий|а|е|і)(?![{_LO}])", current):
            current = ""  # «який перебуває у розпорядженні …» — не посада
        hit = _hit(current, ROLE_CURRENT, section, paragraph, dictionary)
        if hit:
            hits.append(hit)

        if caps:
            target = _target_after(rest[current_end:_sentence_end_after(rest, current_end)])
            hit = _hit(target, ROLE_TARGET, section, paragraph, dictionary)
            if hit:
                hits.append(hit)
    return hits


def _sentence_end_after(text: str, offset: int) -> int:
    """Кінець речення для ВЕЛИКОЇ групи: крапка перед пробілом або кінцем."""
    match = re.search(r"\.(?=\s|$)", text[offset:])
    return offset + match.end() if match else len(text)


def extract_positions(paragraphs: list[str], dictionary=None) -> list[PositionHit]:
    """`dictionary` — `PositionDictionary`; без нього береться стандартний довідник."""
    if dictionary is None:
        from .position_dictionary import load_position_dictionary

        dictionary = load_position_dictionary()
    hits: list[PositionHit] = []
    for text, start, section in iter_blocks(paragraphs):
        hits.extend(extract_positions_from_block(text, start, section, dictionary))
    return hits
