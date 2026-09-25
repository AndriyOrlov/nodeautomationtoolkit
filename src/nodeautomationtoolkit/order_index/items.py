"""Пункти наказу як цілі записи: особа, посади і повний текст.

Пункт — це абзац із «ПРІЗВИЩЕ Ім'я По батькові» плюс усі наступні абзаци до
наступної особи, наступного § чи заголовка («НАКАЗ», «Відповідно до …»):
біографія, освіта, «у ЗС», РНОКПП, «Призначається …», «Підстава: …».

Саме такий запис потрібен генератору наказів: за ІПН або ПІБ береться
найсвіжіший пункт людини, з нього переносяться вивірені біографічні рядки,
а посада призначення стає посадою, з якої особу призначають наступного разу.

Текст пунктів і особисті дані лишаються лише в локальному індексі (файл
`order_positions_index.sqlite` у папці, яку обрав користувач) — до репозиторію
вони не потрапляють (PROJECT_RULES 1.1, 1.2).
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from .extractor import (
    _LEADING_BIO_RE,
    _PERSON_RE,
    ROLE_CURRENT,
    ROLE_TARGET,
    PositionHit,
    extract_positions_from_block,
    iter_blocks,
    normalize_space,
)

_UP = "А-ЯІЇЄҐ"
_LO = "а-яіїєґ"
#: Заголовки, які перериваю пункт: новий наказ, шапка розділу, преамбула.
_HEADING_RE = re.compile(
    r"^(?:НАКАЗ|ЗАТВЕРДЖ|Додаток|Відповідно\s+до|У\s+зв['’]язку|Про\s+[а-яіїєґ])", re.IGNORECASE
)
_IPN_RE = re.compile(r"(?<!\d)(\d{10})(?!\d)")
_ITEM_LABEL_RE = re.compile(r"^(?:Пункт\s+)?(\d{1,3}(?:\.\d{1,3})*)[.)]\s+")
_RANK_RE = re.compile(rf"^(?:{'|'.join(['Генерал', 'Полковник', 'Підполковник', 'Майор', 'Капітан'])})", re.I)


@dataclass
class OrderItem:
    """Пункт наказу з особою, посадами й повним текстом."""

    section: str = ""
    label: str = ""  # «1.», якщо номер пункту стоїть у тексті
    paragraph: int = 0
    rank: str = ""  # звання так, як стоїть у пункті («Полковника»)
    surname: str = ""
    name: str = ""
    patronymic: str = ""
    ipn: str = ""
    current_position: str = ""  # займана посада з тексту (відкрита назва з частиною)
    target_position: str = ""  # посада призначення (ВЕЛИКИМИ, як у наказі)
    hits: list[PositionHit] = field(default_factory=list)
    text: str = ""  # пункт цілком, абзаци через \n

    @property
    def full_name(self) -> str:
        return " ".join(part for part in (self.surname, self.name, self.patronymic) if part)


def _person_parts(text: str) -> tuple[str, str, str, str]:
    """(звання, прізвище, ім'я, по батькові) з початку пункту."""
    match = _PERSON_RE.search(text)
    if not match:
        return "", "", "", ""
    words = match.group(0).strip(" ,").split()
    surname, name, patronymic = (words + ["", "", ""])[:3]
    rank = normalize_space(text[: match.start()]).strip(" ,")
    rank = re.sub(r"^.*?(?<![а-яіїєґ])(?=[А-ЯІЇЄҐ][а-яіїєґ]+\s*$)", "", rank) or rank
    return rank, surname, name, patronymic


def iter_items(paragraphs: list[str], dictionary=None) -> list[OrderItem]:
    """Пункти наказу в порядку документа.

    Береться той самий поділ на блоки, що й для посад: абзац, розірваний посеред
    речення («Старшого лейтенанта СИДОРОВА Петра» / «Петровича, командира …»),
    склеюється, інакше ПІБ у пункті не видно.
    """
    items: list[OrderItem] = []
    current: OrderItem | None = None
    for text, index, section in iter_blocks(paragraphs):
        if current is not None and current.section != section:
            current = None
        has_person = bool(_PERSON_RE.search(text))
        if _HEADING_RE.match(text) and not has_person:
            current = None
            continue
        if has_person:
            label = _ITEM_LABEL_RE.match(text)
            body = text[label.end():] if label else text
            rank, surname, name, patronymic = _person_parts(body)
            current = OrderItem(
                section=section,
                label=label.group(1) if label else "",
                paragraph=index,
                rank=rank,
                surname=surname,
                name=name,
                patronymic=patronymic,
                text=body,
            )
            items.append(current)
            continue
        if current is not None:
            current.text = f"{current.text}\n{text}"

    for item in items:
        _fill_positions(item, dictionary)
    return items


def _fill_positions(item: OrderItem, dictionary) -> None:
    block = " ".join(item.text.split("\n"))
    item.hits = extract_positions_from_block(block, item.paragraph, item.section, dictionary)
    for hit in item.hits:
        if hit.role == ROLE_CURRENT and not item.current_position:
            item.current_position = hit.text
        elif hit.role == ROLE_TARGET and not item.target_position:
            item.target_position = hit.text
    ipn = _IPN_RE.search(_LEADING_BIO_RE.sub("", item.text))
    if ipn:
        item.ipn = ipn.group(1)


def person_key(surname: str, name: str = "", patronymic: str = "") -> str:
    """Ключ пошуку за ПІБ, однаковий у будь-якому відмінку.

    У наказі ПІБ стоїть у знахідному («ІВАНОВА Івана Івановича»), а шукають
    зазвичай у називному, тож кожне слово зводиться до основи:
    «ІВАНОВА Івана Івановича» і «Іванов Іван Іванович» дають один ключ.
    """
    from .position_dictionary import stem

    words = [word for part in (surname, name, patronymic) for word in str(part or "").split()]
    return " ".join(stem(word) for word in words if word)
