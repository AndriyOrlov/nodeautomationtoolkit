"""Розбір рядків плану переміщення — перенесено з VBA-модуля `plan_parser` Excel-генератора.

План переміщення має три текстові стовпці, кожен — один рядок довільного тексту:
- кандидат: «звання (дата), ПІБ, РНОКПП, займана посада, з якого часу, шпк, ВОС, т.р.»;
- вакансія: «посада, шпк, ВОС, т.р., з якого часу вакантна …»;
- особисті дані: «дата народження, освіта (рік закінчення), у ЗС з …».

Порядок витягування той самий, що у VBA: спершу однозначні шматки (ВОС, шпк,
т.р., ПІБ з РНОКПП, «у ЗС», «колишня посада», частини), вирізаються з рядка, а
те, що лишилось, ділиться на звання й посаду. Виправлено відомі вади VBA: частини
не губляться при кількох збігах, посада не ламається від дати в кінці.

Модуль ще НЕ підключено до генераторів. Для сканованих планів перед ним має
стояти розпізнавання тексту — тоді `normalize_plan_text` прибирає його сміття.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

# ── Спільна нормалізація (VBA `normalize`) ──────────────────────────────────
_INVISIBLE = dict.fromkeys(map(ord, "\u200b\u200c\u200d\ufeff"), None)


def normalize_plan_text(text: str) -> str:
    """Пробіли, лапки, невидимі символи; пробіл перед комою прибирається."""
    value = str(text or "").translate(_INVISIBLE).replace("\xa0", " ")
    value = value.replace("“", '"').replace("”", '"').replace("''", '"')
    value = value.replace("‘", "'").replace("’", "'")
    value = re.sub(r"\s+,", ",", value)
    return re.sub(r"\s+", " ", value).strip()


# ── Однозначні шматки ────────────────────────────────────────────────────────
_VOS_RE = re.compile(r"ВОС\s?[-–]\s?\d{6,7}(?:[^\d\s,]\s?/?\s?\d{3}|[^\d\s,])?", re.IGNORECASE)
_SHPK_RE = re.compile(r"шпк\s?(?:[\"«„“][^\"»”]+[\"»”]|-\s*[^,]+|,\s*[^,]+)", re.IGNORECASE)
_TARIFF_RE = re.compile(
    r"\d+\s*т\.?\s?р\.?|т\.?\s?р\.?\s*\d+|\d+\s+тарифн\w*\s+розряд\w*\.?", re.IGNORECASE
)
_NAME_WITH_ID_RE = re.compile(
    r"(?P<name>[^\s,]+\s+[^\s,]+\s+[^\s,]+)\s*,\s*"
    r"(?P<id>\d{10}|ID\s?\d{9}|[А-ЯІЇЄҐ]{2}\s?\d{6})",
    re.IGNORECASE,
)
_NAME_WITHOUT_ID_RE = re.compile(
    r"(?P<name>[^\s,]+\s+[^\s,]+\s+[^\s,]+)\s*,\s*немає\s+(?:ІПН|РНОКПП)", re.IGNORECASE
)
_SERVICE_RE = re.compile(r"[уУ]\s+ЗСУ?\s*[-–]?.+$")
_FORMER_RE = re.compile(r"колишн\w*\s+[^,]+", re.IGNORECASE)
_UNIT_RE = re.compile(
    r"\d+\s+запасної\s+роти|військової\s+частини\s+[A-ZА-ЯІЇЄҐ]\s?\d{4}", re.IGNORECASE
)
_DATE_TAIL_RE = re.compile(r"\s*(?:з|зі|від|зг\.?)\s*\d{1,2}\.\d{2}\.\d{2,4}.*$|\s*\([^)]*\)\s*$")
_BIRTH_RE = re.compile(r"\d{2}\.\d{2}\.(\d{4})|(\d{4})\s?р\.\s?н\.?")
_EDUCATION_RE = re.compile(r"освіта:.+?р\.", re.IGNORECASE)


@dataclass
class Candidate:
    rank: str = ""
    full_name: str = ""
    ipn: str = ""
    position: str = ""
    former_position: str = ""
    units: list[str] = field(default_factory=list)
    vos: str = ""
    shpk: str = ""
    tariff: str = ""
    service: str = ""


@dataclass
class Vacancy:
    rank: str = ""
    position: str = ""
    units: list[str] = field(default_factory=list)
    vos: str = ""
    shpk: str = ""
    tariff: str = ""


@dataclass
class PersonalDetails:
    birth_year: str = ""
    education: str = ""
    service: str = ""


def _cut(pattern: re.Pattern, text: str) -> tuple[str, str]:
    """Перший збіг і рядок без нього."""
    match = pattern.search(text)
    if not match:
        return "", text
    return match.group(0).strip(), text[: match.start()] + text[match.end():]


def _clean(text: str) -> str:
    value = re.sub(r",(\s*,)+", ",", text)
    value = re.sub(r"\s{2,}", " ", value)
    return value.strip(" ,.;")


def normalize_shpk(raw: str) -> str:
    """Будь-який запис шпк → «шпк “значення”» (VBA: лапки англійські U+201C/U+201D)."""
    value = re.sub(r"^\s*шпк", "", raw, flags=re.IGNORECASE)
    value = re.sub("[\"“”„«»'’‘]", "", value)
    value = value.strip(" -,")
    return f"шпк “{value}”" if value else ""


def normalize_tariff(raw: str) -> str:
    digits = re.search(r"\d+", raw or "")
    return f"{digits.group(0)} т.р." if digits else ""


def _extract_common(text: str) -> tuple[str, str, str, str]:
    vos, text = _cut(_VOS_RE, text)
    vos = re.sub(r"\s*[-–]\s*", "-", vos, count=1)
    shpk, text = _cut(_SHPK_RE, text)
    tariff, text = _cut(_TARIFF_RE, text)
    return vos, normalize_shpk(shpk), normalize_tariff(tariff), text


def parse_candidate(line: str) -> Candidate:
    """Рядок «про кандидата» зі стовпця C плану переміщення."""
    text = normalize_plan_text(line)
    result = Candidate()
    result.vos, result.shpk, result.tariff, text = _extract_common(text)

    named = _NAME_WITH_ID_RE.search(text)
    if named:
        result.full_name = named.group("name")
        identifier = re.sub(r"\s+", "", named.group("id")).upper()
        result.ipn = re.sub(r"^(ID|[А-ЯІЇЄҐ]{2})(?=\d)", r"\1 ", identifier)
        text = text[: named.start()] + text[named.end():]
    else:
        unnamed = _NAME_WITHOUT_ID_RE.search(text)
        if unnamed:
            result.full_name = unnamed.group("name")
            result.ipn = "немає ІПН"
            text = text[: unnamed.start()] + text[unnamed.end():]

    result.service, text = _cut(_SERVICE_RE, text)
    result.service = result.service.replace(",", "").strip()

    former, text = _cut(_FORMER_RE, text)
    result.former_position = re.sub(r"^колишн\w*\s+", "", former, flags=re.IGNORECASE)

    result.units = [match.group(0) for match in _UNIT_RE.finditer(text)]
    text = _UNIT_RE.sub("", text)

    text = _clean(text)
    # Посада — усе після першої коми; звання — до неї (VBA: «, .+»).
    head, _, tail = text.partition(",")
    result.position = _clean(_DATE_TAIL_RE.sub("", tail))
    rank = re.sub(r"\([^)]*\)", "", head)
    rank = re.sub(r"\d+\.\d+\.\d+|\b\d+\b", "", rank)
    # ПІБ, якщо не вирізався разом із РНОКПП, стоїть у кінці звання ВЕЛИКИМ прізвищем.
    result.rank = _clean(rank)
    return result


def parse_vacancy(line: str) -> Vacancy:
    """Рядок «про вакансію» зі стовпця B плану переміщення."""
    text = normalize_plan_text(line)
    result = Vacancy()
    result.vos, result.shpk, result.tariff, text = _extract_common(text)
    text = _clean(text)

    matches = list(_UNIT_RE.finditer(text))
    if not matches:
        result.position = _clean(_DATE_TAIL_RE.sub("", text))
        return result
    result.units = [match.group(0) for match in matches]
    result.rank = _clean(text[: matches[0].start()])
    tail = text[matches[-1].end():]
    tail = re.sub(r"\s*(?:з|зі|від|зг\.?)\s*\d{1,2}\.\d{2}\.\d{2,4}\b.*$", "", tail)
    tail = re.sub(r"\s*\([^)]*\)|\s*/\d+", "", tail)
    result.position = _clean(tail)
    return result


def parse_personal_details(line: str) -> PersonalDetails:
    """Рядок «дата народження, освіта, досвід» зі стовпця D плану переміщення."""
    text = normalize_plan_text(line)
    result = PersonalDetails()

    birth = _BIRTH_RE.search(text)
    if birth:
        result.birth_year = f"{birth.group(1) or birth.group(2)} р.н."
        text = text[: birth.start()] + text[birth.end():]

    education = _EDUCATION_RE.search(text)
    if education:
        value = education.group(0)
        if "," in value:
            value = value.split(",")[0]
        text = text.replace(value, "", 1)
        result.education = _clean(re.sub(r"^освіта:\s*", "", value, flags=re.IGNORECASE))

    service = re.search(r"\s?[уУ]\s+ЗС\w?\s?[-–]?.+", text)
    if service:
        value = service.group(0)
        text = text.replace(value, "", 1)
        value = re.sub(r"\s*по\s+т\.ч\.", "", value)
        value = re.sub(r",\s*з\s+", " та з ", value)
        value = _clean(value)
        result.service = f"{value}." if value else ""

    # VBA: усе, що лишилось (навчальний заклад, рік закінчення), — продовження освіти.
    rest = _clean(text)
    if rest:
        result.education = f"{result.education}, {rest}" if result.education else rest
    if result.education and not result.education.endswith("."):
        result.education += "."
    return result
