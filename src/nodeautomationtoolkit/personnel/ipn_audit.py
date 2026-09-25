"""Перевірка РНОКПП у тексті наказу — попередження для журналу.

Викликає перевірка наказу (`order_review`); генератори витягів і повідомлень
її не викликають. Лише знаходить проблеми й не змінює жодного документа.

Рядки журналу НЕ містять ні РНОКПП, ні ПІБ — лише номер пункту й суть проблеми
(журнал знеособлюється, див. AGENT.md про `redact`).
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date

from nodeautomationtoolkit.personnel.ipn import check_ipn

# Пункт наказу: «1. …», «12) …» на початку рядка.
_ITEM_START_RE = re.compile(r"^\s*(\d{1,3})[.)]\s")
# РНОКПП не починається з 0 (це народження до 1927 року), а номер телефону
# «0501234567» — починається. Так телефони виконавців не дають хибних тривог.
# Крапка чи кома ПІСЛЯ номера — кінець речення («1234567890.»), а не частина числа.
_IPN_TOKEN_RE = re.compile(r"(?<![\d.,/+-])[1-9]\d{9}(?!\d)(?![.,/-]\d)")
_BIRTH_YEAR_RE = re.compile(r"\b(19\d{2}|20\d{2})\s*р\.?\s*н\.?", re.IGNORECASE)
_MONTHS = {
    "січня": 1, "лютого": 2, "березня": 3, "квітня": 4, "травня": 5, "червня": 6,
    "липня": 7, "серпня": 8, "вересня": 9, "жовтня": 10, "листопада": 11, "грудня": 12,
}
#: «11.07.1988 р.н.», «11.07.1988 року народження».
_BIRTH_NUMERIC_RE = re.compile(
    r"\b(\d{1,2})\.(\d{1,2})\.((?:19|20)\d{2})\s*(?:р\.?\s*н\.?|року\s+народження)", re.IGNORECASE
)
#: «Народився 11 серпня 1976 року», «Народилася 01.02.1990».
_BORN_RE = re.compile(
    r"\b(?P<verb>Народився|Народилася)\s+"
    r"(?:(?P<day>\d{1,2})\s+(?P<month>[а-яіїєґ]+)\s+(?P<year>(?:19|20)\d{2})"
    r"|(?P<nday>\d{1,2})\.(?P<nmonth>\d{1,2})\.(?P<nyear>(?:19|20)\d{2}))",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class IpnProblem:
    item: str
    kind: str  # "check_digit" | "birth_year" | "birth_date" | "sex"
    detail: str = ""
    #: Сам РНОКПП — лише щоб позначити місце в документі; у `message()` не потрапляє.
    token: str = ""

    def message(self) -> str:
        where = f"пункт {self.item}" if self.item else "текст до першого пункту"
        if self.kind == "check_digit":
            return f"УВАГА: {where} — РНОКПП не проходить контрольну перевірку (можлива описка)."
        if self.kind == "birth_date":
            return f"УВАГА: {where} — дата народження {self.detail} не збігається з РНОКПП."
        if self.kind == "sex":
            return f"УВАГА: {where} — стать у РНОКПП не збігається з «{self.detail}»."
        return f"УВАГА: {where} — рік народження {self.detail} не збігається з РНОКПП."


def _split_items(text: str) -> list[tuple[str, str]]:
    items: list[tuple[str, list[str]]] = [("", [])]
    for line in str(text or "").splitlines():
        match = _ITEM_START_RE.match(line)
        if match:
            items.append((match.group(1), []))
        items[-1][1].append(line)
    return [(label, "\n".join(lines)) for label, lines in items if lines]


def _safe_date(year, month, day) -> date | None:
    try:
        return date(int(year), int(month), int(day))
    except (TypeError, ValueError):
        return None


def _birth_facts(body: str) -> tuple[set[date], set[int], set[str]]:
    """Повні дати, роки й «Народився»/«Народилася» з пункту."""
    dates: set[date] = set()
    years = {int(year) for year in _BIRTH_YEAR_RE.findall(body)}
    verbs: set[str] = set()
    for day, month, year in _BIRTH_NUMERIC_RE.findall(body):
        years.add(int(year))
        dates.add(_safe_date(year, month, day))
    for match in _BORN_RE.finditer(body):
        verbs.add(match.group("verb").casefold())
        if match.group("day"):
            month = _MONTHS.get(match.group("month").casefold())
            year, day = match.group("year"), match.group("day")
        else:
            year, month, day = match.group("nyear"), match.group("nmonth"), match.group("nday")
        years.add(int(year))
        dates.add(_safe_date(year, month, day))
    dates.discard(None)
    return dates, years, verbs


def find_ipn_problems(order_text: str) -> list[IpnProblem]:
    """РНОКПП у пунктах: контрольна цифра, дата (або рік) народження, стать.

    Дата й стать звіряються, лише коли в пункті одна людина — одна дата
    (один рік) і одне «Народився/Народилася».
    """
    problems: list[IpnProblem] = []
    for label, body in _split_items(order_text):
        dates, years, verbs = _birth_facts(body)
        for token in _IPN_TOKEN_RE.findall(body):
            check = check_ipn(token)
            if not check.valid:
                problems.append(IpnProblem(label, "check_digit", token=token))
                continue
            if not check.birth_date:
                continue
            if len(dates) == 1:
                [birth] = dates
                if check.birth_date != birth:
                    problems.append(IpnProblem(label, "birth_date", birth.strftime("%d.%m.%Y"), token))
            elif len(years) == 1:
                [year] = years
                if check.birth_date.year != year:
                    problems.append(IpnProblem(label, "birth_year", str(year), token))
            if len(verbs) == 1 and check.sex:
                [verb] = verbs
                if (verb == "народився") != (check.sex == "Ч"):
                    problems.append(IpnProblem(label, "sex", verb.capitalize(), token))
    return problems


def describe_ipn_problems(order_text: str) -> list[str]:
    """Рядки журналу для `find_ipn_problems`."""
    return [problem.message() for problem in find_ipn_problems(order_text)]
