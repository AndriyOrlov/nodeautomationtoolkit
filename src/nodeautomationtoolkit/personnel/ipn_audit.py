"""Перевірка РНОКПП у тексті наказу — попередження для журналу.

Заготовка: генератори витягів і повідомлень її ще НЕ викликають. Коли буде
підключено, вона лише додає рядки в журнал і не змінює жодного документа.

Рядки журналу НЕ містять ні РНОКПП, ні ПІБ — лише номер пункту й суть проблеми
(журнал знеособлюється, див. AGENT.md про `redact`).
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from nodeautomationtoolkit.personnel.ipn import check_ipn

# Пункт наказу: «1. …», «12) …» на початку рядка.
_ITEM_START_RE = re.compile(r"^\s*(\d{1,3})[.)]\s")
# РНОКПП не починається з 0 (це народження до 1927 року), а номер телефону
# «0501234567» — починається. Так телефони виконавців не дають хибних тривог.
# Крапка чи кома ПІСЛЯ номера — кінець речення («1234567890.»), а не частина числа.
_IPN_TOKEN_RE = re.compile(r"(?<![\d.,/+-])[1-9]\d{9}(?!\d)(?![.,/-]\d)")
_BIRTH_YEAR_RE = re.compile(r"\b(19\d{2}|20\d{2})\s*р\.?\s*н\.?", re.IGNORECASE)


@dataclass(frozen=True)
class IpnProblem:
    item: str
    kind: str  # "check_digit" | "birth_year"
    detail: str = ""

    def message(self) -> str:
        where = f"пункт {self.item}" if self.item else "текст до першого пункту"
        if self.kind == "check_digit":
            return f"УВАГА: {where} — РНОКПП не проходить контрольну перевірку (можлива описка)."
        return f"УВАГА: {where} — рік народження {self.detail} не збігається з РНОКПП."


def _split_items(text: str) -> list[tuple[str, str]]:
    items: list[tuple[str, list[str]]] = [("", [])]
    for line in str(text or "").splitlines():
        match = _ITEM_START_RE.match(line)
        if match:
            items.append((match.group(1), []))
        items[-1][1].append(line)
    return [(label, "\n".join(lines)) for label, lines in items if lines]


def find_ipn_problems(order_text: str) -> list[IpnProblem]:
    """Знаходить у пунктах РНОКПП з неправильною контрольною цифрою або не тим роком."""
    problems: list[IpnProblem] = []
    for label, body in _split_items(order_text):
        birth_years = {int(year) for year in _BIRTH_YEAR_RE.findall(body)}
        for token in _IPN_TOKEN_RE.findall(body):
            check = check_ipn(token)
            if not check.valid:
                problems.append(IpnProblem(label, "check_digit"))
                continue
            # Рік звіряємо, лише коли в пункті одна людина (один рік народження).
            if len(birth_years) == 1 and check.birth_date:
                [year] = birth_years
                if check.birth_date.year != year:
                    problems.append(IpnProblem(label, "birth_year", str(year)))
    return problems


def describe_ipn_problems(order_text: str) -> list[str]:
    """Рядки журналу для `find_ipn_problems`."""
    return [problem.message() for problem in find_ipn_problems(order_text)]
