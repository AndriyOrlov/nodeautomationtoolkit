"""Перевірка РНОКПП (ІПН) — перенесено з формули Excel-генератора.

РНОКПП з 10 цифр містить у собі:
- цифри 1–5 — кількість днів від 31.12.1899 до дати народження;
- цифра 9 — стать (непарна — чоловіча, парна — жіноча);
- цифра 10 — контрольна: сума цифр 1–9 з вагами (-1, 5, 7, 9, 4, 6, 10, 5, 7),
  остача від ділення на 11, потім остача від ділення на 10.

Модуль лише перевіряє й нічого не виправляє: помилку набору в наказі має
побачити людина. Паспорт («АА 123456») і ID-картка («ID 123456789») замість
РНОКПП — не помилка, для них повертається `kind="document"`.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, timedelta

_WEIGHTS = (-1, 5, 7, 9, 4, 6, 10, 5, 7)
_EPOCH = date(1899, 12, 31)
_IPN_RE = re.compile(r"^\d{10}$")
_DOCUMENT_RE = re.compile(r"^(?:ID\s?\d{9}|[А-ЯІЇЄҐA-Z]{2}\s?\d{6})$", re.IGNORECASE)


@dataclass(frozen=True)
class IpnCheck:
    kind: str  # "ipn" | "document" | "missing" | "invalid"
    valid: bool
    birth_date: date | None = None
    sex: str | None = None  # "Ч" | "Ж"
    reason: str = ""


def ipn_check_digit(first_nine: str) -> int:
    total = sum(int(digit) * weight for digit, weight in zip(first_nine, _WEIGHTS, strict=True))
    return total % 11 % 10


def check_ipn(value: str) -> IpnCheck:
    """Перевіряє РНОКПП: формат, контрольну цифру; повертає дату народження й стать."""
    text = re.sub(r"\s+", " ", str(value or "")).strip()
    if not text or re.fullmatch(r"немає\s+(?:ІПН|РНОКПП)", text, re.IGNORECASE):
        return IpnCheck("missing", False, reason="РНОКПП не вказано")
    if _DOCUMENT_RE.match(text):
        return IpnCheck("document", True, reason="замість РНОКПП вказано документ")
    digits = text.replace(" ", "")
    if not _IPN_RE.match(digits):
        return IpnCheck("invalid", False, reason="РНОКПП має складатися з 10 цифр")

    birth_date = _EPOCH + timedelta(days=int(digits[:5]))
    sex = "Ч" if int(digits[8]) % 2 else "Ж"
    if ipn_check_digit(digits[:9]) != int(digits[9]):
        return IpnCheck("invalid", False, birth_date, sex, "не збігається контрольна цифра")
    return IpnCheck("ipn", True, birth_date, sex)


def ipn_matches_birth_year(value: str, year: int | str) -> bool | None:
    """Чи збігається рік народження з РНОКПП. `None` — перевірити нема чим."""
    check = check_ipn(value)
    if check.birth_date is None:
        return None
    try:
        return check.birth_date.year == int(str(year).strip()[:4])
    except ValueError:
        return None
