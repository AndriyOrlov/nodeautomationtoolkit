"""Підстави звільнення з військової служби (стаття 26 Закону 2232-XII).

Підзаголовок у наказі виглядає так (зразки додатка 53):

    У ЗАПАС ЗА ПІДПУНКТОМ «а» (у зв'язку із закінченням строку контракту):
    У ВІДСТАВКУ ЗА ПІДПУНКТОМ «б» (за станом здоров'я):

Перелік підпунктів — окремий довідник `dismissal_reasons.csv`, який користувач
править сам: формулювання в законі змінюються частіше, ніж наш код. Ключ рядка —
«пункт.підпункт» частини п'ятої статті 26 («1.а» — мирний час, «2.к» —
особливий період); за самим підпунктом («а») береться перший відповідний рядок.

«Куди» (у запас чи у відставку) в довіднику — лише типове значення: той самий
підпункт «б» дає і запас, і відставку залежно від висновку ВЛК, тому останнє
слово за користувачем.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from nodeautomationtoolkit.personnel.dictionaries import load_keyed_table, normalize_key

DISMISSAL_FILE = "dismissal_reasons.csv"
RESERVE = "у запас"
RETIREMENT = "у відставку"


@dataclass(frozen=True)
class DismissalGround:
    """Один рядок довідника підстав."""

    key: str
    point: str  # пункт частини п'ятої статті 26
    letter: str  # підпункт
    reason: str  # текст у дужках
    destination: str  # типове «куди»
    note: str = ""

    def subheading(self, destination: str = "") -> str:
        """«У ЗАПАС ЗА ПІДПУНКТОМ «а» (у зв'язку із закінченням строку контракту):»"""
        where = (destination or self.destination or RESERVE).upper()
        reason = f" ({self.reason})" if self.reason else ""
        return f"{where} ЗА ПІДПУНКТОМ «{self.letter}»{reason}:"


def grounds(user_directory: str | Path | None = None) -> list[DismissalGround]:
    """Усі підстави в порядку довідника."""
    result = []
    for row in load_keyed_table(DISMISSAL_FILE, user_directory).values():
        letter = (row.get("Підпункт") or "").strip()
        if not letter:
            continue
        result.append(
            DismissalGround(
                key=(row.get("Ключ") or "").strip(),
                point=(row.get("Пункт") or "").strip(),
                letter=letter,
                reason=(row.get("Причина") or "").strip(),
                destination=(row.get("Куди") or RESERVE).strip(),
                note=(row.get("Примітка") or "").strip(),
            )
        )
    return result


def find_ground(key: str, user_directory: str | Path | None = None) -> DismissalGround | None:
    """Підстава за ключем «1.а», за підпунктом «а» або за словами самої причини."""
    wanted = normalize_key(key)
    if not wanted:
        return None
    table = grounds(user_directory)
    for ground in table:
        if normalize_key(ground.key) == wanted:
            return ground
    for ground in table:
        if normalize_key(ground.letter) == wanted:
            return ground
    # Пошук за текстом причини — лише для цілої фрази: одна літера «щ» інакше
    # знайшлася б усередині «очищення влади» й дала б чужу підставу.
    if len(wanted) >= 4:
        for ground in table:
            if wanted in normalize_key(ground.reason):
                return ground
    return None


def subheading(key: str, destination: str = "", user_directory: str | Path | None = None) -> str:
    """Підзаголовок групи; якщо підстави немає в довіднику — порожньо."""
    ground = find_ground(key, user_directory)
    return ground.subheading(destination) if ground else ""
