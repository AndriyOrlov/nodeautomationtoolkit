"""Дані про особовий склад: довідники, відмінювання, перевірка ІПН.

Окремий пакет, перенесений з Excel-генератора наказів «Нормалізатор»: генератори
витягів, примірників і повідомлень його поки не використовують, тож їхню роботу
він не змінює. Довідники лежать окремими CSV у `dictionaries/` і правляться
користувачем без зміни коду (див. `dictionaries.py`).
"""

from nodeautomationtoolkit.personnel.declension import decline_position, decline_rank
from nodeautomationtoolkit.personnel.dictionaries import CaseDictionary, load_case_dictionary
from nodeautomationtoolkit.personnel.ipn import IpnCheck, check_ipn

__all__ = [
    "CaseDictionary",
    "IpnCheck",
    "check_ipn",
    "decline_position",
    "decline_rank",
    "load_case_dictionary",
]
