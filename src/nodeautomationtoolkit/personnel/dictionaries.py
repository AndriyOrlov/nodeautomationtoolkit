"""Довідники відмінкових форм у CSV-файлах, які користувач править сам.

Формат файлу: UTF-8 (з BOM, щоб Excel відкривав кирилицю), роздільник `;`,
перший рядок — заголовки. Перший стовпець — називний відмінок, далі стовпці з
назвами відмінків («Родовий», «Знахідний», «Орудний» …), стовпець «Примітка»
ігнорується. Порожня клітинка означає, що форми немає.

Стандартні довідники лежать поруч (`dictionaries/`). Якщо передати власну теку,
однойменний файл у ній ДОПОВНЮЄ стандартний: рядки користувача перекривають
стандартні з тим самим називним відмінком, решта лишається.
"""

from __future__ import annotations

import csv
import re
from dataclasses import dataclass, field
from pathlib import Path

DEFAULT_DICTIONARY_DIR = Path(__file__).with_name("dictionaries")

CASE_COLUMNS = {
    "називний": "Н",
    "родовий": "Р",
    "давальний": "Д",
    "знахідний": "З",
    "орудний": "О",
    "місцевий": "М",
}

# Латинські літери, схожі на кириличні: «Cолдат» із латинською C, «командиp»
# з латинською p. В Excel-довіднику такі описки були окремими рядками.
_LATIN_LOOKALIKES = str.maketrans("ABCEHIKMOPTXaceiopxy", "АВСЕНІКМОРТХасеіорху")
_APOSTROPHES_RE = re.compile("[’ʼ`ʻ‘´]")
_DASHES_RE = re.compile("[–—‑]")


def normalize_key(text: str) -> str:
    """Ключ для пошуку: регістр, схожі латинські літери, апострофи, тире, пробіли."""
    value = str(text or "").replace("\xa0", " ").translate(_LATIN_LOOKALIKES)
    value = _APOSTROPHES_RE.sub("'", value)
    value = _DASHES_RE.sub("-", value)
    value = re.sub(r"\s*-\s*", "-", value) if " - " not in value else value
    return re.sub(r"\s+", " ", value).strip().casefold()


@dataclass
class CaseDictionary:
    """Форми слів і словосполучень за відмінками: `forms[відмінок][називний] = форма`."""

    forms: dict[str, dict[str, str]] = field(default_factory=dict)
    _word_cache: dict[str, dict[str, str]] = field(default_factory=dict, repr=False, compare=False)

    def lookup(self, nominative: str, case: str) -> str | None:
        return self.forms.get(case, {}).get(normalize_key(nominative)) or None

    def add(self, nominative: str, case: str, form: str) -> None:
        form = re.sub(r"\s+", " ", str(form or "")).strip()
        if form:
            self.forms.setdefault(case, {})[normalize_key(nominative)] = form
            self._word_cache.pop(case, None)

    def phrases(self, case: str) -> dict[str, str]:
        return self.forms.get(case, {})

    def word_forms(self, case: str) -> dict[str, str]:
        """Форми окремих слів, виведені з фраз («старший водій» дає «старший» → «старшого»).

        Береться лише слово, яке у фразі справді змінюється; перша знайдена
        форма має пріоритет.
        """
        if case not in self._word_cache:
            words: dict[str, str] = {}
            for nominative, form in self.phrases(case).items():
                nominative_words = nominative.split()
                form_words = form.split()
                if len(nominative_words) != len(form_words):
                    continue
                for nominative_word, form_word in zip(nominative_words, form_words, strict=True):
                    if nominative_word != normalize_key(form_word):
                        words.setdefault(nominative_word, form_word.casefold())
            self._word_cache[case] = words
        return self._word_cache[case]


def _read_csv(path: Path, dictionary: CaseDictionary) -> None:
    raw = path.read_text(encoding="utf-8-sig")
    rows = list(csv.reader(raw.splitlines(), delimiter=";"))
    if not rows:
        return
    header = [cell.strip().casefold() for cell in rows[0]]
    columns = [
        (index, CASE_COLUMNS[name]) for index, name in enumerate(header) if name in CASE_COLUMNS
    ]
    nominative_index = next((index for index, case in columns if case == "Н"), 0)
    for row in rows[1:]:
        if nominative_index >= len(row) or not row[nominative_index].strip():
            continue
        nominative = row[nominative_index]
        for index, case in columns:
            if index < len(row):
                dictionary.add(nominative, case, row[index])


def load_keyed_table(
    name: str, user_directory: str | Path | None = None
) -> dict[str, dict[str, str]]:
    """Довідник «ключ → рядок» (ключ — перший стовпець, нормалізований `normalize_key`).

    Для простих списків на кшталт коефіцієнтів шпк чи формулювань. Рядок із
    тим самим ключем у файлі користувача перекриває стандартний.
    """
    table: dict[str, dict[str, str]] = {}
    paths = [DEFAULT_DICTIONARY_DIR / name]
    if user_directory:
        paths.append(Path(user_directory) / name)
    for path in paths:
        if not path.is_file():
            continue
        rows = list(csv.reader(path.read_text(encoding="utf-8-sig").splitlines(), delimiter=";"))
        if not rows:
            continue
        header = [cell.strip() for cell in rows[0]]
        for row in rows[1:]:
            if not row or not row[0].strip():
                continue
            values = {
                column: (row[index].strip() if index < len(row) else "")
                for index, column in enumerate(header)
            }
            table[normalize_key(row[0])] = values
    return table


def load_case_dictionary(name: str, user_directory: str | Path | None = None) -> CaseDictionary:
    """Читає стандартний довідник `name` і доповнює його файлом користувача, якщо він є."""
    dictionary = CaseDictionary()
    default_path = DEFAULT_DICTIONARY_DIR / name
    if default_path.is_file():
        _read_csv(default_path, dictionary)
    if user_directory:
        user_path = Path(user_directory) / name
        if user_path.is_file():
            _read_csv(user_path, dictionary)
    return dictionary
