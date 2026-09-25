"""Пакет `personnel`: довідники, відмінювання звань і посад, РНОКПП.

Перенесено з Excel-генератора наказів. Усі РНОКПП у тестах згенеровані тут же
з вигаданих дат — жодного справжнього номера.
"""

from datetime import date

import pytest

from nodeautomationtoolkit.personnel import check_ipn, decline_position, decline_rank
from nodeautomationtoolkit.personnel.declension import clear_dictionary_cache
from nodeautomationtoolkit.personnel.dictionaries import load_case_dictionary, normalize_key
from nodeautomationtoolkit.personnel.ipn import ipn_check_digit, ipn_matches_birth_year


def _fake_ipn(birth: date, serial: str = "0001") -> str:
    """Вигаданий, але формально правильний РНОКПП для дати народження."""
    days = (birth - date(1899, 12, 31)).days
    first_nine = f"{days:05d}{serial}"
    return first_nine + str(ipn_check_digit(first_nine))


# ── РНОКПП ──────────────────────────────────────────────────────────────────


def test_valid_ipn_gives_birth_date_and_sex():
    male = check_ipn(_fake_ipn(date(1990, 5, 17), "0013"))
    female = check_ipn(_fake_ipn(date(1985, 1, 2), "0024"))

    assert (male.kind, male.valid, male.birth_date, male.sex) == ("ipn", True, date(1990, 5, 17), "Ч")
    assert (female.valid, female.birth_date, female.sex) == (True, date(1985, 1, 2), "Ж")


def test_typo_in_ipn_is_caught_by_check_digit():
    ipn = _fake_ipn(date(1990, 5, 17))
    typo = ipn[:3] + str((int(ipn[3]) + 1) % 10) + ipn[4:]

    result = check_ipn(typo)

    assert (result.kind, result.valid) == ("invalid", False)
    assert "контрольна" in result.reason


@pytest.mark.parametrize(
    ("value", "kind", "valid"),
    [
        ("", "missing", False),
        ("немає ІПН", "missing", False),
        ("ID 123456789", "document", True),
        ("АА 123456", "document", True),
        ("12345", "invalid", False),
    ],
)
def test_other_forms_of_ipn_field(value, kind, valid):
    result = check_ipn(value)
    assert (result.kind, result.valid) == (kind, valid)


def test_birth_year_consistency():
    ipn = _fake_ipn(date(1990, 5, 17))

    assert ipn_matches_birth_year(ipn, "1990 р.н.") is True
    assert ipn_matches_birth_year(ipn, 1991) is False
    assert ipn_matches_birth_year("ID 123456789", 1990) is None


# ── Довідники ───────────────────────────────────────────────────────────────


def test_normalize_key_handles_latin_lookalikes_and_dashes():
    assert normalize_key("Cолдат") == normalize_key("солдат")  # латинська C
    assert normalize_key("штаб–сержант") == "штаб-сержант"
    assert normalize_key("  Старший\xa0 солдат ") == "старший солдат"


def test_user_dictionary_overrides_and_extends_default(tmp_path):
    (tmp_path / "ranks.csv").write_text(
        "Називний;Знахідний;Примітка\nкапітан;КАПІТАНА-ТЕСТ;\nтестовий чин;тестового чина;\n",
        encoding="utf-8-sig",
    )

    dictionary = load_case_dictionary("ranks.csv", tmp_path)

    assert dictionary.lookup("капітан", "З") == "КАПІТАНА-ТЕСТ"
    assert dictionary.lookup("тестовий чин", "З") == "тестового чина"
    assert dictionary.lookup("майор", "З") == "майора"  # стандартний рядок лишився


def test_default_dictionaries_are_readable_and_non_empty():
    ranks = load_case_dictionary("ranks.csv")
    positions = load_case_dictionary("positions.csv")

    assert len(ranks.phrases("З")) >= 40
    assert len(positions.phrases("О")) >= 500
    # Зіпсовані автоматичні форми з Excel («військових частином») не перенесено.
    assert not any("частином" in form for form in positions.phrases("О").values())


# ── Звання ──────────────────────────────────────────────────────────────────


@pytest.mark.parametrize(
    ("rank", "case", "expected"),
    [
        ("капітан", "З", "капітана"),
        ("старший лейтенант", "З", "старшого лейтенанта"),
        ("головний майстер-сержант", "З", "головного майстер-сержанта"),
        ("старшина 1 статті", "З", "старшину 1 статті"),
        ("старшина 1 статті", "Р", "старшини 1 статті"),
        ("капітан 3 рангу", "З", "капітана 3 рангу"),
        ("майор медичної служби", "З", "майора медичної служби"),
        ("Cтарший солдат", "З", "Старшого солдата"),
        ("ПІДПОЛКОВНИК", "З", "ПІДПОЛКОВНИКА"),
    ],
)
def test_decline_rank(rank, case, expected):
    result = decline_rank(rank, case)
    assert (result.text, result.found) == (expected, True)


def test_unknown_rank_is_returned_as_is():
    result = decline_rank("вигадане звання", "З")
    assert (result.text, result.found) == ("вигадане звання", False)


# ── Посади ──────────────────────────────────────────────────────────────────


@pytest.mark.parametrize(
    ("position", "case", "expected"),
    [
        ("командир взводу", "З", "командира взводу"),
        ("командир взводу", "О", "командиром взводу"),
        ("старший водій", "З", "старшого водія"),
        # Фрази цілком немає — слова відмінюються по одному, родові означення лишаються.
        ("бойовий медик роти охорони", "О", "бойовим медиком роти охорони"),
        ("старший водій-електрик взводу", "З", "старшого водія-електрика взводу"),
        ("Командир відділення", "З", "Командира відділення"),
    ],
)
def test_decline_position(position, case, expected):
    result = decline_position(position, case)
    assert (result.text, result.found) == (expected, True)


@pytest.mark.parametrize(
    ("position", "case", "expected"),
    [
        ("старша медична сестра", "З", "старшу медичну сестру"),
        ("старша медична сестра", "О", "старшою медичною сестрою"),
    ],
)
def test_feminine_noun_after_feminine_adjectives(position, case, expected):
    result = decline_position(position, case)
    assert (result.text, result.found) == (expected, True)


@pytest.mark.parametrize(
    "position",
    [
        "старший тестоман",       # прикметник без відомого іменника — не «старшого тестоман»
        "офіцер-тестоман",        # відома лише перша частина — не «офіцера-тестоман»
    ],
)
def test_partial_declension_is_not_returned(position):
    result = decline_position(position, "З")
    assert (result.text, result.found) == (position, False)


def test_unknown_position_is_not_guessed():
    result = decline_position("тестувальник вигаданого підрозділу", "З")
    assert (result.text, result.found) == ("тестувальник вигаданого підрозділу", False)


def test_user_positions_file_is_used(tmp_path):
    (tmp_path / "positions.csv").write_text(
        "Називний;Знахідний;Орудний\nтестувальник;тестувальника;тестувальником\n",
        encoding="utf-8-sig",
    )
    clear_dictionary_cache()

    result = decline_position("тестувальник вигаданого підрозділу", "О", user_directory=tmp_path)

    assert (result.text, result.found) == ("тестувальником вигаданого підрозділу", True)
