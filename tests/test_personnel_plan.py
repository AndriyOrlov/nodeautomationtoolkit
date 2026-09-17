"""Заготовки з Excel-генератора: рівень посади, перевірка РНОКПП у наказі, розбір плану.

Модулі ще не підключено до генераторів. Усі ПІБ, шифри й РНОКПП вигадані.
"""

from datetime import date

import pytest

from nodeautomationtoolkit.personnel.appointment import (
    EQUAL,
    HIGHER,
    LOWER,
    UNKNOWN,
    appointment_wording,
    compare_positions,
    mentions_vlk,
    shpk_level,
    tariff_grade,
)
from nodeautomationtoolkit.personnel.ipn import ipn_check_digit
from nodeautomationtoolkit.personnel.ipn_audit import describe_ipn_problems, find_ipn_problems
from nodeautomationtoolkit.personnel.plan_parser import (
    parse_candidate,
    parse_personal_details,
    parse_vacancy,
)


def _fake_ipn(birth: date, serial: str = "0013") -> str:
    days = (birth - date(1899, 12, 31)).days
    first_nine = f"{days:05d}{serial}"
    return first_nine + str(ipn_check_digit(first_nine))


# ── Рівень посади ───────────────────────────────────────────────────────────


@pytest.mark.parametrize(
    "text", ['шпк "старший солдат"', "шпк “старший солдат”", "шпк - старший солдат", "ШПК, старший солдат"]
)
def test_shpk_level_accepts_any_notation(text):
    assert shpk_level(text) == 2


def test_tariff_grade_notations():
    assert tariff_grade("12 т.р.") == 12
    assert tariff_grade("т.р. 7") == 7
    assert tariff_grade("4 тарифний розряд") == 4
    assert tariff_grade("без розряду") is None


@pytest.mark.parametrize(
    ("from_shpk", "to_shpk", "from_tariff", "to_tariff", "expected"),
    [
        ("шпк “солдат”", "шпк “сержант”", "", "", HIGHER),
        ("шпк “капітан”", "шпк “старший лейтенант”", "", "", LOWER),
        ("шпк “майор”", "шпк “майор”", "15 т.р.", "15 т.р.", EQUAL),
        ("шпк “майор”", "шпк “майор”", "16 т.р.", "15 т.р.", LOWER),
        ("шпк “майор”", "шпк “майор”", "15 т.р.", "16 т.р.", HIGHER),
        ("шпк “солдат”", "шпк “матрос”", "3 т.р.", "3 т.р.", EQUAL),  # корабельне = військовому
        ("", "шпк “сержант”", "", "", UNKNOWN),  # запасна рота
    ],
)
def test_compare_positions(from_shpk, to_shpk, from_tariff, to_tariff, expected):
    assert compare_positions(from_shpk, to_shpk, from_tariff, to_tariff) == expected


def test_appointment_wording_depends_on_vlk():
    assert appointment_wording(HIGHER) == "Призначається на вищу посаду у порядку просування по службі"
    assert "за станом здоров'я" in appointment_wording(LOWER, vlk=True)
    assert "за станом здоров'я" not in appointment_wording(LOWER, vlk=False)
    assert "за станом здоров'я" not in appointment_wording(HIGHER, vlk=True)
    assert appointment_wording(UNKNOWN) == "Призначається"


def test_mentions_vlk():
    assert mentions_vlk("висновок ВЛК від 01.02.2026 № 12")
    assert mentions_vlk("Обмежено придатний до військової служби")
    assert not mentions_vlk("рапорт військовослужбовця")


def test_user_wording_file_overrides_default(tmp_path):
    (tmp_path / "appointment_wording.csv").write_text(
        "Ключ;Текст\nвища;Тестове формулювання\n", encoding="utf-8-sig"
    )
    assert appointment_wording(HIGHER, user_directory=tmp_path) == "Тестове формулювання"
    assert appointment_wording(LOWER, user_directory=tmp_path).startswith("Призначається на нижчу")


# ── РНОКПП у тексті наказу ──────────────────────────────────────────────────


def test_ipn_audit_reports_item_without_personal_data():
    good = _fake_ipn(date(1990, 5, 17))
    typo = good[:-1] + str((int(good[-1]) + 1) % 10)
    order = "\n".join([
        "НАКАЗУЮ:",
        "1. Капітана ТЕСТЕНКА Тест Тестовича ПРИЗНАЧИТИ …",
        "1990 р.н., освіта: вища.",
        f"{good}.",
        "2. Майора ТЕСТОВОГО Петра Петровича ЗВІЛЬНИТИ …",
        "1985 р.н.",
        f"{typo}.",
        "3. Лейтенанта ПРОБНОГО Івана Івановича ПРИЗНАЧИТИ …",
        "1991 р.н.",
        f"{good}.",
        "Виконавець тел. 0501234567",
    ])

    problems = find_ipn_problems(order)
    lines = describe_ipn_problems(order)

    assert [(p.item, p.kind) for p in problems] == [("2", "check_digit"), ("3", "birth_year")]
    joined = "\n".join(lines)
    assert "пункт 2" in joined and "пункт 3" in joined
    assert good not in joined and typo not in joined and "ТЕСТ" not in joined


# ── Розбір плану переміщення ────────────────────────────────────────────────


def test_parse_candidate_full_line():
    result = parse_candidate(
        "старший солдат (12.05.2023) ТЕСТЕНКО Тест Тестович, 1234567890, стрілець 1 механізованого "
        "відділення 1 механізованого взводу військової частини А1111 з 01.06.2023, "
        'шпк "солдат", ВОС-100915А, 3 т.р., у ЗС - з 03.2022'
    )

    assert result.rank == "старший солдат"
    assert result.full_name == "ТЕСТЕНКО Тест Тестович"
    assert result.ipn == "1234567890"
    assert result.position == "стрілець 1 механізованого відділення 1 механізованого взводу"
    assert result.units == ["військової частини А1111"]
    assert (result.vos, result.shpk, result.tariff) == ("ВОС-100915А", "шпк “солдат”", "3 т.р.")
    assert result.service == "у ЗС - з 03.2022"


def test_parse_candidate_without_ipn_and_with_former_position():
    result = parse_candidate(
        "солдат ТЕСТОВИЙ Петро Петрович, немає ІПН, водій автомобільного відділення 2 запасної роти, "
        "колишній стрілець, шпк - солдат, ВОС – 837037А, т.р. 2"
    )

    assert (result.rank, result.full_name, result.ipn) == ("солдат", "ТЕСТОВИЙ Петро Петрович", "немає ІПН")
    assert result.position == "водій автомобільного відділення"
    assert result.former_position == "стрілець"
    assert result.units == ["2 запасної роти"]
    assert (result.vos, result.shpk, result.tariff) == ("ВОС-837037А", "шпк “солдат”", "2 т.р.")


def test_parse_vacancy_splits_rank_units_and_position():
    result = parse_vacancy(
        "старший солдат 3 запасної роти військової частини А2222 командир відділення, "
        "шпк «сержант», ВОС-100915А, 5 т.р., з 10.01.2024 (вакантна)"
    )

    assert result.rank == "старший солдат"
    assert result.units == ["3 запасної роти", "військової частини А2222"]
    assert result.position == "командир відділення"
    assert (result.shpk, result.tariff) == ("шпк “сержант”", "5 т.р.")


def test_parse_vacancy_without_units():
    result = parse_vacancy("навідник-оператор, шпк, старший солдат, ВОС-1009151, 4 тарифний розряд")

    assert (result.position, result.units) == ("навідник-оператор", [])
    assert (result.shpk, result.tariff) == ("шпк “старший солдат”", "4 т.р.")


def test_parse_personal_details():
    result = parse_personal_details(
        "15.03.1995, освіта: повна загальна середня, ЗОШ № 1 у 2012 р., у ЗС - з 02.2022, з 05.2023 по т.ч."
    )

    assert result.birth_year == "1995 р.н."
    assert result.education == "повна загальна середня, ЗОШ № 1 у 2012 р."
    assert result.service == "у ЗС - з 02.2022 та з 05.2023."
