"""Шифрування частини разом із корпусом: без дублів і без відкритих ознак.

Закріплює три дефекти, знайдені на справжніх наказах (усі дані тут вигадані):

1. Стовпець D словника — це ПОСИЛАННЯ на рядок корпусу, і майже завжди в
   ньому стоїть скорочення («51 АК»), а не стовпець A того рядка. Прямий
   пошук за ключем його не розвʼязував, і в текст ішло саме скорочення, яке
   далі шифрував ще й власний патерн корпусу:
   «військової частини А0077 військової частини 51 АК військової частини А0051».
2. Почесне найменування в лапках («Тестовий Яр») лишалося після шифру —
   і як зайвий текст, і як відкрита ознака частини в закритому повідомленні.
3. Той самий шифр корпусу підряд не згортався, якщо між двома входженнями
   стояло почесне найменування.
"""

import pytest

from nodeautomationtoolkit.builtin_nodes.message_order import (
    _collapse_unit_phrase_repeats,
    _spans_source_and_destination,
    cipher_unit_names,
    reflow_soft_breaks,
)
from nodeautomationtoolkit.builtin_nodes.recipient_mapping import is_tck_entry


def _entry(open_name, cipher, abbreviation="", corps=""):
    return {
        "open_name": open_name,
        "cipher": cipher,
        "abbreviation": abbreviation,
        "corps": corps,
    }


def _mapping(corps_column, corps_abbreviation="51 АК"):
    """Словник із двох рядків: бригада та її корпус (усе вигадане)."""
    return {
        "77 окрема тестова бригада": _entry(
            "77 окрема тестова бригада", "А0077", "77 отбр", corps_column
        ),
        "51 армійський корпус": _entry(
            "51 армійський корпус", "А0051", corps_abbreviation
        ),
    }


@pytest.mark.parametrize(
    "corps_column, corps_abbreviation",
    [
        ("51 АК", "51 АК"),      # звичайний випадок: D = скорочення корпусу
        ("51 АК", ""),           # стовпець C корпусу порожній
        ("51АК", ""),            # скорочення без пробілу
        ("51 армійський корпус", "51 АК"),  # D = стовпець A корпусу
    ],
)
def test_corps_reference_resolves_to_cipher(corps_column, corps_abbreviation):
    text = "начальника служби 77 окремої тестової бригади"
    result, _, _ = cipher_unit_names(text, _mapping(corps_column, corps_abbreviation))
    assert result == "начальника служби військової частини А0077 військової частини А0051"
    assert "АК" not in result


def test_honorific_in_quotes_is_removed_after_cipher():
    text = "командира взводу 77 окремої тестової бригади «Тестовий Яр»"
    result, _, _ = cipher_unit_names(text, _mapping("51 АК"))
    assert result == "командира взводу військової частини А0077 військової частини А0051"


@pytest.mark.parametrize("quotes", ["«{}»", "“{}”", '"{}"'])
def test_honorific_removed_for_every_quote_style(quotes):
    text = "офіцера 77 окремої тестової бригади " + quotes.format("Тестовий Яр")
    result, _, _ = cipher_unit_names(text, _mapping("51 АК"))
    assert "Тестовий Яр" not in result
    assert result.rstrip() == "офіцера військової частини А0077 військової частини А0051"


def test_corps_named_after_honorific_is_not_duplicated():
    """Корпус, названий у наказі слідом за частиною, не додає другий шифр."""
    text = (
        "командира взводу 77 окремої тестової бригади «Тестовий Яр» "
        "51 армійського корпусу"
    )
    result, _, _ = cipher_unit_names(text, _mapping("51 АК"))
    assert result == "командира взводу військової частини А0077 військової частини А0051"


def test_uppercase_destination_keeps_uppercase():
    text = "КОМАНДИРОМ ВЗВОДУ 77 ОКРЕМОЇ ТЕСТОВОЇ БРИГАДИ «ТЕСТОВИЙ ЯР»"
    result, _, _ = cipher_unit_names(text, _mapping("51 АК"))
    assert result == "КОМАНДИРОМ ВЗВОДУ ВІЙСЬКОВОЇ ЧАСТИНИ А0077 ВІЙСЬКОВОЇ ЧАСТИНИ А0051"


def test_separate_mentions_are_not_collapsed():
    """Дві РІЗНІ згадки лишаються: згортається лише повтор підряд."""
    text = "зі складу 77 окремої тестової бригади до 51 армійського корпусу"
    result, _, _ = cipher_unit_names(text, _mapping("51 АК"))
    assert result == (
        "зі складу військової частини А0077 військової частини А0051 "
        "до військової частини А0051"
    )


def test_collapse_keeps_case_of_first_phrase():
    """Лишається ПЕРШИЙ зворот — він несе відмінок речення та регістр."""
    assert (
        _collapse_unit_phrase_repeats("у військовій частині військової частини А0000")
        == "у військовій частині А0000"
    )
    assert (
        _collapse_unit_phrase_repeats("ВІЙСЬКОВОЇ ЧАСТИНИ ВІЙСЬКОВОЇ ЧАСТИНИ А0000")
        == "ВІЙСЬКОВОЇ ЧАСТИНИ А0000"
    )


def test_locative_case_survives_corps_suffix():
    """Відмінок ставиться на першу половину, корпус лишається родовим (9.5)."""
    result, _, _ = cipher_unit_names(
        "у 77 окремій тестовій бригаді", _mapping("51 АК")
    )
    assert result == "у військовій частині А0077 військової частини А0051"


# ── Ланцюг підпорядкованості: батальйон → бригада → корпус ───────────────────
#
# Нумерований батальйон має власний шифр, тому ТРИ ланки — це норма, а не
# дубль. Прибирати можна лише ПОВТОР того самого шифру.


def _chain_mapping(battalion_corps_column):
    """Батальйон у бригаді, бригада в корпусі (усе вигадане).

    `battalion_corps_column` — те, що стоїть у стовпці D батальйону. Стовпець
    називається «Корпус», тому там буває і корпус, і бригада; обидва варіанти
    мають давати той самий ланцюг.
    """
    return {
        "11 окремий тестовий батальйон": _entry(
            "11 окремий тестовий батальйон", "А1111", "11 отб", battalion_corps_column
        ),
        "22 окрема тестова бригада": _entry(
            "22 окрема тестова бригада", "А2222", "22 отбр", "33 АК"
        ),
        "33 армійський корпус": _entry("33 армійський корпус", "А3333", "33 АК"),
    }


@pytest.mark.parametrize(
    "battalion_corps_column",
    ["22 окрема тестова бригада", "33 АК"],
)
def test_three_links_of_subordination_are_kept(battalion_corps_column):
    text = (
        "командира взводу 11 окремого тестового батальйону "
        "22 окремої тестової бригади 33 армійського корпусу"
    )
    result, _, _ = cipher_unit_names(text, _chain_mapping(battalion_corps_column))
    assert result == (
        "командира взводу військової частини А1111 "
        "військової частини А2222 військової частини А3333"
    )


def test_corps_from_column_d_does_not_break_chain_order():
    """Корпус у стовпці D батальйону не має ставати перед бригадою.

    Було: «А1111 А3333 А2222 А3333» — корпус тягнули і батальйон, і бригада.
    Лишається ОСТАННЯ поява шифру: більше зʼєднання завжди стоїть далі.
    """
    text = "командира взводу 11 окремого тестового батальйону 22 окремої тестової бригади"
    result, _, _ = cipher_unit_names(text, _chain_mapping("33 АК"))
    assert result == (
        "командира взводу військової частини А1111 "
        "військової частини А2222 військової частини А3333"
    )


def test_chain_keeps_case_of_first_link():
    text = "у 11 окремому тестовому батальйоні 22 окремої тестової бригади"
    result, _, _ = cipher_unit_names(text, _chain_mapping("33 АК"))
    assert result.startswith("у військовій частині А1111 військової частини А2222")


def test_chain_keeps_uppercase():
    text = (
        "КОМАНДИРОМ ВЗВОДУ 11 ОКРЕМОГО ТЕСТОВОГО БАТАЛЬЙОНУ "
        "22 ОКРЕМОЇ ТЕСТОВОЇ БРИГАДИ 33 АРМІЙСЬКОГО КОРПУСУ"
    )
    result, _, _ = cipher_unit_names(text, _chain_mapping("33 АК"))
    assert result == (
        "КОМАНДИРОМ ВЗВОДУ ВІЙСЬКОВОЇ ЧАСТИНИ А1111 "
        "ВІЙСЬКОВОЇ ЧАСТИНИ А2222 ВІЙСЬКОВОЇ ЧАСТИНИ А3333"
    )


def test_chain_collapse_does_not_reach_across_other_words():
    """Дві різні ланки речення не зшиваються в один ланцюг."""
    assert (
        _collapse_unit_phrase_repeats(
            "зі складу військової частини А1111 військової частини А3333 "
            "до військової частини А3333"
        )
        == "зі складу військової частини А1111 військової частини А3333 "
        "до військової частини А3333"
    )


@pytest.mark.parametrize("separator", [" ", "\x0b", "\n"])
def test_chain_survives_soft_break_inside_name(separator):
    """Назву частини в наказі часто рве Shift+Enter (правило 4.2.8).

    Ланцюг має зібратися однаково, а перенос ПОЗА назвою — лишитись.
    """
    text = (
        f"командира{separator}гірсько-штурмового взводу "
        f"11 окремого тестового батальйону{separator}"
        f"22 окремої тестової бригади  33 армійського{separator}корпусу"
    )
    result, _, _ = cipher_unit_names(text, _chain_mapping("33 АК"))
    assert result == (
        f"командира{separator}гірсько-штурмового взводу військової частини А1111 "
        "військової частини А2222 військової частини А3333"
    )


# ── Мʼякі переноси у змісті повідомлення (правило закінчення) ────────────────


def test_soft_break_mid_phrase_is_stitched():
    """Розрив посеред фрази зшивається: інакше рядок лишається напівпорожнім."""
    assert reflow_soft_breaks(
        "командира взводу 11 окремого тестового батальйону\x0b22 окремої бригади"
    ) == "командира взводу 11 окремого тестового батальйону 22 окремої бригади"


@pytest.mark.parametrize("ending", [",", ".", ":", ";"])
def test_soft_break_after_finished_line_is_kept(ending):
    """Біографічний блок побудований переносами після закінчення — не чіпаємо."""
    text = f"1990 р. н., освіта: ТВІ у 2012 р.{ending}\x0bу ЗС - із 08.2008."
    assert reflow_soft_breaks(text) == text


def test_soft_break_stitching_removes_ragged_spaces():
    assert reflow_soft_breaks("командира  \x0b  взводу") == "командира взводу"


def test_reflow_leaves_text_without_breaks_untouched():
    text = "1. Лейтенанта ТЕСТЕНКА Андрія Андрійовича, командира взводу."
    assert reflow_soft_breaks(text) == text


# ── Запобіжник від проковтування тексту (правила 4.2.9 та 9.5.5) ─────────────
#
# Найгірший можливий наслідок — зникнення тексту наказу. Один збіг тягнувся
# від «звідки» (малими) до «КУДИ» (ВЕЛИКИМИ) і замінявся одним шифром.


_TCK_MAPPING = {
    "Тестівський обласний територіальний центр комплектування та соціальної підтримки":
        _entry(
            "Тестівський обласний територіальний центр комплектування та соціальної підтримки",
            "Тестівський ОТЦК та СП",
            "Тестівський ОТЦК та СП",
        ),
}


def test_match_does_not_swallow_text_between_source_and_destination():
    text = (
        "офіцера відділення рекрутингу та комплектування Прикладівського районного "
        "територіального центру комплектування та соціальної підтримки Тестівської "
        "області – НАЧАЛЬНИКОМ ГРУПИ ПСИХОЛОГІЧНОЇ ПІДТРИМКИ ПЕРСОНАЛУ "
        "ПРИКЛАДІВСЬКОГО РАЙОННОГО ТЕРИТОРІАЛЬНОГО ЦЕНТРУ КОМПЛЕКТУВАННЯ ТА "
        "СОЦІАЛЬНОЇ ПІДТРИМКИ ЦІЄЇ САМОЇ ОБЛАСТІ"
    )
    result, _, _ = cipher_unit_names(text, _TCK_MAPPING)
    assert result == text, "жоден символ пункту не має зникнути"


def test_both_halves_of_item_are_ciphered_separately():
    """Тире-роздільник не перетинається, але обидві половини шифруються."""
    text = (
        "командира взводу 11 окремого тестового батальйону – "
        "КОМАНДИРОМ ВЗВОДУ 22 ОКРЕМОЇ ТЕСТОВОЇ БРИГАДИ"
    )
    result, _, _ = cipher_unit_names(text, _chain_mapping("33 АК"))
    # Ліворуч названо лише батальйон, тож його ланцюг — «батальйон + корпус зі
    # стовпця D». Бригада з ВЕЛИКОЇ половини в цей ланцюг не потрапляє.
    assert result == (
        "командира взводу військової частини А1111 військової частини А3333 – "
        "КОМАНДИРОМ ВЗВОДУ ВІЙСЬКОВОЇ ЧАСТИНИ А2222 ВІЙСЬКОВОЇ ЧАСТИНИ А3333"
    )


@pytest.mark.parametrize(
    "matched, expected",
    [
        ("Тестівської області – НАЧАЛЬНИКОМ ГРУПИ ПСИХОЛОГІЧНОЇ ПІДТРИМКИ", True),
        ("11 окремого тестового батальйону", False),
        ("22 ОКРЕМОЇ ТЕСТОВОЇ БРИГАДИ", False),
        ("Тестівський ОТЦК та СП", False),          # абревіатура — не межа пункту
        ("22 окремої тестової бригади 33 АК", False),
        ("", False),
    ],
)
def test_source_destination_boundary_detection(matched, expected):
    assert _spans_source_and_destination(matched) is expected


def test_hyphen_inside_name_still_matches():
    """Заборона стосується тире з пробілами, а не дефіса всередині слова."""
    mapping = {
        "42 окрема гірсько-штурмова бригада": _entry(
            "42 окрема гірсько-штурмова бригада", "А4242", "42 огшбр"
        ),
    }
    result, _, _ = cipher_unit_names("командира взводу 42 окремої гірсько-штурмової бригади", mapping)
    assert result == "командира взводу військової частини А4242"


# ── ТЦК у змісті лишається ПОВНОЮ відкритою назвою (правило 9.5.6) ──────────


def test_tck_row_never_ciphers_content():
    """Рядок ТЦК не бере участі в шифруванні: назва лишається повною."""
    text = (
        "офіцера Тестівського обласного територіального центру комплектування "
        "та соціальної підтримки"
    )
    result, replaced, rows = cipher_unit_names(text, _TCK_MAPPING)
    assert result == text
    assert replaced == 0
    assert rows == []


def test_tck_row_does_not_hide_unit_ciphering():
    """Поряд із ТЦК звичайна частина шифрується як завжди."""
    mapping = dict(_TCK_MAPPING)
    mapping.update(_chain_mapping("33 АК"))
    text = (
        "офіцера Тестівського обласного територіального центру комплектування "
        "та соціальної підтримки, командира взводу 22 окремої тестової бригади"
    )
    result, _, _ = cipher_unit_names(text, mapping)
    assert result == (
        "офіцера Тестівського обласного територіального центру комплектування "
        "та соціальної підтримки, командира взводу військової частини А2222 "
        "військової частини А3333"
    )


@pytest.mark.parametrize(
    "entry_value, expected",
    [
        ({"open_name": "Тестовий обласний ТЦК та СП", "cipher": ""}, True),
        ({"open_name": "Тестовий центр комплектування та соціальної підтримки"}, True),
        ({"open_name": "22 окрема тестова бригада", "cipher": "А2222"}, False),
        ({"open_name": "33 армійський корпус", "cipher": "А3333"}, False),
    ],
)
def test_tck_entry_detection(entry_value, expected):
    assert is_tck_entry(entry_value) is expected
