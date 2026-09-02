"""Нечіткий пошук назви частини не має «зшивати» уламки двох різних частин.

Між словами назви навмисно дозволено широкий проміжок — там стоять почесні
найменування. Але номер ІНШОЇ частини у цьому проміжку означає, що збіг
склеївся з двох різних частин у переліку.
"""

import pytest

from nodeautomationtoolkit.builtin_nodes.recipient_mapping import (
    _build_unit_fuzzy_pattern,
    _stem_variants,
    map_military_units,
)

TARGET = "158 окрема бригада підтримки"


@pytest.mark.parametrize(
    "text",
    [
        "до 158 окремої бригади підтримки",
        "158 окремої ордена Богдана Хмельницького бригади підтримки",
        "158 окремої бригади підтримки 22 армійського корпусу",
        "у 158 окремій бригаді підтримки",
    ],
)
def test_real_name_is_found(text):
    assert _build_unit_fuzzy_pattern(TARGET).search(text), text


@pytest.mark.parametrize(
    "text",
    [
        # «158» від батальйону + «бригади підтримки» від 47-ї
        "до 158 окремого батальйону зв'язку та 47 окремої бригади підтримки",
        "до 47 окремої бригади підтримки",
        "до 158 окремого батальйону зв'язку",
    ],
)
def test_pieces_of_two_units_do_not_stitch(text):
    assert not _build_unit_fuzzy_pattern(TARGET).search(text), text


def test_wrong_unit_does_not_get_an_extract():
    """Наскрізна перевірка: у витяги потрапляє лише справді згадана частина."""
    mapping = {
        TARGET: {
            "open_name": TARGET,
            "cipher": "А1158",
            "abbreviation": "158 обрп",
            "corps": "",
            "recipient_to": "Командиру військової частини А1158",
        },
        "158 окремий батальйон зв'язку": {
            "open_name": "158 окремий батальйон зв'язку",
            "cipher": "А1159",
            "abbreviation": "158 обз",
            "corps": "",
            "recipient_to": "Командиру військової частини А1159",
        },
        "47 окрема бригада підтримки": {
            "open_name": "47 окрема бригада підтримки",
            "cipher": "А4747",
            "abbreviation": "47 обрп",
            "corps": "",
            "recipient_to": "Командиру військової частини А4747",
        },
    }
    text = "§ 1\n1. Призначити до 158 окремого батальйону зв'язку та 47 окремої бригади підтримки."

    keys = list(map_military_units(text=text, mapping=mapping)["unit_paragraphs"].keys())

    assert any("А1159" in key for key in keys), keys
    assert any("А4747" in key for key in keys), keys
    assert not any("А1158" in key for key in keys), f"хибно додано 158 обрп: {keys}"


# ── Випадний голосний: вузол → вузла ─────────────────────────────────────────
_FLEETING_MAPPING = {
    "5 вузол зв'язку": {
        "open_name": "5 вузол зв'язку",
        "cipher": "А0005",
        "abbreviation": "5 вз",
        "corps": "",
        "recipient_to": "Командиру військової частини А0005",
        "destination_where": "м. Тестове",
    },
}


@pytest.mark.parametrize(
    "phrase",
    ["5 вузол зв'язку", "5 вузла зв'язку", "5 вузлу зв'язку", "5 вузлом зв'язку"],
)
def test_fleeting_vowel_name_is_found_in_every_case(phrase):
    """«о» в останньому складі зникає при відмінюванні: вузол → вузла.

    Шаблон будується з називного відмінка (стовпець A), тому без окремого
    варіанта стема жодна відмінкова форма не знаходилась і пункт лишався
    без адресата.
    """
    order = (
        "§ 1\n"
        "Відповідно до пункту 1 ПРИЗНАЧИТИ:\n"
        f"1. Майора ТЕСТЕНКА, начальника {phrase}, командиром роти.\n"
        "\nКомандир військової частини А0001\nполковник   Петро ТЕСТОВИЙ\n"
    )
    routes = map_military_units(text=order, mapping=_FLEETING_MAPPING)

    assert "5 вз А0005" in routes["unit_paragraphs"]


def test_stem_variants_do_not_touch_names_without_alternation():
    """Полк, центр, бригада не мають випадного голосного — стем один."""
    assert _stem_variants("полк") == ["полк"]
    assert _stem_variants("центр") == ["центр"]
    assert _stem_variants("вузол") == ["вузол", "вузл"]


def test_air_defense_command_post_matches_by_stable_core_from_column_a():
    table_name = (
        "10 командний пункт протиповітряної оборони оперативного командування"
    )
    mapping = {
        table_name: {
            "open_name": table_name,
            "cipher": "А0010",
            "abbreviation": "10 КПППО",
            "corps": "",
            "recipient_to": "Командиру військової частини А0010",
            "destination_where": "м. Тестове",
        }
    }
    item = (
        "1. Капітана ПРИХОДЬКА, начальника групи обліку та розподілу "
        "військовослужбовців 69 батальйону резерву – ОПЕРАТИВНИМ ЧЕРГОВИМ "
        "ВІДДІЛЕННЯ ОПЕРАТИВНИХ ЧЕРГОВИХ 10 КОМАНДНОГО ПУНКТУ "
        "ПРОТИПОВІТРЯНОЇ ОБОРОНИ."
    )

    routes = map_military_units(text=item, mapping=mapping)

    assert routes["unmatched_items"] == []
    assert list(routes["unit_paragraphs"]) == ["10 КПППО А0010"]


# ── Спільний корінь не має зливати дві різні частини ──────────────────────────
_SAME_NUMBER_MAPPING = {
    "99 окремий полк зв’язку": {
        "open_name": "99 окремий полк зв’язку",
        "cipher": "А0099",
        "abbreviation": "99 опз",
        "corps": "",
        "recipient_to": "Командиру військової частини А0099",
        "destination_where": "м. Тестове",
    },
    "99 окремий радіотехнічний батальйон": {
        "open_name": "99 окремий радіотехнічний батальйон",
        "cipher": "А0098",
        "abbreviation": "99 ортб",
        "corps": "",
        "recipient_to": "Командиру військової частини А0098",
        "destination_where": "м. Друге",
    },
}

_RADIO_ITEM = (
    "2. Капітана ТЕСТЕНКА Тараса Тарасовича, начальника групи програмного "
    "забезпечення роти мобільного зв’язку радіорелейного батальйону "
    "99 окремого полку зв’язку – КОМАНДИРОМ РАДІОРЕЛЕЙНОГО ВЗВОДУ – "
    "НАЧАЛЬНИКОМ СТАНЦІЇ РАДІОРЕЛЕЙНОЇ РОТИ БАТАЛЬЙОНУ ОПОРНИХ МЕРЕЖ "
    "ЗВ’ЯЗКУ ЦЬОГО САМОГО ПОЛКУ."
)


def test_shared_root_does_not_merge_two_units_with_the_same_number():
    """«радіотехнічний» не має ловити «РАДІОРЕЛЕЙНОГО».

    Обидві частини мають номер 99, тож захист «у проміжку немає іншого номера»
    тут не працює. Загальний шаблон брав перші 4 літери — «раді» збігалося з
    «РАДІОРЕЛЕЙНОГО», і витяг ішов ще й на батальйон, якого в наказі немає.
    """
    order = (
        "§ 1\nВідповідно до пункту 1 ПРИЗНАЧИТИ:\n\n"
        + _RADIO_ITEM
        + "\n\nКомандир військової частини А0001\nполковник   Петро ТЕСТОВИЙ\n"
    )
    routes = map_military_units(text=order, mapping=_SAME_NUMBER_MAPPING)

    assert list(routes["unit_paragraphs"]) == ["99 опз А0099"]


def test_shortened_word_form_is_still_matched():
    """Стеля префікса не має зламати скорочену форму слова.

    «ремонтно-відновлювальний» у таблиці й «ремонтно-відновного» в наказі
    розходяться після 6-го символу, тому 6 — саме стеля, а не бажане значення.
    """
    pattern = _build_unit_fuzzy_pattern("208 окремий ремонтно-відновлювальний полк")

    assert pattern.search("208 окремого ремонтно-відновного полку")
