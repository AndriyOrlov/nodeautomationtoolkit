"""Нечіткий пошук назви частини не має «зшивати» уламки двох різних частин.

Між словами назви навмисно дозволено широкий проміжок — там стоять почесні
найменування. Але номер ІНШОЇ частини у цьому проміжку означає, що збіг
склеївся з двох різних частин у переліку.
"""

import pytest

from nodeautomationtoolkit.builtin_nodes.recipient_mapping import (
    _build_unit_fuzzy_pattern,
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
