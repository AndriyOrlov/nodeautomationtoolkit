"""Назва примірника має читатися МОДУЛЕМ ВИТЯГІВ (правило 3.3).

Дефект, який це закриває: примірник називався `прим_2_17.08.2026_413.docx`,
де немає ні «№ …», ні «від …». `extract_metadata_from_filename` шукає саме ці
маркери, тож переданий на вкладку витягів примірник приходив БЕЗ номера й дати
наказу, і поля доводилось заповнювати руками.

Ключова перевірка тут — не текст назви, а ЗАМИКАННЯ КОЛА:
`build_copy_two_filename` → `extract_metadata_from_filename` має повернути
рівно ті реквізити, з яких назву склали.
"""

import pytest

from generate_extracts import (
    COPY_FILENAME_PREFIX,
    build_copy_two_filename,
    extract_metadata_from_filename,
    is_generated_copy_filename,
)


# --------------------------------------------------------------------------
# Замикання кола: назва → реквізити
# --------------------------------------------------------------------------


@pytest.mark.parametrize(
    "order_num, order_date",
    [
        ("413", "17.08.2026"),
        ("55", "01.01.2026"),
        ("134-вк", "20.05.2025"),
        ("б-н", "03.03.2026"),
        ("1", "31.12.2026"),
    ],
)
def test_generated_name_gives_the_metadata_back(order_num, order_date):
    """Те, з чого склали назву, має з неї й прочитатися."""
    filename = build_copy_two_filename(order_num, order_date, "Наказ.docx")

    assert extract_metadata_from_filename(filename) == (order_num, order_date)


def test_the_exact_shape_agreed_with_the_user():
    assert (
        build_copy_two_filename("413", "17.08.2026", "Наказ.docx")
        == "2,3_№413 від 17.08.2026.docx"
    )


def test_old_shape_was_unreadable_and_that_was_the_bug():
    """Фіксуємо причину: у старій назві не було жодного маркера."""
    assert extract_metadata_from_filename("прим_2_17.08.2026_413.docx") == ("", "")


# --------------------------------------------------------------------------
# Реквізити не вигадуються
# --------------------------------------------------------------------------


@pytest.mark.parametrize(
    "order_num, expected_token",
    [("б/н", "б-н"), ("123/45", "123-45"), ("355/1", "355-1")],
)
def test_slash_in_the_number_stays_whole(order_num, expected_token):
    """Скісна риска у назві файлу неможлива — але номер не має обрізатись.

    З підкресленням «б/н» читалось назад як «б», а «123/45» як «123»: символ
    `_` не входить у шаблон пошуку номера, тож читання зупинялось на ньому.
    """
    filename = build_copy_two_filename(order_num, "17.08.2026", "Наказ.docx")

    assert f"№{expected_token} " in filename
    assert extract_metadata_from_filename(filename) == (expected_token, "17.08.2026")


def test_missing_metadata_keeps_the_order_own_name():
    """Номер і дата беруться лише з назви наказу — вигадувати їх не можна."""
    assert (
        build_copy_two_filename("", "", "Наказ без реквізитів.docx")
        == "2,3_Наказ без реквізитів.docx"
    )


def test_fallback_still_carries_whatever_the_order_name_had():
    """Якщо реквізити були в назві наказу, вони лишаються й у примірнику.

    Дата тут неповна, тож спрацьовує запасна гілка — але номер із власної
    назви наказу нікуди не дівається й читається далі.
    """
    filename = build_copy_two_filename("", "", "Наказ № 77 від 2026.docx")

    assert extract_metadata_from_filename(filename)[0] == "77"


def test_source_folders_are_not_part_of_the_name():
    filename = build_copy_two_filename("", "", r"C:\\Накази\\Наказ_355.docx")

    assert "\\\\" not in filename
    assert filename == "2,3_Наказ_355.docx"


# --------------------------------------------------------------------------
# Готовий примірник не можна брати як наказ
# --------------------------------------------------------------------------


def test_generated_copy_is_recognised_and_skipped():
    filename = build_copy_two_filename("413", "17.08.2026", "Наказ.docx")

    assert is_generated_copy_filename(filename) is True


def test_legacy_name_is_still_recognised():
    """Файли, зроблені до перейменування, лежать у теках і далі."""
    assert is_generated_copy_filename("прим_2_17.08.2026_413.docx") is True


def test_a_real_order_is_not_mistaken_for_a_copy():
    assert is_generated_copy_filename("Наказ № 413 від 17.08.2026.docx") is False
    assert is_generated_copy_filename("Наказ_355.docx") is False


def test_prefix_is_matched_only_at_the_start():
    """Наказ, у назві якого просто трапилось «2,3», не має бути пропущений."""
    assert is_generated_copy_filename("Наказ п. 2,3_додаток.docx") is False


def test_recognition_ignores_the_folder_part():
    assert is_generated_copy_filename(r"C:\\Примірники_2\\2,3_№413 від 17.08.2026.docx") is True


def test_prefix_constant_matches_the_built_name():
    filename = build_copy_two_filename("413", "17.08.2026", "Наказ.docx")

    assert filename.startswith(COPY_FILENAME_PREFIX + "_")
