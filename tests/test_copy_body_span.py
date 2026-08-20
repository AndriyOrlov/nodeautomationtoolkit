"""Межі тіла наказу для примірника.

Примірник має містити пункти РАЗОМ ІЗ ПІДПИСАНТОМ, але БЕЗ службової
таблиці розсилки. Підписант визначається тією самою логікою, що й у витягах.
"""

import importlib.util
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))


def _load_generator_module():
    spec = importlib.util.spec_from_file_location(
        "generate_extracts_span_tests", PROJECT_ROOT / "generate_extracts.py"
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


generator = _load_generator_module()
find_content_start_line = generator.find_content_start_line
find_distribution_cutoff_line = generator.find_distribution_cutoff_line
find_order_signer = generator._find_order_signer

ORDER = "\n".join(
    [
        "МІНІСТЕРСТВО ОБОРОНИ УКРАЇНИ",              # 0
        "НАКАЗ",                                      # 1
        "",                                           # 2
        "§ 1",                                        # 3  ← початок тіла
        "1. Солдата Петренка Петра призначити.",      # 4
        "1980 р.н.",                                  # 5
        "",                                           # 6
        "Командир військової частини А0000",          # 7  ← підписант
        "полковник            Іван ПЕТРЕНКО",         # 8
        "",                                           # 9
        "Розрахунок розсилки",                        # 10 ← відсікається
        "1. Архів",                                   # 11
    ]
)


def test_body_starts_at_paragraph_sign():
    assert find_content_start_line(ORDER) == 3


def test_distribution_table_is_cut_off():
    assert find_distribution_cutoff_line(ORDER) == 10


def test_signer_is_inside_copied_span():
    """Підписант має потрапити у примірник, а не бути відсіченим."""
    signer = find_order_signer(ORDER)
    assert signer is not None
    assert find_content_start_line(ORDER) <= signer["start_line"]
    assert signer["start_line"] < find_distribution_cutoff_line(ORDER)


def test_span_excludes_distribution_but_keeps_signer():
    start = find_content_start_line(ORDER)
    cutoff = find_distribution_cutoff_line(ORDER)
    copied = ORDER.splitlines()[start:cutoff]

    assert any("Командир військової частини" in line for line in copied)
    assert any("ПЕТРЕНКО" in line for line in copied)
    assert not any("Розрахунок розсилки" in line for line in copied)
    assert not any("Архів" in line for line in copied)


ORDER_WITH_TWO_SERVICE_BLOCKS = "\n".join(
    [
        "§ 1",                                        # 0
        "1. Пункт.",                                  # 1
        "",                                           # 2
        "Командир військової частини А0000",          # 3 ← підписант
        "полковник            Іван ПЕТРЕНКО",         # 4
        "",                                           # 5
        "Розрахунок розсилки витягів із наказу:",     # 6 ← ПЕРШИЙ службовий блок
        "1",
        "Розрахунок розсилки електронних повідомлень",
        "Згідно з оригіналом",                        # 9 ← ОСТАННІЙ маркер
    ]
)


def test_cutoff_takes_first_service_block_not_the_last():
    """Сканування з кінця повертало останній маркер, і таблиці розсилки
    лишалися всередині примірника."""
    signer = find_order_signer(ORDER_WITH_TWO_SERVICE_BLOCKS)
    assert signer is not None

    first_block = generator.find_service_block_line(
        ORDER_WITH_TWO_SERVICE_BLOCKS, signer["start_line"] + 1
    )
    last_marker = find_distribution_cutoff_line(ORDER_WITH_TWO_SERVICE_BLOCKS)

    assert first_block == 6
    assert last_marker == 9
    assert first_block < last_marker


def test_no_distribution_tables_reach_the_copy():
    signer = find_order_signer(ORDER_WITH_TWO_SERVICE_BLOCKS)
    cutoff = generator.find_service_block_line(
        ORDER_WITH_TWO_SERVICE_BLOCKS, signer["start_line"] + 1
    )
    copied = ORDER_WITH_TWO_SERVICE_BLOCKS.splitlines()[:cutoff]

    assert any("ПЕТРЕНКО" in line for line in copied)
    assert not any("Розрахунок розсилки" in line for line in copied)
    assert not any("Згідно з оригіналом" in line for line in copied)


CELL_MARK = chr(7)  # службовий знак кінця комірки таблиці у Word


class _FakeRange:
    def __init__(self, text):
        self.Text = text


class _FakeParagraph:
    def __init__(self, text):
        self.Range = _FakeRange(text)


class _FakeParagraphs:
    def __init__(self, texts):
        self._items = [_FakeParagraph(t) for t in texts]
        self.Count = len(self._items)

    def __call__(self, index):
        return self._items[index - 1]


class _FakeDoc:
    def __init__(self, texts):
        self.Paragraphs = _FakeParagraphs(texts)


def test_span_survives_table_inside_the_order():
    """Раніше межі рахувались через зіставлення рядків тексту з абзацами.

    `Content.Text` не розбиває комірки таблиці на рядки, а `doc.Paragraphs`
    рахує кожну окремо — нумерація зсувалась і підписант відрізався.
    """
    doc = _FakeDoc(
        [
            "НАКАЗ\r",                                  # 1 шапка
            "§ 1\r",                                     # 2 ← початок тіла
            "1. Пункт.\r",                               # 3
            "Комірка А" + CELL_MARK,                     # 4 таблиця
            "Комірка Б" + CELL_MARK,                     # 5 таблиця
            "\r",                                        # 6
            "Командир частини\r",                        # 7 підписант
            "полковник І. ПЕТРЕНКО\r",                   # 8 ← кінець копіювання
            "\r",                                        # 9
            "Розрахунок розсилки витягів із наказу:\r",  # 10 службовий блок
            "1\r",                                       # 11
            "Згідно з оригіналом\r",                     # 12
        ]
    )

    first, last = generator.App._order_body_span(doc)

    assert first == 2
    assert last == 8


def test_marker_word_inside_an_item_does_not_cut_the_signer():
    """Службовий блок шукається ПІСЛЯ підписанта.

    Якщо шукати від початку тіла, слово-маркер у тексті пункту обрізало
    наказ разом зі званням і прізвищем підписанта.
    """
    doc = _FakeDoc(
        [
            "§ 2\r",                                     # 1
            "11. Пункт із фразою розсилка: усередині\r",  # 2 ← пастка
            "\r",
            "Тимчасово виконуючий обов'язки\r",           # 4 підписант
            "командувача військ\r",                       # 5
            "полковник   І. ПЕТРЕНКО\r",                  # 6 ← звання + прізвище
            "\r",
            "Розрахунок розсилки витягів із наказу:\r",   # 8
            "Згідно з оригіналом\r",                      # 9
        ]
    )

    first, last = generator.App._order_body_span(doc)

    assert first == 1
    assert last == 6


def test_span_stops_before_first_service_block():
    doc = _FakeDoc(
        [
            "§ 1\r",
            "1. Пункт.\r",
            "Командир\r",
            "Розрахунок розсилки витягів із наказу:\r",
            "Згідно з оригіналом\r",
        ]
    )

    first, last = generator.App._order_body_span(doc)

    assert (first, last) == (1, 3)


def test_order_without_distribution_table_keeps_everything():
    text = "\n".join(["§ 1", "1. Пункт.", "", "Командир", "полковник   Іван ПЕТРЕНКО"])
    assert find_distribution_cutoff_line(text) == len(text.splitlines())
