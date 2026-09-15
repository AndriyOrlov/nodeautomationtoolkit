"""Номер пункту, який парсер не впізнавав, зливав пункт у шапку наказу.

Відтворення наказу №459 (15.09.2026): пункт 1 набрано «1.Капітана» без
пропуску (або з автонумерацією Word). Він не ставав пунктом, увесь його текст
ставав шапкою, частини з нього — контекстом для всіх наступних пунктів, і
пункт 2 про внутрішнє переміщення в центрі потрапляв у розсилку частин
пункту 1. Номери частин тут ВИГАДАНІ, але тієї самої довжини, що в наказі.
"""

import importlib.util
import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from nodeautomationtoolkit.builtin_nodes.recipient_mapping import (  # noqa: E402
    map_military_units,
    normalize_item_numbering,
)

A = "’"


def _entry(name: str, cipher: str, abbreviation: str, corps: str = "") -> dict:
    return {
        "open_name": name,
        "cipher": cipher,
        "abbreviation": abbreviation,
        "corps": corps,
        "recipient_to": f"Командиру військової частини {cipher}",
        "destination_where": "м. Тестове",
    }


MAPPING = {
    "57 окрема механізована бригада": _entry("57 окрема механізована бригада", "А2057", "57 омбр", corps="27 АК"),
    "27 армійський корпус": _entry("27 армійський корпус", "А2027", "27 АК"),
    f"89 окремий полк зв{A}язку": _entry(f"89 окремий полк зв{A}язку", "А2089", "89 опз"),
    "557 центр підготовки підрозділів": _entry("557 центр підготовки підрозділів", "А2557", "557 цпп"),
}

HEADING = "Відповідно до статті 1 нижчепойменованих військовослужбовців ПРИЗНАЧИТИ:"
BODY_1 = (
    "Капітана ТЕСТЕНКА Андрія Андрійовича, заступника командира механізованої роти механізованого батальйону\n"
    f"57 окремої механізованої бригади 27 армійського корпусу – ЗАСТУПНИКОМ КОМАНДИРА РОТИ ЗВ{A}ЯЗКУ "
    f"З ОЗБРОЄННЯ БАТАЛЬЙОНУ МОБІЛЬНИХ ВУЗЛІВ ЗВ{A}ЯЗКУ 89 ОКРЕМОГО ПОЛКУ ЗВ{A}ЯЗКУ,"
)
BODY_2 = (
    "Майора ПРИКЛАДЕНКА Петра Петровича, старшого викладача циклової комісії вогневої підготовки школи "
    "підготовки підрозділів протиповітряної оборони 557 центру підготовки\n"
    "підрозділів – НАЧАЛЬНИКОМ ЦИКЛОВОЇ КОМІСІЇ ВОГНЕВОЇ ПІДГОТОВКИ ШКОЛИ ПІДГОТОВКИ ПІДРОЗДІЛІВ "
    "ПРОТИПОВІТРЯНОЇ ОБОРОНИ ЦЬОГО САМОГО ЦЕНТРУ,"
)
BIO = "1990 р. н., освіта: вища у 2012 р.\n1234567890."


def _extracts(result) -> dict[str, list[str]]:
    return {key: [item["label"] for item in data["items"]] for key, data in result["unit_paragraphs"].items()}


@pytest.mark.parametrize(
    "first_number",
    ["1. ", "1.", "1)", "1.​", "﻿1.", "1. "],
    ids=["з пропуском", "без пропуску", "дужка без пропуску", "нульова ширина", "BOM", "нерозривний"],
)
def test_first_item_is_never_swallowed_into_the_heading(first_number):
    text = "\n".join([HEADING, first_number + BODY_1, BIO, "2. " + BODY_2, BIO])

    result = map_military_units(text=text, mapping=MAPPING)

    # Мітка пункту повторює набраний номер: «1.» → «Пункт 1.», «1)» → «Пункт 1».
    audit = {row["label"].rstrip("."): row for row in result["routing_audit"]}
    assert set(audit) == {"Пункт 1", "Пункт 2"}
    # Пункт 2 не отримує частин пункту 1 як «адресата із шапки».
    assert audit["Пункт 2"]["context_recipients"] == "—"
    assert "шапки" not in audit["Пункт 2"]["applied_rules"]
    extracts = {key: [label.rstrip(".") for label in labels] for key, labels in _extracts(result).items()}
    assert extracts == {
        "27 АК А2027": ["Пункт 1"],
        "89 опз А2089": ["Пункт 1"],
        "557 цпп А2557": ["Пункт 2"],
    }


def test_real_heading_with_a_unit_still_applies_to_every_item():
    """Правило 4.3 не змінилось: частина, названа в шапці, — вихідна для всіх пунктів."""
    heading = "Відповідно до статті 1 військовослужбовців 57 окремої механізованої бригади ПРИЗНАЧИТИ:"
    text = "\n".join([heading, "1." + BODY_1, BIO, "2." + BODY_2, BIO])

    result = map_military_units(text=text, mapping=MAPPING)

    audit = {row["label"].rstrip("."): row for row in result["routing_audit"]}
    assert set(audit) == {"Пункт 1", "Пункт 2"}
    assert audit["Пункт 2"]["context_recipients"] != "—"


# ── сама нормалізація ─────────────────────────────────────────────────────
@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        ("1.Капітана", "1. Капітана"),
        ("2)Майора", "2) Майора"),
        ("3.1.Лейтенанта", "3.1. Лейтенанта"),
        ("4.«Про військовий обовʼязок»", "4. «Про військовий обовʼязок»"),
        ("1.​Капітана", "1. Капітана"),
        ("﻿5.Сержанта", "5. Сержанта"),
        ("шапка\x0b6.Рядового", "шапка\x0b6. Рядового"),
    ],
)
def test_glued_item_numbers_get_a_space(raw, expected):
    assert normalize_item_numbering(raw) == expected


@pytest.mark.parametrize(
    "line",
    [
        "1. Капітана ТЕСТЕНКА",
        "01.09.2026 року",
        "1985 р. н., освіта: вища",
        "1234567890.",
        "з 02.09.2026 до 02.09.2028.",
        "Пункт 1.Капітана",
        "§ 1",
    ],
)
def test_other_lines_are_left_alone(line):
    assert normalize_item_numbering(line) == line


def test_line_count_is_preserved_and_normalization_is_idempotent():
    raw = "§ 1\n1.Капітана\n\n2)Майора\x0bпродовження\n1985 р. н."
    once = normalize_item_numbering(raw)
    assert once.count("\n") == raw.count("\n")
    assert once.splitlines() == ["§ 1", "1. Капітана", "", "2) Майора", "продовження", "1985 р. н."]
    assert normalize_item_numbering(once) == once


# ── автонумерація Word у тексті наказу ────────────────────────────────────
def _load_generator():
    spec = importlib.util.spec_from_file_location(
        "generate_extracts_item_numbering_tests", PROJECT_ROOT / "generate_extracts.py"
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class _ListFormat:
    def __init__(self, list_string):
        self.ListString = list_string


class _Range:
    def __init__(self, text, start, list_string=""):
        self.Text = text
        self.Start = start
        self.ListFormat = _ListFormat(list_string)


class _Paragraph:
    def __init__(self, text, start, list_string=""):
        self.Range = _Range(text, start, list_string)


class _Collection:
    def __init__(self, items):
        self._items = items
        self.Count = len(items)

    def __call__(self, index):
        return self._items[index - 1]

    def __iter__(self):
        return iter(self._items)


class _AutoNumberedDoc:
    """Абзаци з автонумерацією: номера в `Range.Text` немає, він у `ListString`."""

    def __init__(self, rows):
        paragraphs, start = [], 0
        for text, list_string in rows:
            paragraphs.append(_Paragraph(text + "\r", start, list_string))
            start += len(text) + 1
        self.Paragraphs = _Collection(paragraphs)
        self.ListParagraphs = _Collection([p for p in paragraphs if p.Range.ListFormat.ListString])


def test_word_auto_numbering_becomes_part_of_the_order_text():
    generator = _load_generator()
    doc = _AutoNumberedDoc([
        (HEADING, ""),
        ("Капітана ТЕСТЕНКА Андрія Андрійовича, командира роти.", "1."),
        ("1990 р. н.", ""),
        ("Майора ПРИКЛАДЕНКА Петра Петровича, начальника служби.", "2)"),
        ("Пункт маркованого списку.", "•"),
        ("3.Лейтенанта ЗРАЗКОВА, набраного вручну.", "3."),
    ])

    lines = generator.read_document_text(doc).splitlines()

    assert lines == [
        HEADING,
        "1. Капітана ТЕСТЕНКА Андрія Андрійовича, командира роти.",
        "1990 р. н.",
        "2) Майора ПРИКЛАДЕНКА Петра Петровича, начальника служби.",
        "Пункт маркованого списку.",
        "3. Лейтенанта ЗРАЗКОВА, набраного вручну.",
    ]
