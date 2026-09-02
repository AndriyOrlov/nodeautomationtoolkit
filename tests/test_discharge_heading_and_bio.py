"""Накази про звільнення: шапка не дублюється, розрив стоїть перед «р. н.».

Дві вади, що виявилися разом на реальному наказі:

1. Шапку «Відповідно до … Закону України ⏎ “Про військовий обов'язок…”
   ЗВІЛЬНИТИ з військової служби:» розриває мʼякий перенос. Обидва уламки
   проходили як шапка, ставали ОКРЕМИМИ сегментами ієрархії, і обидва вказували
   на ОДИН абзац Word — у витягу шапка друкувалася ДВІЧІ.
2. `is_biographical_paragraph` шукала «р.н.» без пробілу, а в наказах (і в
   офіційному зразку) стоїть «р. н.». Через це першим біографічним абзацом
   вважався рядок ІПН, і обовʼязковий порожній абзац зʼїжджав на нього —
   зайвий розрив після «Підлягає направленню на військовий облік до …».
"""

import importlib.util
import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from nodeautomationtoolkit.builtin_nodes.recipient_mapping import (
    _starts_a_new_heading,
    map_military_units,
)


def _load_generator():
    spec = importlib.util.spec_from_file_location(
        "generate_extracts_discharge_tests", PROJECT_ROOT / "generate_extracts.py"
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


generator = _load_generator()

UNIT = "1 окремий тестовий загін"
MAPPING = {
    UNIT: {
        "open_name": UNIT,
        "cipher": "А0001",
        "abbreviation": "1 отз",
        "corps": "",
        "recipient_to": "Командиру військової частини А0001",
        "destination_where": "м. Тестове",
    }
}

HEAD_LINE_1 = "Відповідно до пункту 2 частини четвертої статті 26 Закону України"
HEAD_LINE_2 = (
    "“Про військовий обов’язок і військову службу” нижчепойменованих осіб "
    "офіцерського складу ЗВІЛЬНИТИ з військової служби:"
)
SUBHEADING = (
    "У ЗАПАС ЗА ПІДПУНКТОМ “б” (за станом здоров’я – за наявності інвалідності):"
)
ITEM = (
    "1. Полковника ЗВІЛЬНЕНКА Андрія Андрійовича, офіцера 1 окремого тестового "
    "загону, звільнити з військової служби."
)
SIGNER = "\nКомандир військової частини А0001\nполковник   Петро ТЕСТОВИЙ\n"


def _heading_ranges():
    order = (
        "§ 1\n"
        + HEAD_LINE_1 + "\n"
        + HEAD_LINE_2 + "\n\n"
        + SUBHEADING + "\n\n"
        + ITEM + "\n1971 р. н.\n1234567890.\n"
        + SIGNER
    )
    routes = map_military_units(text=order, mapping=MAPPING)
    entry = next(iter(routes["unit_paragraphs"].values()))
    return entry["items"][0]["heading_ranges"]


def test_wrapped_discharge_heading_is_one_segment_not_two():
    """Шапка, розірвана переносом, має лишатися ОДНИМ сегментом ієрархії."""
    ranges = _heading_ranges()

    # § + шапка + підшапка = рівно три рівні
    assert len(ranges) == 3, ranges
    assert ranges[1] == (1, 3), f"шапку розбито на сегменти: {ranges}"


def test_no_source_line_is_copied_into_the_extract_twice():
    """Головний симптом: рядок шапки потрапляв у витяг двічі."""
    seen: dict[int, int] = {}
    for start, end in _heading_ranges():
        for line in range(start, end + 1):
            seen[line] = seen.get(line, 0) + 1

    duplicated = sorted(line for line, count in seen.items() if count > 1)
    assert not duplicated, f"рядки копіюються двічі: {duplicated}"


@pytest.mark.parametrize(
    "line",
    ["1971 р.н.", "1971 р. н.", "1971 р.  н.", "1971 року народження"],
    ids=["злитно", "з-пробілом", "два-пробіли", "словами"],
)
def test_birth_year_line_is_recognised_in_every_spelling(line):
    """Саме «р. н.» із пробілом стоїть в офіційному зразку."""
    assert generator.is_biographical_paragraph(line)


def test_registration_line_is_not_biographical():
    """«Підлягає направленню…» — не початок біографічного блоку.

    Інакше обовʼязковий порожній абзац ставав ПІСЛЯ неї, перед ІПН.
    """
    assert not generator.is_biographical_paragraph(
        "Підлягає направленню на військовий облік до Тестового обласного "
        "територіального центру комплектування та соціальної підтримки."
    )


@pytest.mark.parametrize(
    ("line", "expected"),
    [
        (HEAD_LINE_2, False),
        ("ремонтно-відновлювального полку ЗВІЛЬНИТИ:", False),
        ("і ПРИЗНАЧИТИ до цього самого полку:", False),
        (SUBHEADING, True),
        ("ПО ОСОБОВОМУ СКЛАДУ:", True),
        (HEAD_LINE_1, True),
        ("§ 2", True),
    ],
)
def test_new_heading_is_told_apart_from_a_wrapped_fragment(line, expected):
    assert _starts_a_new_heading(line) is expected


# ── Підшапка-підстава, що закінчується КРАПКОЮ, а не двокрапкою ──────────────
_SUB_COLON = "У ЗАПАС ЗА ПІДПУНКТОМ “б” (за станом здоровʼя – за наявності інвалідності):"
_SUB_PERIOD = (
    "У ВІДСТАВКУ ЗА ПІДПУНКТОМ “б” (за станом здоровʼя – на підставі висновку "
    "(постанови) військово-лікарської комісії про непридатність до військової служби)."
)
_DISCHARGE_ITEM = (
    "7. Капітана ТЕСТЕНКА Тараса Тарасовича, інженера 1 окремого тестового загону."
)


def _routes_for(subheadings):
    order = (
        "§ 1\n"
        + HEAD_LINE_1 + "\n"
        + HEAD_LINE_2 + "\n\n"
        + "\n\n".join(subheadings) + "\n\n"
        + _DISCHARGE_ITEM + "\nНародився 10 січня 1973 року.\n987654321.\n"
        + SIGNER
    )
    routes = map_military_units(text=order, mapping=MAPPING)
    entry = next(iter(routes["unit_paragraphs"].values()))
    return entry["items"][0]


@pytest.mark.parametrize(
    "subheading",
    [_SUB_COLON, _SUB_PERIOD],
    ids=["закінчується-двокрапкою", "закінчується-крапкою"],
)
def test_discharge_subheading_is_a_heading_however_it_ends(subheading):
    """Підстава звільнення не завжди закінчується двокрапкою.

    «У ВІДСТАВКУ ЗА ПІДПУНКТОМ “б” (… до військової служби).» закінчується
    КРАПКОЮ. Такий рядок не впізнавався як підшапка, поглинався сусіднім
    блоком і друкувався у витягу не на своєму місці — «підшапка переміщалася
    вниз», а поруч із нею лишалася чужа, вже неактуальна підстава.
    """
    item = _routes_for([subheading])

    assert len(item["heading_ranges"]) == 3, item["heading_ranges"]
    assert subheading[:24] in item["parent_heading"]


def test_neighbouring_discharge_subheadings_replace_each_other():
    """Дві підстави поспіль — у витяг іде ЛИШЕ остання.

    Інакше у витягу стояли обидві, і незрозуміло, за яким саме підпунктом
    звільнено. Друга підстава тут закінчується крапкою — саме той випадок,
    що ламався.
    """
    item = _routes_for([_SUB_COLON, _SUB_PERIOD])
    heading = item["parent_heading"]

    assert "У ЗАПАС" not in heading, heading
    assert "У ВІДСТАВКУ" in heading
    assert len(item["heading_ranges"]) == 3, item["heading_ranges"]


# ── Підстава НАСТУПНОЇ групи не має приліплюватися до попереднього пункту ────
_NEXT_BASIS_SPLIT = [
    "У ЗАПАС ЗА ПІДПУНКТОМ “г” (через сімейні обставини або з інших поважних",
    "причин, перелік яких визначається частиною дванадцятою цієї статті (якщо",
    "військовослужбовці не висловили бажання продовжувати військову службу) –",
    "перебування на утриманні у військовослужбовця трьох і більше дітей):",
]


def test_next_group_basis_does_not_leak_into_the_previous_item():
    """Підстава наступної групи стоїть ПІСЛЯ пункту й розірвана переносом.

    Її перший уламок не закінчується двокрапкою, тому раніше він не вважався
    підшапкою, поглинався попереднім пунктом — і у витягу під пунктом стояла
    ЧУЖА, вже наступна підстава.
    """
    order = (
        "§ 2\n"
        + HEAD_LINE_1 + "\n"
        + HEAD_LINE_2 + "\n\n"
        + SUBHEADING + "\n\n"
        + "13. Підполковника ТЕСТЕНКА Тараса Тарасовича, начальника четвертого "
          "відділу 1 окремого тестового загону.\n"
          "Народився 19 червня 1991 року.\n987654321.\n\n"
        + "\n".join(_NEXT_BASIS_SPLIT) + "\n\n"
        + "14. Майора ІНШЕНКА Івана Івановича, офіцера 1 окремого тестового загону.\n"
        + SIGNER
    )
    routes = map_military_units(text=order, mapping=MAPPING)
    items = {
        str(item["label"]).strip(): item
        for entry in routes["unit_paragraphs"].values()
        for item in entry["items"]
    }

    thirteenth = "\n".join(items["Пункт 13."].get("lines", []))
    assert "ПІДПУНКТОМ “г”" not in thirteenth, thirteenth
    assert "ПІДПУНКТОМ “б”" in items["Пункт 13."]["parent_heading"]
    assert "ПІДПУНКТОМ “г”" in items["Пункт 14."]["parent_heading"]
