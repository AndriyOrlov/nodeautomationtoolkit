"""Біографія пункту не обривається на рядку, схожому на шапку.

Рядок без номера вважався новою шапкою розділу, якщо в ньому є «відповідно
до», «згідно з», «військовослужбовців» або він закінчується двокрапкою —
навіть посеред біографії пункту. Пункт обривався на такому рядку: у витяг
ішла лише частина біографії, а хвіст ставав «шапкою» НАСТУПНОГО пункту й
потрапляв у чужий витяг. Залежало від формулювання — тому «рандомно».

Два запобіжники:
- ВІДСТУП: абзац біографії має лівий відступ від 4 см (правило з перевірки
  наказу). `read_document_text` збирає такі рядки в `bio_lines`, і всередині
  пункту вони ніколи не стають шапкою.
- ТЕКСТ (коли відступів немає): посеред пункту шапкою є лише рядок, що сам її
  починає (§, «Відповідно до …», «Згідно з …», ВЕЛИКІ літери) або закінчується
  двокрапкою.

Дані вигадані.
"""

import importlib.util
import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from nodeautomationtoolkit.builtin_nodes.recipient_mapping import map_military_units  # noqa: E402

UNIT = "військова частина А1000"
MAPPING = {"100 окрема бригада": UNIT}

ORDER = """НАКАЗ

§ 1

Відповідно до пункту 1 Положення нижчепойменованих військовослужбовців ЗВІЛЬНИТИ з займаних посад і ПРИЗНАЧИТИ:

1. Капітана ТЕСТОВОГО Тест Тестовича, командира роти 100 окремої бригади - КОМАНДИРОМ БАТАЛЬЙОНУ 100 ОКРЕМОЇ БРИГАДИ, ВОС - 0000000.
1990 р. н., освіта: ТВІ у 2012 р.,
у ЗС - із 08.2008.
{bio}
Останній рядок біографії.

{next_heading}2. Лейтенанта ІНШОГО Іна Іновича, командира взводу 100 окремої бригади - КОМАНДИРОМ РОТИ 100 ОКРЕМОЇ БРИГАДИ.
"""


def _items(bio: str, next_heading: str = "", with_indents: bool = False):
    text = ORDER.format(bio=bio, next_heading=next_heading)
    lines = text.splitlines()
    first_bio_line = next(i for i, line in enumerate(lines) if "р. н." in line)
    last_bio_line = next(i for i, line in enumerate(lines) if line.startswith("Останній"))
    bio_lines = set(range(first_bio_line, last_bio_line + 1)) if with_indents else None
    routes = map_military_units(text=text, mapping=MAPPING, bio_lines=bio_lines)
    return routes["unit_paragraphs"][UNIT]["items"], last_bio_line


def _assert_bio_kept(items, last_bio_line):
    first, second = items
    assert first["source_end_line"] >= last_bio_line
    # Хвіст біографії першої особи не стає шапкою другого пункту.
    assert "Останній рядок біографії" not in second["parent_heading"]


@pytest.mark.parametrize(
    "bio",
    [
        "Призначається згідно з планом переміщення.",
        "Призначається на посаду відповідно до штату.",
        "Із числа військовослужбовців за контрактом.",
    ],
)
def test_heading_words_inside_biography_do_not_end_the_item(bio):
    _assert_bio_kept(*_items(bio))


@pytest.mark.parametrize(
    "bio",
    [
        "Призначається згідно з планом переміщення.",
        "Вислуга років у ЗС:",
        "Звільняється з правом носіння військової форми одягу згідно з наказом:",
        "У ЗАПАС ЗА ПІДПУНКТОМ «а» як виняток.",
        "ВОС - 0000000.",
    ],
)
def test_indented_biography_is_never_a_heading(bio):
    # Двокрапку в кінці рядка чи ВЕЛИКІ літери текстом від шапки не відрізнити —
    # тут вирішує лише відступ абзацу.
    _assert_bio_kept(*_items(bio, with_indents=True))


@pytest.mark.parametrize("with_indents", [False, True])
@pytest.mark.parametrize(
    "heading",
    [
        "§ 2",
        "Відповідно до пункту 2 Положення нижчепойменованих ПРИЗНАЧИТИ:",
        "Згідно з рапортом ПРИЗНАЧИТИ:",
        "У ЗАПАС ЗА ПІДПУНКТОМ «а» (у зв'язку із закінченням строку контракту):",
        "Нижчепойменованих офіцерів ПРИЗНАЧИТИ:",
    ],
)
def test_real_heading_after_an_item_still_starts_a_new_group(heading, with_indents):
    items, last_bio_line = _items(
        "Призначається на вищу посаду.", next_heading=f"{heading}\n\n", with_indents=with_indents
    )
    first, second = items
    assert first["source_end_line"] < last_bio_line + 3
    assert heading.split()[0] in second["parent_heading"]


# ── Збирання відступів із Word ───────────────────────────────────────────────

def _load_generator():
    spec = importlib.util.spec_from_file_location(
        "generate_extracts_bio_indent_tests", PROJECT_ROOT / "generate_extracts.py"
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class _Range:
    def __init__(self, text):
        self.Text = text
        self.Start = 0


class _Paragraph:
    def __init__(self, text, left_indent):
        self.Range = _Range(text)
        self.LeftIndent = left_indent


class _Paragraphs:
    def __init__(self, items):
        self._items = items
        self.Count = len(items)

    def __call__(self, index):
        return self._items[index - 1]

    def __iter__(self):
        return iter(self._items)


class _Doc:
    def __init__(self, items):
        self.Paragraphs = _Paragraphs(items)


def test_read_document_text_collects_indented_lines_with_soft_breaks():
    generator = _load_generator()
    cm = 72 / 2.54
    doc = _Doc(
        [
            _Paragraph("1. Капітана ТЕСТОВОГО, командира роти.\r", 0),
            # мʼякий перенос усередині абзацу біографії дає ДВА рядки
            _Paragraph("1990 р. н., освіта: ТВІ,\vу ЗС - із 08.2008.\r", 6 * cm),
            _Paragraph("Призначається згідно з планом.\r", 4 * cm),
            _Paragraph("\r", 8 * cm),  # порожній абзац — не біографія
            _Paragraph("По тестовій частині:\r", 3.9 * cm),
            _Paragraph("Мішаний абзац\r", 9999999),
        ]
    )
    bio_lines: set[int] = set()
    lines = generator.read_document_text(doc, bio_lines=bio_lines).splitlines()

    assert sorted(bio_lines) == [1, 2, 3]
    assert [lines[i] for i in sorted(bio_lines)] == [
        "1990 р. н., освіта: ТВІ,",
        "у ЗС - із 08.2008.",
        "Призначається згідно з планом.",
    ]
