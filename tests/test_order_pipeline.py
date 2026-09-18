"""Чотири етапи генератора наказів разом: індексація, генерація, перевірка, збірка.

Усе на вигаданих файлах у тимчасовій теці (PROJECT_RULES 1.1).
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest
from docx import Document
from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from nodeautomationtoolkit.order_generator import pipeline  # noqa: E402
from nodeautomationtoolkit.order_generator.build import OrderDocumentParts  # noqa: E402
from nodeautomationtoolkit.order_generator.compose import OrderParams  # noqa: E402

PLAN_HEADER = [
    "№ з/п",
    "Найменування посади, що підлягає комплектуванню",
    "Військове звання, прізвище, ім'я та по батькові кандидата",
    "Дата народження, освіта",
    "Оцінювання",
    "Підстава включення до плану",
    "Відмітка про реалізацію",
]

PLAN_ROWS = [
    [
        "1",
        "командир механізованої роти 901 механізованого батальйону "
        "(шпк «майор», ВОС-0210003, 12 т.р.), вакантна з 01.08.2026",
        "капітан (12.05.2023) ІВАНЕНКО Олексій Вікторович, 3508402997, "
        "командир механізованого взводу 901 механізованого батальйону, з 01.09.2024 "
        "(шпк «капітан», ВОС-0210003, 10 т.р.)",
        "11.07.1988, Національна академія сухопутних військ (отр) у 2010 р., у ЗС - із 08.2006",
        "відповідає займаній посаді, 85 балів",
        "план переміщення від 01.09.2026",
        "",
    ],
    [
        "2",
        "начальник штабу 901 механізованого батальйону (шпк «підполковник», ВОС-0210003, 14 т.р.), "
        "вакантна з 01.09.2026",
        "майор (03.02.2022) БОНДАРЕНКО Сергій Петрович, 2649423014, "
        "заступник командира механізованого батальйону, з 01.03.2023 "
        "(шпк «майор», ВОС-0210003, 13 т.р.)",
        "02.03.1985, Одеський ІСВ у 2007 р., у ЗС - із 08.2003",
        "відповідає займаній посаді, 90 балів",
        "план переміщення від 01.09.2026",
        "",
    ],
]


@pytest.fixture
def plan_file(tmp_path) -> Path:
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["ЗАТВЕРДЖУЮ"])
    sheet.append(["ПЛАН переміщення військовослужбовців"])
    sheet.append(PLAN_HEADER)
    for row in PLAN_ROWS:
        sheet.append(row)
    path = tmp_path / "План переміщення.xlsx"
    workbook.save(path)
    return path


def _write_order(path: Path, paragraphs: list[str]) -> None:
    document = Document()
    for text in paragraphs:
        document.add_paragraph(text)
    path.parent.mkdir(parents=True, exist_ok=True)
    document.save(str(path))


def test_plan_alone_makes_an_order(plan_file, tmp_path):
    records = pipeline.records_from_plan(plan_file)
    assert [record.full_name for record in records] == [
        "ІВАНЕНКО Олексій Вікторович",
        "БОНДАРЕНКО Сергій Петрович",
    ]

    result = pipeline.run(
        records,
        OrderParams(points="пункту 45", number="525", date="17.09.2026", section="§ 1"),
        tmp_path / "готове",
        parts=OrderDocumentParts(
            signer_position="Командувач Сухопутних військ Збройних Сил України",
            signer_rank="генерал-лейтенант",
            signer_name="І. ПЕТРЕНКО",
            executor="Виконавець: О. КОВАЛЬ",
        ),
    )
    assert result.document is not None and result.document.exists()
    assert result.document.name == "Наказ № 525 від 17.09.2026.docx"

    text = "\n".join(p.text for p in Document(str(result.document)).paragraphs).replace(" ", " ")
    # Прізвища за абеткою, наскрізна нумерація.
    assert text.index("БОНДАРЕНКА") < text.index("ІВАНЕНКА")
    assert "1. Майора БОНДАРЕНКА Сергія Петровича" in text
    assert "2. Капітана ІВАНЕНКА Олексія Вікторовича" in text
    assert "КОМАНДИРОМ МЕХАНІЗОВАНОЇ РОТИ" in text
    assert "Призначається на вищу посаду" in text
    assert text.strip().endswith("Виконавець: О. КОВАЛЬ")


def test_index_supplies_the_position_a_person_holds(plan_file, tmp_path, monkeypatch):
    from nodeautomationtoolkit.order_index import store as index_store
    from nodeautomationtoolkit.order_index.word_reader import WordTextReader

    class _NoWord(WordTextReader):
        """У тестах Word не запускається; .docx читається напряму."""

        def _read_with_word(self, copy):
            raise OSError("Word у тестах не запускається")

    monkeypatch.setattr(index_store, "WordTextReader", _NoWord)
    orders, index = tmp_path / "накази", tmp_path / "індекс"
    _write_order(
        orders / "Наказ №12 від 05.03.2024.docx",
        [
            "§ 1",
            "Капітана ІВАНЕНКА Олексія Вікторовича, командира механізованого взводу 901 "
            "механізованого батальйону - КОМАНДИРОМ МЕХАНІЗОВАНОЇ РОТИ 901 МЕХАНІЗОВАНОГО "
            "БАТАЛЬЙОНУ, ВОС - 0210003.",
            "1988 р.н., освіта: Національна академія сухопутних військ (отр) у 2010 р., "
            "у ЗС - із 08.2006.",
            "3508402997.",
        ],
    )
    summary = pipeline.index_orders(orders, index)
    assert summary.indexed == 1 and summary.items >= 1

    records = pipeline.records_from_plan(plan_file, index)
    ivanenko = next(record for record in records if str(record.surname) == "ІВАНЕНКО")
    # Посада, НА ЯКУ призначили торік, стає посадою, З ЯКОЇ призначають тепер.
    assert str(ivanenko.current.text).startswith("КОМАНДИРОМ МЕХАНІЗОВАНОЇ РОТИ")
    assert ivanenko.current.text.source == "наказ"
    assert ivanenko.birth.source == "наказ"
    assert any("Пункт узято з" in note for note in ivanenko.notes)


def test_errors_stop_the_assembly(tmp_path):
    from nodeautomationtoolkit.order_generator.record import MANUAL, PersonRecord, Value

    record = PersonRecord(
        rank=Value("капітан", MANUAL),
        surname=Value("ІВАНЕНКО", MANUAL),
        name=Value("Олексій", MANUAL),
        patronymic=Value("Вікторович", MANUAL),
    )
    messages: list[str] = []
    result = pipeline.run(
        [record], OrderParams(), tmp_path / "готове", log=messages.append
    )
    assert result.document is None
    assert not (tmp_path / "готове").exists()
    assert any("виправте помилки" in message for message in messages)
