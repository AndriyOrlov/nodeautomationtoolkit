"""Наказ про звільнення з військової служби: підпункти, вислуга, ТЦК, форма.

Формулювання звірені зі зразками додатка 53, підстави — зі статті 26 Закону
України «Про військовий обов'язок і військову службу». Дані вигадані.
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from nodeautomationtoolkit.order_generator.build import (  # noqa: E402
    OrderDocumentParts,
    _kind,
    build_order_document,
    content_lines,
)
from nodeautomationtoolkit.order_generator.check import ERROR, check_draft  # noqa: E402
from nodeautomationtoolkit.order_generator.compose import (  # noqa: E402
    DISMISSAL,
    OrderParams,
    birth_in_words,
    compose_order,
)
from nodeautomationtoolkit.order_generator.record import MANUAL, PersonRecord, Position, Value  # noqa: E402
from nodeautomationtoolkit.personnel import dismissal  # noqa: E402

NBSP = " "
LAW_POINTS = "пункту другого частини п'ятої статті 26"


def person(
    surname="БОНДАР",
    name="Руслан",
    patronymic="Володимирович",
    rank="полковник",
    position="заступник начальника факультету",
    ground="1.а",
    destination="у запас",
    calendar="29 років 3 місяці",
    privileged="29 років 11 місяців",
    registration="Галицько-Франківського ОРТЦК та СП м. Львова",
    ipn="2652323639",
    birth="11.08.1976",
    uniform="так",
    note="Чинність контракту припиняється 18.12.2026",
) -> PersonRecord:
    return PersonRecord(
        rank=Value(rank, MANUAL),
        surname=Value(surname, MANUAL),
        name=Value(name, MANUAL),
        patronymic=Value(patronymic, MANUAL),
        ipn=Value(ipn, MANUAL),
        birth=Value(birth, MANUAL),
        current=Position(text=Value(position, MANUAL)),
        dismissal=Value(ground, MANUAL),
        destination=Value(destination, MANUAL),
        service_calendar=Value(calendar, MANUAL),
        service_privileged=Value(privileged, MANUAL),
        registration=Value(registration, MANUAL),
        uniform=Value(uniform, MANUAL),
        dismissal_note=Value(note, MANUAL),
    )


def params(**extra) -> OrderParams:
    values = {"action": DISMISSAL, "law_points": LAW_POINTS, "unit": "Збройних Сил України"}
    values.update(extra)
    return OrderParams(**values)


def plain(text: str) -> str:
    return text.replace(NBSP, " ")


# ─────────────────────────────── довідник ────────────────────────────────────
def test_grounds_come_from_the_dictionary():
    table = dismissal.grounds()
    assert len(table) > 20
    first = dismissal.find_ground("1.а")
    assert first is not None and first.reason == "у зв'язку із закінченням строку контракту"
    # За самим підпунктом теж знаходиться.
    assert dismissal.find_ground("г").reason.startswith("у зв'язку із скороченням штатів")
    assert dismissal.find_ground("немає такої") is None


def test_subheading_matches_the_sample():
    assert dismissal.subheading("1.а") == (
        "У ЗАПАС ЗА ПІДПУНКТОМ «а» (у зв'язку із закінченням строку контракту):"
    )
    assert dismissal.subheading("1.б", "у відставку") == (
        "У ВІДСТАВКУ ЗА ПІДПУНКТОМ «б» (за станом здоров'я):"
    )


# ─────────────────────────────── генерація ───────────────────────────────────
def test_dismissal_item_follows_the_sample_wording():
    draft = compose_order([person()], params())
    lines = [plain(line) for line in draft.text.splitlines()]

    assert lines[0].startswith("Відповідно до пункту другого частини п'ятої статті 26 Закону України")
    assert lines[0].endswith("ЗВІЛЬНИТИ з військової служби:")
    assert lines[2] == "У ЗАПАС ЗА ПІДПУНКТОМ «а» (у зв'язку із закінченням строку контракту):"
    assert lines[4] == (
        "1. Полковника БОНДАРА Руслана Володимировича, заступника начальника факультету."
    )
    assert lines[5] == (
        "Народився 11 серпня 1976 року. Вислуга років у ЗС: календарна - 29 років 3 місяці, "
        "пільгова - 29 років 11 місяців."
    )
    assert lines[6] == (
        "Підлягає направленню на військовий облік до Галицько-Франківського ОРТЦК та СП м. Львова."
    )
    assert lines[7] == "2652323639."
    assert lines[8] == "Звільняється з правом носіння військової форми одягу."
    assert lines[9] == "Чинність контракту припиняється 18.12.2026."


def test_birth_in_words():
    assert birth_in_words("11.08.1976") == "Народився 11 серпня 1976 року"
    assert birth_in_words("1976") == "1976 р.н."
    assert birth_in_words("") == ""


def test_without_the_uniform_right():
    draft = compose_order([person(uniform="ні")], params())
    assert "Звільняється без права носіння військової форми одягу." in plain(draft.text)


def test_people_are_grouped_by_subparagraph_and_numbered_through():
    draft = compose_order(
        [
            person(surname="ЯНЧАК", ground="1.г", ipn=""),
            person(surname="АНТОНЮК", ground="1.б", destination="у відставку", ipn=""),
            person(surname="БОНДАР", ground="1.а", ipn=""),
            person(surname="АЛЬОШИН", ground="1.г", ipn=""),
        ],
        params(),
    )
    assert [item.number for item in draft.items] == [1, 2, 3, 4]
    assert [str(item.record.surname) for item in draft.items] == [
        "БОНДАР",  # «а»
        "АНТОНЮК",  # «б»
        "АЛЬОШИН",  # «г», за абеткою
        "ЯНЧАК",
    ]
    text = plain(draft.text)
    assert text.count("ЗА ПІДПУНКТОМ") == 3
    assert text.index("«а»") < text.index("«б»") < text.index("«г»")


def test_dismissal_has_no_appointment_wording():
    text = plain(compose_order([person()], params()).text)
    assert "Призначається" not in text
    assert "ПРИЗНАЧИТИ" not in text


def test_non_breaking_spaces_are_applied_here_too():
    text = compose_order([person()], params()).text
    assert f"у{NBSP}ЗС" in text
    assert f"до{NBSP}Галицько-Франківського" in text


# ──────────────────────────────── перевірка ──────────────────────────────────
def test_check_passes_on_a_complete_dismissal():
    draft = compose_order([person()], params(number="530", date="18.09.2026"))
    result = check_draft(draft)
    assert result.ready, [problem.line() for problem in result.errors]


def test_check_wants_the_subparagraph():
    result = check_draft(compose_order([person(ground="")], params()))
    assert not result.ready
    assert any("підпункт" in problem.what for problem in result.errors)


def test_check_reports_an_unknown_subparagraph():
    result = check_draft(compose_order([person(ground="щ")], params()))
    assert any(
        problem.level == ERROR and "довіднику підстав" in problem.what
        for problem in result.problems
    )


def test_check_does_not_ask_for_a_new_position():
    result = check_draft(compose_order([person()], params()))
    assert not any("на яку призначається" in problem.what for problem in result.problems)


def test_check_warns_about_the_missing_registration():
    result = check_draft(compose_order([person(registration="")], params()))
    assert result.ready
    assert any("ТЦК" in problem.what for problem in result.warnings)


# ───────────────────────────────── збірка ────────────────────────────────────
def test_subheading_is_kept_with_the_first_item(tmp_path):
    from docx import Document

    draft = compose_order(
        [person(), person(surname="АНТОНЮК", ground="1.б", destination="у відставку", ipn="")],
        params(),
    )
    lines = content_lines(draft, OrderDocumentParts())
    subheading = next(index for index, line in enumerate(lines) if "ЗА ПІДПУНКТОМ" in line)
    assert _kind(lines[subheading]) == "heading"

    path = build_order_document(draft, tmp_path / "dismissal.docx")
    paragraphs = Document(str(path)).paragraphs
    heading = next(p for p in paragraphs if "ЗА ПІДПУНКТОМ" in p.text)
    assert heading.paragraph_format.keep_with_next is True
