"""Ядро генератора наказів: розбір плану, джерела й поєднання їх в один запис."""

from nodeautomationtoolkit.order_generator import merge, sources
from nodeautomationtoolkit.order_generator.names import FullName, accusative, split_full_name
from nodeautomationtoolkit.order_generator.plan import (
    PlanEntry,
    parse_biography,
    parse_candidate,
    parse_vacancy,
)
from nodeautomationtoolkit.order_generator.record import MANUAL, ORDER, PLAN
from nodeautomationtoolkit.order_index.store import PersonItem
from nodeautomationtoolkit.order_index.word_reader import WordTextReader


class _NoWordReader(WordTextReader):
    """У тестах Word не запускається: .doc із вмістом .docx читається напряму."""

    def _read_with_word(self, copy):
        raise OSError("Word у тестах не запускається")


def _write_docx(path, paragraphs):
    from docx import Document

    document = Document()
    for text in paragraphs:
        document.add_paragraph(text)
    document.save(path)


# Рядки за зразком додатка 16 до Інструкції (наказ МО № 170); дані вигадані.
VACANCY = (
    "Начальник штабу - перший заступник командира 901 механізованого батальйону "
    "(шпк «підполковник», ВОС-0210003, 30), вакантна з 20.10.2025"
)
CANDIDATE = (
    "Майор (19.01.2024) ІВАНЕНКО Олексій Вікторович, 1234567890, командир механізованої роти "
    "901 механізованого батальйону з 26.06.2023 (шпк «майор», ВОС-0210003, 25)"
)
BIOGRAPHY = "11.07.1988, Національна академія сухопутних військ (отр) у 2010 р., у ЗС - із 08.2006."

PREVIOUS_ITEM = PersonItem(
    id=1,
    rank="Капітана",
    surname="ІВАНЕНКА",
    name="Олексія",
    patronymic="Вікторовича",
    ipn="1234567890",
    current_position="командира механізованого взводу 901 механізованого батальйону",
    target_position="КОМАНДИРОМ МЕХАНІЗОВАНОЇ РОТИ 901 МЕХАНІЗОВАНОГО БАТАЛЬЙОНУ",
    section="§ 1",
    label="",
    text=(
        "Капітана ІВАНЕНКА Олексія Вікторовича, командира механізованого взводу 901 механізованого "
        "батальйону - КОМАНДИРОМ МЕХАНІЗОВАНОЇ РОТИ 901 МЕХАНІЗОВАНОГО БАТАЛЬЙОНУ, ВОС - 0210003.\n"
        "1988 р.н., освіта: Національна академія сухопутних військ (отр) у 2010 р.,\n"
        "у ЗС - із 08.2006.\n1234567890.\nКандидат технічних наук."
    ),
    order_number="12",
    order_date="2024-03-05",
    path="C:/накази/Наказ №12 від 05.03.2024.docx",
)


def _plan_entry() -> PlanEntry:
    candidate, problems = parse_candidate(CANDIDATE)
    return PlanEntry(
        number="1",
        vacancy=parse_vacancy(VACANCY),
        candidate=candidate,
        biography=parse_biography(BIOGRAPHY),
        evaluation="Займаній посаді відповідає. Гідний до просування по службі. 41,5",
        basis="Пункт 16 Резерву командувача Сухопутних військ Збройних Сил України",
        raw=(),
        problems=problems,
    )


def test_plan_row_is_parsed():
    entry = _plan_entry()
    assert entry.vacancy.position == "Начальник штабу - перший заступник командира 901 механізованого батальйону"
    assert (entry.vacancy.params.shpk, entry.vacancy.params.vos, entry.vacancy.params.tariff) == (
        "підполковник", "0210003", "30",
    )
    assert entry.vacancy.note.startswith("вакантна з 20.10.2025")
    candidate = entry.candidate
    assert (candidate.rank, candidate.rank_date, candidate.ipn) == ("Майор", "19.01.2024", "1234567890")
    assert (candidate.surname, candidate.name, candidate.patronymic) == ("ІВАНЕНКО", "Олексій", "Вікторович")
    assert candidate.position == "командир механізованої роти 901 механізованого батальйону"
    assert candidate.since == "26.06.2023"
    assert (candidate.params.shpk, candidate.params.tariff) == ("майор", "25")
    assert entry.biography.birth_year == "1988"
    assert entry.biography.service_since == "08.2006"
    assert entry.biography.education.startswith("Національна академія сухопутних військ")
    assert entry.problems == []


def test_names_accusative():
    assert accusative(split_full_name("ІВАНЕНКО Олексій Вікторович")) == "ІВАНЕНКА Олексія Вікторовича"
    assert accusative(split_full_name("КОВАЛЬ Ігор Петрович")) == "КОВАЛЯ Ігоря Петровича"
    assert accusative(split_full_name("ЖОВТИЙ Павло Ілліч")) == "ЖОВТОГО Павла Ілліча"
    assert accusative(FullName("ЗІНЧЕНКО", "Світлана", "Петрівна")) == "ЗІНЧЕНКО Світлану Петрівну"
    assert accusative(FullName("КОВАЛЬЧУК", "Олена", "Іванівна")) == "КОВАЛЬЧУК Олену Іванівну"


def test_plan_and_previous_order_are_combined():
    record = merge.build_record(
        plan=sources.from_plan(_plan_entry()),
        order=sources.from_order_item(PREVIOUS_ITEM),
    )
    # Посада, на яку призначили торік, стала посадою, з якої призначають тепер.
    assert str(record.current.text) == "КОМАНДИРОМ МЕХАНІЗОВАНОЇ РОТИ 901 МЕХАНІЗОВАНОГО БАТАЛЬЙОНУ"
    assert record.current.text.source == ORDER
    # Реквізити займаної посади в наказі не пишуть — добираються з плану.
    assert (str(record.current.shpk), str(record.current.tariff)) == ("майор", "25")
    assert record.current.shpk.source == PLAN
    # Нова посада — з плану.
    assert str(record.target.text).startswith("Начальник штабу - перший заступник командира")
    assert record.target.text.source == PLAN
    # Біографія — з наказу, бо там вона вже вивірена; звання — з плану, воно свіжіше.
    assert (str(record.birth), record.birth.source) == ("1988", ORDER)
    assert str(record.extra_bio) == "Кандидат технічних наук."
    assert (str(record.rank), record.rank.source) == ("Майор", PLAN)
    assert str(record.basis).startswith("Пункт 16 Резерву")
    assert any("Наказ №12" in note or "наказ № 12" in note for note in record.notes)
    assert record.problems == []
    assert record.sources()["target.text"] == PLAN


def test_plan_without_previous_order():
    record = merge.build_record(plan=sources.from_plan(_plan_entry()))
    assert str(record.current.text) == "командир механізованої роти 901 механізованого батальйону"
    assert record.current.text.source == PLAN
    assert (str(record.birth), record.birth.source) == ("1988", PLAN)
    assert record.problems == []


def test_order_and_manual_new_position():
    record = merge.build_record(
        order=sources.from_order_item(PREVIOUS_ITEM),
        manual=sources.from_manual({
            "target.text": "начальник штабу 902 механізованого батальйону",
            "target.shpk": "підполковник",
            "basis": "рапорт від 01.09.2026",
        }),
    )
    assert str(record.target.text) == "начальник штабу 902 механізованого батальйону"
    assert record.target.text.source == MANUAL
    assert str(record.current.text).startswith("КОМАНДИРОМ МЕХАНІЗОВАНОЇ РОТИ")
    assert str(record.basis) == "рапорт від 01.09.2026"
    # Єдине зауваження: ПІБ прийшло з пункту наказу, а там воно у знахідному.
    assert len(record.problems) == 1 and "знахідному відмінку" in record.problems[0]


def test_conflict_between_plan_and_order_is_reported():
    plan = sources.from_plan(_plan_entry())
    other_item = PersonItem(**{**PREVIOUS_ITEM.__dict__,
                               "target_position": "КОМАНДИРОМ ТАНКОВОЇ РОТИ 902 ТАНКОВОГО БАТАЛЬЙОНУ"})
    record = merge.build_record(plan=plan, order=sources.from_order_item(other_item))
    assert any("не збігається" in problem for problem in record.problems)


def test_manual_only_reports_what_is_missing():
    record = merge.build_record(manual=sources.from_manual({"surname": "ІВАНЕНКО"}))
    assert "немає посади, на яку призначається" in record.problems
    assert "немає РНОКПП" in record.problems


# ── Будь-який документ на вході ──────────────────────────────────────────────
PODANNIA = """ПОДАННЯ
1. Майор ІВАНЕНКО Олексій Вікторович, командир механізованої роти 901 механізованого батальйону,
призначений наказом командувача Сухопутних військ від 05.03.2024 № 12, шпк «майор», ВОС-0210003, 25 т.р.
2. Реєстраційний номер облікової картки платника податків 1234567890, 11.07.1988 р.н.,
освіта: Національна академія сухопутних військ (отр) у 2010 р., у ЗС - із 08.2006.
3. Подається до призначення на посаду: начальник штабу - перший заступник командира
901 механізованого батальйону, шпк «підполковник», ВОС-0210003, 30 т.р.
На призначення на зазначену посаду згоден. Підстава: рапорт від 01.09.2026.
"""

RAPORT = """Рапорт (Заява)
Командиру військової частини А1111
Повідомляю, що маю право на звільнення з військової служби та бажаю звільнитися з військової служби
у запас відповідно до підпункту «б» пункту 2 частини шостої статті 26 Закону України
«Про військовий обов'язок і військову службу».
Капітан ПЕТРЕНКО Сергій Миколайович, командир механізованого взводу, 2345678901, 03.02.1992 р.н.
"""


def _write(tmp_path, name, text):
    path = tmp_path / name
    path.write_text(text, encoding="utf-8")
    return path


def test_any_document_is_read_softly(tmp_path):
    from nodeautomationtoolkit.order_generator.documents import read_document

    facts = read_document(_write(tmp_path, "подання.txt", PODANNIA))
    assert facts.kind == "подання"
    assert [person.full_name for person in facts.people] == ["ІВАНЕНКО Олексій Вікторович"]
    assert facts.first("ipn") == "1234567890"
    assert facts.first("birth") == "11.07.1988"
    assert facts.first("service_since") == "08.2006"
    assert facts.first("shpk") == "майор"
    assert facts.first("vos") == "0210003"
    assert "згоден" in facts.first("consent").casefold()
    assert facts.first("order_reference") == "наказ від 05.03.2024 № 12"
    assert facts.first("basis") == "рапорт від 01.09.2026"
    assert "командир роти" in facts.positions

    report = read_document(_write(tmp_path, "рапорт.txt", RAPORT))
    assert report.kind == "рапорт"
    assert report.first("ipn") == "2345678901"
    assert "звільнитися" in report.first("discharge").casefold()
    assert report.first("law_reference").startswith("підпункт")


def test_document_is_checked_against_orders(tmp_path, monkeypatch):
    from nodeautomationtoolkit.order_generator.resolve import resolve_document
    from nodeautomationtoolkit.order_index import store as index_store

    monkeypatch.setattr(index_store, "WordTextReader", _NoWordReader)
    orders, index = tmp_path / "orders", tmp_path / "index"
    orders.mkdir()
    _write_docx(orders / "Наказ №12 від 05.03.2024.docx", [
        "§ 1",
        "Капітана ІВАНЕНКА Олексія Вікторовича, командира механізованого взводу 901 механізованого "
        "батальйону - КОМАНДИРОМ МЕХАНІЗОВАНОЇ РОТИ 901 МЕХАНІЗОВАНОГО БАТАЛЬЙОНУ, ВОС - 0210003.",
        "1988 р.н., освіта: Національна академія сухопутних військ (отр) у 2010 р., у ЗС - із 08.2006.",
        "1234567890.",
    ])
    index_store.build_index(orders, index)

    resolved = resolve_document(_write(tmp_path, "подання.txt", PODANNIA), index)
    assert resolved.kind == "подання"
    person = resolved.people[0]
    record = person.record
    # Біографія — з наказу, посада «з якої» — та, на яку призначили торік.
    assert str(record.birth) == "1988" and record.birth.source == "наказ"
    assert str(record.current.text).startswith("КОМАНДИРОМ МЕХАНІЗОВАНОЇ РОТИ")
    assert str(record.target.text)
    assert any("збігається з наказом" in note for note in record.notes)
    assert any("Посада в документі збігається" in note for note in record.notes)
    assert record.problems == []


def test_document_without_person_in_index(tmp_path, monkeypatch):
    from nodeautomationtoolkit.order_generator.resolve import resolve_document
    from nodeautomationtoolkit.order_index import store as index_store

    monkeypatch.setattr(index_store, "WordTextReader", _NoWordReader)
    orders, index = tmp_path / "orders", tmp_path / "index"
    orders.mkdir()
    _write_docx(orders / "Наказ №1 від 01.01.2024.docx", ["§ 1", "Немає жодного пункту."])
    index_store.build_index(orders, index)
    resolved = resolve_document(_write(tmp_path, "рапорт.txt", RAPORT), index)
    record = resolved.people[0].record
    assert any("не знайдено" in note for note in record.notes)
    assert str(record.ipn) == "2345678901"
    assert "немає посади, на яку призначається" in record.problems
