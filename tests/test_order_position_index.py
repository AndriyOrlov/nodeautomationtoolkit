"""Тестовий індексатор посад: вигадані накази за структурою додатка 53."""

import os
import shutil
import time

from docx import Document

from nodeautomationtoolkit.order_index import store
from nodeautomationtoolkit.order_index.extractor import (
    ROLE_CURRENT,
    ROLE_TARGET,
    extract_positions,
    order_metadata,
    position_key,
)
from nodeautomationtoolkit.order_index.position_dictionary import load_position_dictionary
from nodeautomationtoolkit.order_index.unit_split import split_position_unit
from nodeautomationtoolkit.order_index.word_reader import (
    WordTextReader,
    sniff_suffix,
    split_word_text,
)

APPOINT = [
    "§ 1",
    "Відповідно до пункту 5 Положення нижчепойменованих осіб офіцерського складу ЗВІЛЬНИТИ з займаних посад і ПРИЗНАЧИТИ:",
    "Капітана ІВАНОВА Івана Івановича, командира механізованої роти механізованого батальйону військової частини А1111 - "
    "КОМАНДИРОМ МЕХАНІЗОВАНОГО БАТАЛЬЙОНУ ВІЙСЬКОВОЇ ЧАСТИНИ А2222, ВОС - 0211003.",
    "1990 р.н., освіта: вигаданий університет у 2012 р., у ЗС - із 08.2008.",
    "1234567890.",
    "Старшого лейтенанта СИДОРОВА Петра",
    "Петровича, заступника командира роти військової частини А1111 - КОМАНДИРОМ РОТИ ВІЙСЬКОВОЇ ЧАСТИНИ А1111.",
    "§ 2",
    "Відповідно до статті 26 Закону капітана ПЕТРЕНКА Олега Олеговича, командира механізованої роти механізованого "
    "батальйону військової частини А1111, ЗВІЛЬНИТИ з займаної посади і ПРИЗНАЧИТИ ЗАСТУПНИКОМ КОМАНДИРА БАТАЛЬЙОНУ "
    "ВІЙСЬКОВОЇ ЧАСТИНИ А2222, ВОС - 0211003.",
    "§ 3",
    "Майора КОВАЛЕНКА Андрія Андрійовича, начальника штабу військової частини А1111, який перебуває у розпорядженні, "
    "ІМЕНУВАТИ КОВАЛЕМ.",
    "Солдата БОНДАРЕНКА Василя Васильовича, стрільця військової частини А1111, ЗВІЛЬНИТИ З ВІЙСЬКОВОЇ СЛУЖБИ У ЗАПАС.",
]


def _found(hits, role):
    return [(hit.position, hit.known) for hit in hits if hit.role == role]


def test_current_and_target_positions():
    hits = extract_positions(APPOINT)
    assert _found(hits, ROLE_CURRENT) == [
        ("командир роти", True),
        ("заступник командира роти", True),
        ("командир роти", True),
        ("начальник штабу військової частини", True),
        ("стрілець", True),
    ]
    assert _found(hits, ROLE_TARGET) == [
        ("командир батальйону", True),
        ("командир роти", True),
        ("заступник командира батальйону", True),
    ]
    # Повний текст із частиною лишається для перегляду.
    assert hits[0].text == "командира механізованої роти механізованого батальйону військової частини А1111"
    assert [hit.section for hit in hits[:3]] == ["§ 1", "§ 1", "§ 1"]


def test_birth_year_after_name_is_not_a_position():
    hits = extract_positions([
        "Полковника ІВАНОВА Івана Івановича, 1946 р.н., старшого офіцера відділу кадрів військової частини А1111.",
        "Майора ПЕТРОВА Петра Петровича, 1954 р.н.",
        "Капітана СИДОРОВА Олега Олеговича, 12.03.1990 р. н., 1234567890, водія-електрика взводу забезпечення.",
    ])
    assert [(hit.position, hit.known) for hit in hits] == [("старший офіцер відділу", True), ("водій-електрик", True)]


def test_biography_and_non_positions_are_ignored():
    hits = extract_positions(["1990 р.н., освіта: вигаданий університет у 2012 р.", "Підстава: рапорт від 01.02.2025."])
    assert hits == []


def test_dictionary_matches_any_case_and_register():
    dictionary = load_position_dictionary()
    cases = {
        "командира механізованої роти військової частини А1111": "командир роти",
        "КОМАНДИРОМ МЕХАНІЗОВАНОЇ РОТИ": "командир роти",
        "командиру 2 механізованої роти": "командир роти",
        "заступника командира 72 окремої механізованої бригади": "заступник командира бригади",
        "заступника командира з морально-психологічного забезпечення батальйону":
            "заступник командира з морально-психологічного забезпечення",
        "стрільця-помічника гранатометника механізованого відділення": "стрілець-помічник гранатометника",
        "СТАРШИМ БОЙОВИМ МЕДИКОМ ВЗВОДУ": "старший бойовий медик",
        "колишньому старшому помічнику військового коменданта": "старший помічник коменданта",
        "командира механізованого відділення": "командир відділення",
    }
    for text, expected in cases.items():
        match = dictionary.match_at_start(text)
        assert match is not None and match.nominative == expected, text
    assert dictionary.match_at_start("1946 р.н.") is None
    assert dictionary.match_at_start("який перебуває у розпорядженні") is None


def test_unknown_position_is_marked():
    hits = extract_positions(["Майора СИДОРОВА Олега Олеговича, фахівця з вигаданих справ військової частини А1111."])
    assert [(hit.position, hit.known) for hit in hits] == [("фахівця з вигаданих справ", False)]


def test_key_and_metadata():
    assert position_key("Командира «роти» – взводу.") == position_key('командира "роти"-взводу')
    assert order_metadata("Наказ №123 від 05.03.2025") == ("123", "2025-03-05")
    assert order_metadata("без реквізитів", "НАКАЗ\n«12» березня 2024 року № 45") == ("45", "2024-03-12")


def _write_order(path, paragraphs):
    document = Document()
    for text in paragraphs:
        document.add_paragraph(text)
    document.save(path)


class _NoWord(WordTextReader):
    """Без справжнього Word: .doc із вмістом .docx читається напряму, решта — помилка."""

    def _read_with_word(self, copy):
        raise OSError("Word у тестах не запускається")


def test_build_index_is_incremental(tmp_path, monkeypatch):
    monkeypatch.setattr(store, "WordTextReader", _NoWord)
    orders = tmp_path / "orders"
    (orders / "2024").mkdir(parents=True)
    index = tmp_path / "index"
    _write_order(orders / "Наказ №10 від 01.02.2024.docx", APPOINT)
    _write_order(orders / "2024" / "Наказ №11 від 03.04.2025.docx", APPOINT[:5])
    # .doc, який насправді .docx, читається без Word.
    shutil.copy(orders / "2024" / "Наказ №11 від 03.04.2025.docx", orders / "Наказ №12 від 05.05.2025.doc")
    (orders / "зіпсований №1 від 01.01.2020.doc").write_bytes(b"\xd0\xcf\x11\xe0 broken")
    (orders / "~$Наказ №10 від 01.02.2024.docx").write_bytes(b"lock")

    first = store.build_index(orders, index)
    assert (first.indexed, first.failed) == (3, 1)

    rows = {row.display: row for row in store.load_positions(index)}
    company = rows["командир роти"]
    # займана: 2 у №10 + 1 у №11 + 1 у №12; призначення: 1 у №10
    assert (company.mentions, company.orders, company.units, company.current, company.target) == (5, 3, 1, 4, 1)
    assert (company.first_date, company.last_date) == ("2024-02-01", "2025-05-05")
    occurrences = store.load_occurrences(index, company.id)
    assert [o.number for o in occurrences] == ["12", "11", "10", "10", "10"]
    assert {o.unit for o in occurrences} == {"військової частини А1111", "ВІЙСЬКОВОЇ ЧАСТИНИ А1111"}
    assert all(row.known for row in rows.values())

    second = store.build_index(orders, index)
    assert (second.indexed, second.unchanged) == (0, 4)

    os.remove(orders / "2024" / "Наказ №11 від 03.04.2025.docx")
    third = store.build_index(orders, index)
    assert third.removed == 1
    rows = {row.display: row for row in store.load_positions(index)}
    assert rows["командир роти"].orders == 2
    stats = store.load_stats(index)
    assert (stats.ok, stats.errors) == (2, 1)


def test_changed_file_is_reindexed(tmp_path):
    orders, index = tmp_path / "orders", tmp_path / "index"
    orders.mkdir()
    order = orders / "Наказ №1 від 01.01.2025.docx"
    _write_order(order, APPOINT[:3])
    store.build_index(orders, index)
    time.sleep(0.05)
    _write_order(order, APPOINT[5:7])
    summary = store.build_index(orders, index)
    assert summary.indexed == 1
    displays = {row.display for row in store.load_positions(index)}
    assert displays == {"заступник командира роти", "командир роти"}


def test_same_position_in_different_units_is_one_row(tmp_path):
    orders, index = tmp_path / "orders", tmp_path / "index"
    orders.mkdir()
    for number, unit in enumerate(("А1111", "А2222", "А3333"), start=1):
        _write_order(
            orders / f"Наказ №{number} від 0{number}.01.2025.docx",
            [f"Капітана ІВАНОВА Івана Івановича, командира {number} механізованої роти військової частини {unit}."],
        )
    store.build_index(orders, index)
    rows = store.load_positions(index)
    assert [(row.display, row.mentions, row.units) for row in rows] == [("командир роти", 3, 3)]


def test_word_text_helpers(tmp_path):
    assert split_word_text("перший\rклітинка\x07\rрядок\x0bпродовження") == [
        "перший", "клітинка", "", "рядок продовження"
    ]
    fake_doc = tmp_path / "наказ.doc"
    _write_order(fake_doc, ["текст"])
    assert sniff_suffix(fake_doc) == ".docx"


def test_split_position_unit():
    cases = {
        "командира механізованої роти механізованого батальйону військової частини А1111":
            ("командира механізованої роти механізованого батальйону", "військової частини А1111"),
        "командира 2 механізованої роти 1 механізованого батальйону військової частини А1111":
            ("командира механізованої роти механізованого батальйону", "військової частини А1111"),
        "заступника командира 72 окремої механізованої бригади":
            ("заступника командира бригади", "72 окремої механізованої бригади"),
        "КОМАНДИРОМ РОТИ ВІЙСЬКОВОЇ ЧАСТИНИ А2222": ("КОМАНДИРОМ РОТИ", "ВІЙСЬКОВОЇ ЧАСТИНИ А2222"),
        "штаб-сержанта 2 категорії служби захисту інформації Кадрового центру":
            ("штаб-сержанта 2 категорії служби захисту інформації", "Кадрового центру"),
        "НАЧАЛЬНИКОМ ГРУПИ ЛОГІСТИКИ (S-4) 00 ОКРЕМОГО БАТАЛЬЙОНУ":
            ("НАЧАЛЬНИКОМ ГРУПИ ЛОГІСТИКИ (S-4)", "00 ОКРЕМОГО БАТАЛЬЙОНУ"),
        "начальника штабу батальйону": ("начальника штабу батальйону", ""),
    }
    for text, expected in cases.items():
        assert split_position_unit(text) == expected, text


def _build_script():
    import importlib.util
    from pathlib import Path

    path = Path(__file__).resolve().parents[1] / "scripts" / "order_index" / "build_mo317_dictionaries.py"
    spec = importlib.util.spec_from_file_location("build_mo317_dictionaries", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_mo317_vos_expansion():
    expand = _build_script().expand_vos
    assert expand("220001-228001, 240001, 279001") == (
        ["220", "221", "222", "223", "224", "225", "226", "227", "228", "240", "279"],
        ["220001", "221001", "222001", "223001", "224001", "225001", "226001", "227001", "228001", "240001", "279001"],
        False,
        [],
    )
    vos, codes, all_vos, excluded = expand("усі ВОС, крім ВОС: 113258, 116258 - 118258, 967258;")
    assert (vos, codes, all_vos, excluded) == ([], [], True, ["113", "116", "117", "118", "967"])
    assert expand("Усі ВОС") == ([], [], True, [])


def test_mo317_dictionary_gives_codes_and_vos():
    from nodeautomationtoolkit.order_index.position_dictionary import (
        load_mo317,
        load_vos_names,
        strip_brackets,
    )

    assert strip_brackets("Льотчик (літака)") == "льотчик"
    assert strip_brackets("Повітряний стрілець - радист") == "повітряний стрілець-радист"
    mo317 = load_mo317()
    driver = mo317["механік-водій"]
    assert {"258", "259"} <= driver.codes
    assert driver.all_vos
    assert load_vos_names()["100"] == "Стрілецькі"
    # Назва з переліку МО, якої немає в таблиці Excel, теж розпізнається в наказі.
    match = load_position_dictionary().match_at_start("водія електронавантажувача автомобільної роти")
    assert match is not None and match.nominative == "водій електронавантажувача"


def test_mo444_officer_vos_dictionary():
    from nodeautomationtoolkit.order_index.position_dictionary import (
        MO444_VOS_REPLACEMENT_FILE,
        load_officer_vos,
    )

    officer_vos = load_officer_vos()
    assert len(officer_vos) > 400
    assert officer_vos["061800"]["Найменування ВОС"] == "Льотчик"
    assert officer_vos["021000"]["Група"] == "02"
    assert "020100" in officer_vos["021000"]["Допустимо без додаткової підготовки"]
    assert all(len(code) == 6 and code.isdigit() for code in officer_vos)
    header = MO444_VOS_REPLACEMENT_FILE.read_text(encoding="utf-8-sig").splitlines()[0]
    assert header == "Код ВОС, що підлягає заміні;Код ВОС, на який здійснюється заміна"


def test_position_forms():
    from nodeautomationtoolkit.order_index.position_forms import all_forms

    assert all_forms("старший офіцер відділу") == {
        "Р": "старшого офіцера відділу", "Д": "старшому офіцеру відділу", "О": "старшим офіцером відділу"
    }
    assert all_forms("стрілець-помічник гранатометника")["Р"] == "стрільця-помічника гранатометника"
    assert all_forms("водій")["О"] == "водієм"
    assert all_forms("головна медична сестра")["Д"] == "головній медичній сестрі"
    assert all_forms("черговий частини")["Р"] == "чергового частини"
    assert all_forms("штаб-сержант")["О"] == "штаб-сержантом"
    assert all_forms("начальник штабу - перший заступник командира батальйону")["Р"] == (
        "начальника штабу - першого заступника командира батальйону"
    )


def test_all_positions_dictionary_is_the_only_source():
    import csv

    from nodeautomationtoolkit.order_index.position_dictionary import ALL_POSITIONS_FILE

    with open(ALL_POSITIONS_FILE, encoding="utf-8-sig", newline="") as handle:
        rows = {row["Називний"]: row for row in csv.DictReader(handle, delimiter=";")}
    assert len(rows) > 1300
    rifleman = rows["стрілець"]
    assert (rifleman["Родовий"], rifleman["Орудний"]) == ("стрільця", "стрільцем")
    assert rifleman["Код посади (МО 317)"] and rifleman["Номери ВОС (МО 317)"]
    assert "офіцерський" in rows["командир роти"]["Склад"]
    dictionary = load_position_dictionary()
    for text, expected in {
        "начальника штабу - першого заступника командира 1 механізованого батальйону": (
            "начальник штабу - перший заступник командира батальйону"
        ),
        "старшого авіаційного техніка групи": "старший авіаційний технік",
        "начальником медичної служби бригади": "начальник медичної служби",
    }.items():
        match = dictionary.match_at_start(text)
        assert match is not None and match.nominative == expected, text


ITEM_ORDER_2024 = [
    "§ 1",
    "Відповідно до пункту 5 Положення нижчепойменованих осіб офіцерського складу ЗВІЛЬНИТИ з займаних посад і ПРИЗНАЧИТИ:",
    "Капітана ІВАНОВА Івана Івановича, командира механізованої роти військової частини А1111 - "
    "КОМАНДИРОМ МЕХАНІЗОВАНОГО БАТАЛЬЙОНУ ВІЙСЬКОВОЇ ЧАСТИНИ А2222, ВОС - 0211003.",
    "1990 р.н., освіта: вигаданий університет у 2012 р., у ЗС - із 08.2008.",
    "1234567890.",
    "Призначається на вищу посаду у порядку просування по службі з шпк «капітан» на шпк «майор».",
]


def test_items_keep_person_and_full_text(tmp_path, monkeypatch):
    from nodeautomationtoolkit.order_index.items import iter_items

    items = iter_items(ITEM_ORDER_2024)
    assert len(items) == 1
    item = items[0]
    assert (item.rank, item.full_name, item.ipn) == ("Капітана", "ІВАНОВА Івана Івановича", "1234567890")
    assert item.current_position.startswith("командира механізованої роти")
    assert item.target_position.startswith("КОМАНДИРОМ МЕХАНІЗОВАНОГО БАТАЛЬЙОНУ")
    # Біографія і «Призначається …» лишаються в пункті — їх переносять у новий наказ.
    assert "1990 р.н." in item.text and "Призначається на вищу посаду" in item.text
    assert item.section == "§ 1"


def test_find_person_items_by_ipn_and_name(tmp_path, monkeypatch):
    monkeypatch.setattr(store, "WordTextReader", _NoWord)
    orders, index = tmp_path / "orders", tmp_path / "index"
    orders.mkdir()
    _write_order(orders / "Наказ №7 від 01.02.2024.docx", ITEM_ORDER_2024)
    _write_order(orders / "Наказ №9 від 05.06.2025.docx", [
        "§ 2",
        "Капітана ІВАНОВА Івана Івановича, командира механізованого батальйону військової частини А2222 - "
        "НАЧАЛЬНИКОМ ШТАБУ ВІЙСЬКОВОЇ ЧАСТИНИ А2222, ВОС - 0211003.",
        "1990 р.н., освіта: вигаданий університет у 2012 р., у ЗС - із 08.2008.",
        "1234567890.",
    ])
    summary = store.build_index(orders, index)
    assert summary.items == 2

    latest = store.find_person_items(index, ipn="1234567890")
    assert [item.order_number for item in latest] == ["9", "7"]
    assert latest[0].target_position.startswith("НАЧАЛЬНИКОМ ШТАБУ")
    # Посада, на яку призначили торік, — це посада, з якої призначають тепер.
    assert latest[1].target_position.casefold().startswith("командиром механізованого батальйону")
    assert latest[0].current_position.casefold().startswith("командира механізованого батальйону")

    by_name = store.find_person_items(index, full_name="Іванова Івана Івановича")
    assert [item.id for item in by_name] == [item.id for item in latest]
    assert store.find_person_items(index, full_name="ІВАНОВ") == by_name
    assert store.find_person_items(index, ipn="0000000000", full_name="Нікого") == []
    stats = store.load_stats(index)
    assert (stats.items, stats.people) == (2, 1)


def test_subdivision_chain_is_split():
    from nodeautomationtoolkit.order_index.subdivisions import chain_key, split_chain

    chain = split_chain(
        "командир ударного взводу безпілотних систем роти безпілотних систем командного пункту "
        "безпілотних систем окремого батальйону безпілотних систем полку безпілотних систем"
    )
    assert [link.text for link in chain] == [
        "ударного взводу безпілотних систем",
        "роти безпілотних систем",
        "командного пункту безпілотних систем",
        "окремого батальйону безпілотних систем",
        "полку безпілотних систем",
    ]
    assert [link.kind for link in chain] == ["взводу", "роти", "пункту", "батальйону", "полку"]
    # Номер частини лишається у своїй ланці, а в ключі його немає: «5 роти» = «роти».
    numbered = split_chain("КОМАНДИРОМ ТАНКОВОЇ РОТИ 5 ОКРЕМОЇ ТАНКОВОЇ БРИГАДИ")
    assert [link.text for link in numbered] == ["ТАНКОВОЇ РОТИ", "5 ОКРЕМОЇ ТАНКОВОЇ БРИГАДИ"]
    assert chain_key("5 окремої танкової бригади") == chain_key("ОКРЕМОЇ ТАНКОВОЇ БРИГАДИ")


def test_subdivisions_are_indexed(tmp_path, monkeypatch):
    monkeypatch.setattr(store, "WordTextReader", _NoWord)
    orders, index = tmp_path / "orders", tmp_path / "index"
    orders.mkdir()
    _write_order(orders / "Наказ №3 від 07.07.2025.docx", [
        "§ 1",
        "Капітана ІВАНОВА Івана Івановича, командира ударного взводу безпілотних систем роти безпілотних "
        "систем окремого батальйону безпілотних систем полку безпілотних систем - КОМАНДИРОМ РОТИ "
        "БЕЗПІЛОТНИХ СИСТЕМ ОКРЕМОГО БАТАЛЬЙОНУ БЕЗПІЛОТНИХ СИСТЕМ ПОЛКУ БЕЗПІЛОТНИХ СИСТЕМ, ВОС - 0210003.",
        "1990 р.н. 1234567890.",
    ])
    store.build_index(orders, index)
    rows = {row.display: row for row in store.load_subdivisions(index)}
    assert "роти безпілотних систем" in rows
    assert "ударного взводу безпілотних систем" in rows
    company = rows["роти безпілотних систем"]
    # Рота згадується і в посаді «з якої», і в посаді «на яку».
    assert (company.mentions, company.orders, company.kind) == (2, 1, "роти")
    assert rows["полку безпілотних систем"].mentions == 2
