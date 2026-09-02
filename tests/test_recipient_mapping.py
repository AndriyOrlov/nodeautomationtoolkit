from nodeautomationtoolkit.builtin_nodes.recipient_mapping import (
    groups_to_ciphers,
    read_recipient_mapping,
    map_military_units,
)
from nodeautomationtoolkit.core.table_types import DataTable


def test_reads_csv_mapping_and_builds_markers(tmp_path):
    source = tmp_path / "mapping.csv"
    source.write_text(
        "Відкрите найменування;Шифр;Куди направляється\n"
        "АЛЬФА;Ш-01;Відділ 1\nБРАВО;Ш-01;Відділ 1\n",
        encoding="utf-8",
    )

    result = read_recipient_mapping(str(source))

    assert result["markers"] == ["АЛЬФА", "БРАВО"]
    assert result["mapping"]["АЛЬФА"]["cipher"] == "Ш-01"
    assert isinstance(result["table"], DataTable)


def test_finds_headers_below_a_title_row(tmp_path):
    source = tmp_path / "mapping_with_title.csv"
    source.write_text(
        "ТЕСТОВА ТАБЛИЦЯ;;;\n"
        "Відкрите найменування;Шифр;Куди направляється;Примітка\n"
        "99 тестова бригада;ТЕСТ-А9001;Тестове управління;Вигадано\n",
        encoding="utf-8",
    )

    result = read_recipient_mapping(str(source))

    assert result["count"] == 1
    assert result["mapping"]["99 тестова бригада"]["cipher"] == "ТЕСТ-А9001"


def test_converts_and_merges_groups_with_same_cipher():
    result = groups_to_ciphers(
        groups={"АЛЬФА": "1. Перший", "БРАВО": "2. Другий", "НЕМАЄ": "3. Третій"},
        counts={"АЛЬФА": 1, "БРАВО": 1, "НЕМАЄ": 1},
        mapping={
            "АЛЬФА": {"cipher": "Ш-01", "destination": "Відділ 1"},
            "БРАВО": {"cipher": "Ш-01", "destination": "Відділ 1"},
        },
    )

    assert list(result["documents"]) == ["Ш-01"]
    assert result["documents"]["Ш-01"]["count"] == 2
    assert "1. Перший" in result["documents"]["Ш-01"]["content"]
    assert "2. Другий" in result["documents"]["Ш-01"]["content"]
    assert result["missing"] == ["НЕМАЄ"]
    assert len(result["report"].rows) == 3


def test_analyze_senders_and_split_by_senders():
    from nodeautomationtoolkit.builtin_nodes.recipient_mapping import (
        analyze_senders,
        split_by_senders,
    )

    text = (
        "НАКАЗ командира військової частини А0000\n"
        "15 січня 2026 року № 100, м. Київ\n\n"
        "§ 1\n"
        "1. Пункт про 160 окрему механізовану бригаду призначити.\n"
        "2. Пункт про 167 окрему механізовану бригаду звільнити.\n"
    )

    mapping = {
        "160 окрема механізована бригада": {"open_name": "160 окрема механізована бригада", "cipher": "в/ч А1600"},
        "167 окрема механізована бригада": {"open_name": "167 окрема механізована бригада", "cipher": "в/ч А1670"},
    }

    res = analyze_senders(text=text, mapping=mapping)
    assert isinstance(res["sender_paragraphs"], dict)
    assert len(res["senders_list"]) > 0
    assert isinstance(res["table"], DataTable)

    res_split = split_by_senders(text=text, mapping=mapping)
    assert isinstance(res_split["blocks"], dict)
    assert res_split["senders_count"] > 0


def test_army_corps_prioritization_over_subordinate_units():
    from nodeautomationtoolkit.builtin_nodes.recipient_mapping import analyze_senders

    text = (
        "§ 1\n"
        "1. Молодшого лейтенанта призначити у 167 окрему механізовану бригаду 10-го армійського корпусу.\n"
        "2. Сержанта звільнити з 154 окремої механізованої бригади 3 АК.\n"
    )

    mapping = {
        "10 армійський корпус": {
            "open_name": "10 армійський корпус",
            "cipher": "А0010",
            "abbreviation": "10АК",
        },
        "3 армійський корпус": {
            "open_name": "3 армійський корпус",
            "cipher": "А0003",
            "abbreviation": "3АК",
        },
    }
    res = analyze_senders(text=text, mapping=mapping)
    assert any(sender.startswith("10АК") for sender in res["senders_list"])
    assert any(sender.startswith("3АК") for sender in res["senders_list"])
    key10 = next(sender for sender in res["sender_paragraphs"] if sender.startswith("10АК"))
    assert len(res["sender_paragraphs"][key10]) == 1


def test_tck_full_wording_extracts_only_oblast():
    from nodeautomationtoolkit.builtin_nodes.recipient_mapping import _extract_tck_region_hints, analyze_senders

    text = (
        "§ 1\n"
        "1. Направити документи до Ковельського районного територіального центру "
        "комплектування та соціальної підтримки Волинської області.\n"
    )

    extracted = _extract_tck_region_hints(text)
    assert "Волинська" in extracted or "Волинський" in extracted

    mapping = {
        "Волинський обласний ТЦК та СП": {
            "open_name": "Волинський обласний ТЦК та СП",
            "cipher": "Волинський ОТЦК та СП",
        },
    }
    analysis = analyze_senders(text=text, mapping=mapping)
    assert analysis["senders_list"]


def test_generate_decision_order():
    from nodeautomationtoolkit.builtin_nodes.recipient_mapping import generate_decision_order

    text = (
        "НАКАЗ командира 160 окремої механізованої бригади\n"
        "15 січня 2026 року № 100, м. Київ\n\n"
        "§ 1\n"
        "1. Молодшого лейтенанта призначити у 167 окрему механізовану бригаду.\n"
        "2. Військовослужбовця цієї самої бригади звільнити з посади.\n"
        "3. Командирові цього самого полку провести розслідування.\n"
    )

    mapping = {
        "167 окрема механізована бригада": {"open_name": "167 окрема механізована бригада", "cipher": "в/ч А1670"},
    }

    res = generate_decision_order(text=text, mapping=mapping, new_header="ЗАКРИТИЙ НАКАЗ командира в/ч А0000")
    decision_text = res["decision_text"]

    assert "ЗАКРИТИЙ НАКАЗ командира в/ч А0000" in decision_text
    assert "НАКАЗ командира 160 окремої механізованої бригади" not in decision_text
    assert "А1670" in decision_text
    assert "цієї самої військової частини" in decision_text
    assert res["replaced_count"] >= 3


def test_order_block_constructor_pipeline():
    from nodeautomationtoolkit.builtin_nodes.recipient_mapping import (
        parse_to_blocks,
        filter_transform_blocks,
        assemble_from_blocks,
    )

    text = (
        "НАКАЗ командира 160 окремої механізованої бригади\n"
        "15 січня 2026 року № 100, м. Київ\n\n"
        "§ 1\n"
        "1. Молодшого лейтенанта призначити у 167 окрему механізовану бригаду.\n"
        "2. Військовослужбовця цієї самої бригади звільнити з посади.\n"
        "Підстава: рапорт командира.\n"
    )

    mapping = {
        "167 окрема механізована бригада": {"open_name": "167 окрема механізована бригада", "cipher": "в/ч А1670"},
    }

    # 1. Parse into blocks
    parsed = parse_to_blocks(text=text)
    blocks = parsed["blocks"]
    assert len(blocks) >= 4  # header, section, item 1, item 2, basis

    # 2. Filter & transform blocks
    transformed = filter_transform_blocks(blocks=blocks, mapping=mapping, replace_unit_phrases=True)
    tf_blocks = transformed["blocks"]
    assert transformed["modified_count"] >= 2

    # 3. Assemble back
    assembled = assemble_from_blocks(blocks=tf_blocks, new_header="ЗАКРИТИЙ НАКАЗ командира в/ч А0000")
    final_text = assembled["text"]

    assert "ЗАКРИТИЙ НАКАЗ командира в/ч А0000" in final_text
    assert "А1670" in final_text
    assert "цієї самої військової частини" in final_text
    assert "Підстава: рапорт командира" in final_text


def test_tck_is_not_inferred_from_column_c_when_column_a_does_not_match():
    mapping = {
        "99 окремий розвідувальний батальйон": {
            "cipher": "в/ч А1234",
            "open_name": "99 окремий розвідувальний батальйон",
            "abbreviation": "99ОРБ",
        },
        "Камінь-Каширський районний територіальний центр комплектування та соціальної підтримки": {
            "cipher": "Камінь-Каширський РТЦК",
            "open_name": "Камінь-Каширський районний територіальний центр комплектування та соціальної підтримки",
            "abbreviation": "Волинський ОТЦК та СП",
        },
        "Одеський обласний ТЦК та СП": {
            "cipher": "Одеський ОТЦК та СП",
            "open_name": "Одеський обласний ТЦК та СП",
            "abbreviation": "Одеський ОТЦК та СП",
        },
    }
    text = (
        "Військовослужбовців 99 окремого розвідувального батальйону направити до "
        "Волинського обласного територіального центру комплектування та соціальної підтримки:\n\n"
        "1. Майор ІВАНОВ І.І. - звільнити з військової служби у запас. "
        "Підлягає направленню на військовий облік до Волинського обласного ТЦК та СП."
    )
    res = map_military_units(text=text, mapping=mapping)
    paragraphs = res.get("unit_paragraphs", {})
    assert "99ОРБ А1234" in paragraphs
    assert "Волинський ОТЦК та СП" not in paragraphs
    assert "Одеський ОТЦК та СП" not in paragraphs


def test_all_9_items_preserved_for_both_senders():
    mapping = {
        "99 окремий розвідувальний батальйон": {
            "cipher": "в/ч А1234",
            "open_name": "99 окремий розвідувальний батальйон",
            "abbreviation": "99ОРБ",
        },
        "Волинський обласний ТЦК та СП": {
            "cipher": "Волинський ОТЦК та СП",
            "open_name": "Волинський обласний територіальний центр комплектування та соціальної підтримки",
            "abbreviation": "Волинський ОТЦК та СП",
        },
    }
    items_str = "\n".join([f"{i}. Майор ТЕСТОВ{i} Т.Т. - призначити на посаду." for i in range(1, 10)])
    text = (
        "Військовослужбовців 99 окремого розвідувального батальйону направити до "
        "Волинського обласного територіального центру комплектування та соціальної підтримки:\n\n"
        + items_str
    )
    res = map_military_units(text=text, mapping=mapping)
    paragraphs = res.get("unit_paragraphs", {})
    assert len(paragraphs["99ОРБ А1234"]["items"]) == 9
    assert len(paragraphs["Волинський ОТЦК та СП"]["items"]) == 9


def test_reads_headerless_table_with_ABCD_columns(tmp_path):
    source = tmp_path / "headerless_mapping.csv"
    source.write_text(
        "99 окремий розвідувальний батальйон;в/ч А1234;99ОРБ;10 армійський корпус;м. Київ\n"
        "256 окрема рота;в/ч А2560;256ОР;;м. Львів\n",
        encoding="utf-8",
    )

    result = read_recipient_mapping(str(source))
    mapping = result["mapping"]

    assert "99 окремий розвідувальний батальйон" in mapping
    assert mapping["99 окремий розвідувальний батальйон"]["cipher"] == "в/ч А1234"
    assert mapping["99 окремий розвідувальний батальйон"]["abbreviation"] == "99ОРБ"
    assert mapping["99 окремий розвідувальний батальйон"]["corps"] == "10 армійський корпус"

    assert "256 окрема рота" in mapping
    assert mapping["256 окрема рота"]["cipher"] == "в/ч А2560"
    assert mapping["256 окрема рота"]["abbreviation"] == "256ОР"


def test_tck_is_additive_to_unit_match():
    mapping = {
        "160 окрема механізована бригада": {
            "cipher": "в/ч А1600",
            "open_name": "160 окрема механізована бригада",
            "abbreviation": "160 ОМБр",
        },
        "Волинський обласний ТЦК та СП": {
            "cipher": "Волинський ОТЦК та СП",
            "open_name": "Волинський обласний територіальний центр комплектування та соціальної підтримки",
            "abbreviation": "Волинський ОТЦК та СП",
        },
    }
    text = (
        "НАКАЗ командира 160 окремої механізованої бригади\n\n"
        "1. Капітана ПЕТРЕНКА П.П., командира роти 160 окремої механізованої бригади, "
        "звільнити у запас. Підлягає направленню на облік до Волинського обласного ТЦК та СП."
    )
    res = map_military_units(text=text, mapping=mapping)
    paragraphs = res.get("unit_paragraphs", {})
    assert "160 ОМБр А1600" in paragraphs
    assert "Волинський ОТЦК та СП" in paragraphs


def test_hash_priority_rule_overrides_standard_match():
    mapping = {
        "3 штурмовий батальйон": {
            "cipher": "в/ч А3333",
            "open_name": "3 штурмовий батальйон",
            "abbreviation": "3ШБ",
        },
        "#3 штурмовий батальйон": {
            "cipher": "в/ч А4444",
            "open_name": "#3 штурмовий батальйон",
            "abbreviation": "3ШБ",
        },
    }
    text = (
        "НАКАЗ командира 10 армійського корпусу\n\n"
        "1. Майора ІВАНОВА І.І. призначити у 3 штурмовий батальйон."
    )
    res = map_military_units(text=text, mapping=mapping)
    paragraphs = res.get("unit_paragraphs", {})
    # Has priority row #3 штурмовий батальйон -> в/ч А4444
    assert "3ШБ А4444" in paragraphs
    assert "3ШБ А3333" not in paragraphs


def test_ivano_frankivsk_tck_normalization():
    text = (
        "НАКАЗ командира 10 армійського корпусу\n\n"
        "1. Направити до Івано-Франківського обласного ТЦК та СП.\n"
        "2. Підлягає направленню до Франківського ОТЦК та СП."
    )
    mapping = {
        "Івано-Франківський обласний ТЦК та СП": {
            "open_name": "Івано-Франківський обласний ТЦК та СП",
            "cipher": "Івано-Франківський ОТЦК та СП",
        },
    }
    res = map_military_units(text=text, mapping=mapping)
    paragraphs = res.get("unit_paragraphs", {})
    assert "Івано-Франківський ОТЦК та СП" in paragraphs
    assert "Іваноський-Франківський ОТЦК та СП" not in paragraphs
    assert "Франківський ОТЦК та СП" not in paragraphs
    # Both items 1 and 2 mapped to the same unified Івано-Франківський ОТЦК та СП
    assert len(paragraphs["Івано-Франківський ОТЦК та СП"]["items"]) == 2


def test_deduplicates_by_column_b_cipher():
    mapping = {
        "160 окрема механізована бригада": {
            "cipher": "в/ч А1600",
            "open_name": "160 окрема механізована бригада",
            "abbreviation": "160 ОМБр",
        },
        "160 мехбр": {
            "cipher": "А1600",
            "open_name": "160 мехбр",
            "abbreviation": "160 мехбр",
        },
    }
    text = (
        "НАКАЗ командира 10 армійського корпусу\n\n"
        "1. Капітана Петренка направити до 160 окремої механізованої бригади.\n"
        "2. Майора Іванова направити до 160 мехбр."
    )
    res = map_military_units(text=text, mapping=mapping)
    paragraphs = res.get("unit_paragraphs", {})
    # Both rows mapping to cipher А1600 are merged under a single unified recipient key
    assert len(paragraphs) == 1
    assert "160 ОМБр А1600" in paragraphs
    assert len(paragraphs["160 ОМБр А1600"]["items"]) == 2


def test_corps_cipher_deduplication():
    mapping = {
        "25 армійський корпус": {
            "cipher": "в/ч А2525",
            "open_name": "25 армійський корпус",
            "abbreviation": "25 АК",
        },
        "100 ОМБр": {
            "cipher": "в/ч А1000",
            "open_name": "100 ОМБр",
            "corps": "25АК",
        },
    }
    text = (
        "НАКАЗ командира 25 армійського корпусу\n\n"
        "1. Перевести до складу 25 армійського корпусу.\n"
        "2. Передати майно 25АК."
    )
    res = map_military_units(text=text, mapping=mapping)
    paragraphs = res.get("unit_paragraphs", {})
    # Should consolidate 25АК and 25 армійський корпус into a SINGLE entry for А2525
    assert len(paragraphs) == 1
    key = list(paragraphs.keys())[0]
    assert "А2525" in key
    assert len(paragraphs[key]["items"]) == 2


def test_read_recipient_mapping_6_columns(tmp_path):
    source = tmp_path / "mapping_6_cols.csv"
    source.write_text(
        "Відкрите найменування;Шифр;Скорочення;Корпус;Кому;Куди\n"
        "Львівський обласний ТЦК та СП;Львівський ОТЦК;Львівський ОТЦК;;Начальнику Львівського обласного територіального центру комплектування та соціальної підтримки;м. Львів\n"
        "72 окрема механізована бригада;А2167;72 омбр;10АК;Командиру військової частини А2167;м. Біла Церква\n",
        encoding="utf-8",
    )

    result = read_recipient_mapping(str(source))
    assert result["count"] == 2
    lviv_entry = result["mapping"]["Львівський обласний ТЦК та СП"]
    assert lviv_entry["recipient_to"] == "Начальнику Львівського обласного територіального центру комплектування та соціальної підтримки"
    assert lviv_entry["destination_where"] == "м. Львів"

    brigade_entry = result["mapping"]["72 окрема механізована бригада"]
    assert brigade_entry["recipient_to"] == "Командиру військової частини А2167"
    assert brigade_entry["destination_where"] == "м. Біла Церква"


def test_tck_recipient_and_destination_in_extracts():
    mapping = {
        "Львівський обласний територіальний центр комплектування та соціальної підтримки": {
            "cipher": "Львівський ОТЦК та СП",
            "open_name": "Львівський обласний територіальний центр комплектування та соціальної підтримки",
            "abbreviation": "Львівський ОТЦК та СП",
            "recipient_to": "Начальнику Львівського обласного територіального центру комплектування та соціальної підтримки",
            "destination_where": "м. Львів",
        },
    }
    text = (
        "НАКАЗ командира військової частини А0000\n\n"
        "11. Солдата призначити, призваного Сихівським РТЦК та СП м. Львова."
    )
    res = map_military_units(text=text, mapping=mapping)
    paragraphs = res.get("unit_paragraphs", {})
    assert "Львівський ОТЦК та СП" in paragraphs
    tck_data = paragraphs["Львівський ОТЦК та СП"]
    assert tck_data["recipient_to"] == "Начальнику Львівського обласного територіального центру комплектування та соціальної підтримки"
    assert tck_data["destination_where"] == "м. Львів"


def test_missing_destination_marked_as_kudy():
    mapping = {
        "72 окрема механізована бригада": {
            "cipher": "А2167",
            "open_name": "72 окрема механізована бригада",
            "recipient_to": "Командиру військової частини А2167",
            "destination_where": "",  # Порожньо в Excel
        },
    }
    text = (
        "НАКАЗ командира військової частини А0000\n\n"
        "1. Перевести до 72 окремої механізованої бригади."
    )
    res = map_military_units(text=text, mapping=mapping)
    paragraphs = res.get("unit_paragraphs", {})
    key = list(paragraphs.keys())[0]
    unit_data = paragraphs[key]
    assert unit_data["destination_where"] == "КУДИ"


def test_xlsx_formulas_and_empty_recipient_preserves_destination(tmp_path):
    import openpyxl

    xlsx_file = tmp_path / "test_dict.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Словник"
    ws.append(["Відкрита назва", "Шифр", "Скорочення", "Корпус", "Кому", "Куди"])
    # Row 1: Formula in column B and E
    ws.append(["93 окрема механізована бригада", "А1302", "93 омбр", "10 АК", "=CONCAT(\"Командиру \", \"А1302\")", "м. Черкаське"])
    # Row 2: Empty column E, but Column F is filled
    ws.append(["14 окрема механізована бригада", "А1008", "14 омбр", "", "", "м. Володимир"])
    wb.save(str(xlsx_file))

    res = read_recipient_mapping(str(xlsx_file))
    m = res["mapping"]

    entry_93 = m["93 окрема механізована бригада"]
    assert entry_93["cipher"] == "А1302"
    assert entry_93["destination_where"] == "м. Черкаське"

    entry_14 = m["14 окрема механізована бригада"]
    assert entry_14["cipher"] == "А1008"
    assert entry_14["recipient_to"] == ""
    assert entry_14["destination_where"] == "м. Володимир"


def test_military_registration_in_bio_lines_ignored_for_unit_routing():
    mapping = {
        "72 окрема механізована бригада": {
            "cipher": "А2167",
            "open_name": "72 окрема механізована бригада",
            "abbreviation": "72 омбр",
        },
        "Рівненський РТЦК та СП": {
            "cipher": "Рівненський РТЦК",
            "open_name": "Рівненський РТЦК та СП",
            "abbreviation": "Рівненський РТЦК",
        },
    }
    text = (
        "НАКАЗ КОМАНДИРА ВІЙСЬКОВОЇ ЧАСТИНИ А0000\n\n"
        "1. Солдата ІВАНОВА Івана Івановича призначити на посаду до 72 окремої механізованої бригади.\n"
        "1995 р.н., освіта вища, ІПН 1234567890, ВОС-100915, підлягає направленню на військовий облік до Рівненського РТЦК та СП Рівненської області.\n"
        "Підстава: рапорт командира.\n"
    )
    res = map_military_units(text=text, mapping=mapping)
    paragraphs = res.get("unit_paragraphs", {})
    # 72 омбр must be present
    assert any("72" in k or "А2167" in k for k in paragraphs.keys())
    # Рівненський РТЦК must NOT be present as a recipient
    assert not any("Рівнен" in k for k in paragraphs.keys())


def test_female_soldier_multiline_bio_ignored_for_unit_routing():
    mapping = {
        "14 окрема механізована бригада": {
            "cipher": "А1008",
            "open_name": "14 окрема механізована бригада",
            "abbreviation": "14 омбр",
        },
        "Рівненський ОМТЦК та СП": {
            "cipher": "Рівненський ОМТЦК",
            "open_name": "Рівненський ОМТЦК та СП",
            "abbreviation": "Рівненський ОМТЦК",
        },
    }
    text = (
        "НАКАЗ КОМАНДИРА ВІЙСЬКОВОЇ ЧАСТИНИ А0000\n\n"
        "1. Сержанта ПЕТРЕНКО Олену Василівну призначити до 14 окремої механізованої бригади.\n"
        "Народилась 15 березня 1981 року.\n"
        "У ЗС – із 10.2010\n"
        "Підлягає направленню на військовий облік до Рівненського ОМТЦК та СП Рівненської області.\n"
        "Підстава: наказ Головнокомандувача ЗС України.\n"
    )
    res = map_military_units(text=text, mapping=mapping)
    paragraphs = res.get("unit_paragraphs", {})
    # 14 омбр must be present
    assert any("14" in k or "А1008" in k for k in paragraphs.keys())
    # Рівненський ОМТЦК must NOT be present
    assert not any("Рівнен" in k for k in paragraphs.keys())




