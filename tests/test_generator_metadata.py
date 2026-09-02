from generate_extracts import (
    back_page_tag_values,
    build_message_recipient_list,
    build_copy_two_filename,
    extract_metadata_from_text,
    find_unmatched_open_unit_spans,
    format_ukr_date,
)


def test_extracts_order_metadata_from_order_header_text():
    text = "НАКАЗ\n‘03’ серпня 2026 року м. Київ № 390\n"

    assert extract_metadata_from_text(text) == ("390", "03.08.2026")


def test_missing_or_invalid_date_is_not_replaced_with_current_date():
    assert format_ukr_date("") == ""


def test_slash_becomes_new_line_for_certifier_and_executor():
    """«/» у посаді розбиває її на абзаци — правило спільне для витягів і примірників."""
    from generate_extracts import _slash_to_lines

    assert _slash_to_lines("Начальник штабу / полковник") == "Начальник штабу\rполковник"
    assert _slash_to_lines("") == ""


def test_slash_rule_keeps_unit_codes_and_dates_intact():
    """«в/ч» та числові дроби на кшталт «1/2» розбивати не можна."""
    from generate_extracts import _slash_to_lines

    assert _slash_to_lines("Командир в/ч А0000") == "Командир в/ч А0000"
    assert _slash_to_lines("наказ 1/2") == "наказ 1/2"


def test_order_number_tag_always_has_number_sign_without_space():
    """`{{номер_наказу}}` має бути «№555» в усіх режимах — без пробілу."""
    values = back_page_tag_values("555", "06.08.2026")
    assert values["{{номер_наказу}}"] == "№555"
    assert "№ " not in values["{{номер_наказу}}"]


def test_back_page_tags_use_only_filename_metadata_and_copy_two_label():
    assert back_page_tag_values("396", "06.08.2026") == {
        "{{згідно_з_оригіналом}}": "Згідно з оригіналом",
        "{{примірник}}": "Примірник № 2",
        "{{номер_наказу}}": "№396",
        "{{дата_наказу}}": "“06” серпня 2026 року",
    }
    assert back_page_tag_values("", "") == {
        "{{згідно_з_оригіналом}}": "Згідно з оригіналом",
        "{{примірник}}": "Примірник № 2",
    }


def test_copy_two_filename_does_not_invent_missing_order_metadata():
    assert build_copy_two_filename("396", "06.08.2026", "source.docx") == "2,3_№396 від 06.08.2026.docx"
    assert build_copy_two_filename("", "", "source.docx") == "2,3_source.docx"


def test_unmatched_open_units_are_identified_for_yellow_highlighting():
    text = "Направити до 12 окремої механізованої бригади та військової частини А0000."
    spans = find_unmatched_open_unit_spans(text)

    assert spans
    assert text[spans[0][0]:spans[0][1]] == "12 окремої механізованої бригади"


def test_message_content_uses_existing_closed_order_transformations():
    from nodeautomationtoolkit.builtin_nodes.recipient_mapping import generate_decision_order

    mapping = {
        "12 окрема механізована бригада": {
            "open_name": "12 окрема механізована бригада",
            "cipher": "А1234",
            "abbreviation": "12 ОМБр",
        }
    }
    result = generate_decision_order(
        text="§ 1\nНаправити до 12 окремої механізованої бригади. До цієї самої бригади.",
        mapping=mapping,
        new_header="",
    )

    assert "військової частини А1234" in result["decision_text"]
    assert "цієї самої військової частини" in result["decision_text"]


def test_message_recipients_list_corps_before_subordinate_unit_as_separate_rows():
    from nodeautomationtoolkit.builtin_nodes.recipient_mapping import map_military_units

    mapping = {
        "10 армійський корпус": {
            "open_name": "10 армійський корпус",
            "cipher": "А0010",
            "abbreviation": "10 АК",
            "recipient_to": "Командиру 10 АК",
        },
        "12 окрема механізована бригада": {
            "open_name": "12 окрема механізована бригада",
            "cipher": "А1234",
            "abbreviation": "12 ОМБр",
            "corps": "10 армійський корпус",
            "recipient_to": "Командиру 12 ОМБр",
        },
    }
    routes = map_military_units(
        text="§ 1\n1. Направити до 12 окремої механізованої бригади.",
        mapping=mapping,
    )

    assert build_message_recipient_list(mapping, routes) == [
        "Командиру 10 АК військової частини А0010",
        "Командиру 12 ОМБр військової частини А1234",
    ]


def test_message_recipient_does_not_repeat_military_unit_phrase():
    from nodeautomationtoolkit.builtin_nodes.recipient_mapping import map_military_units

    mapping = {
        "окрема частина": {
            "open_name": "окрема частина",
            "cipher": "А0000",
            "recipient_to": "Командиру військової частини",
        }
    }
    routes = map_military_units(text="1. До окрема частина.", mapping=mapping)

    assert build_message_recipient_list(mapping, routes) == ["Командиру військової частини А0000"]


def test_all_matched_recipients_of_one_item_receive_extracts():
    from nodeautomationtoolkit.builtin_nodes.recipient_mapping import map_military_units

    mapping = {
        "перша частина": {"open_name": "перша частина", "cipher": "А0001", "abbreviation": "1Ч"},
        "друга частина": {"open_name": "друга частина", "cipher": "А0002", "abbreviation": "2Ч"},
        "третя частина": {"open_name": "третя частина", "cipher": "А0003", "abbreviation": "3Ч"},
    }
    result = map_military_units(
        "§ 1\n1. До перша частина, друга частина, третя частина.", mapping=mapping
    )

    assert len(result["unit_paragraphs"]) == 3
    assert all(len(data["items"]) == 1 for data in result["unit_paragraphs"].values())


def test_routing_audit_lists_every_numbered_item_without_truncation():
    from nodeautomationtoolkit.builtin_nodes.recipient_mapping import map_military_units

    mapping = {
        "перша частина": {"open_name": "перша частина", "cipher": "А0001"},
    }
    text = "§ 1\n" + "\n".join(
        f"{number}. Направити до перша частина." for number in range(1, 12)
    )

    result = map_military_units(text=text, mapping=mapping)

    assert len(result["routing_audit"]) == 11
    assert result["unmatched_items"] == []
    assert len(next(iter(result["unit_paragraphs"].values()))["items"]) == 11


def test_unmatched_item_is_returned_for_excel_control():
    from nodeautomationtoolkit.builtin_nodes.recipient_mapping import map_military_units

    result = map_military_units("§ 1\n1. Пункт без адресата.", mapping={})

    assert len(result["unmatched_items"]) == 1
    assert result["routing_audit"][0]["final_recipients"] == "—"


def test_internal_reference_keeps_explicitly_named_tck_as_recipient():
    from nodeautomationtoolkit.builtin_nodes.recipient_mapping import map_military_units

    mapping = {
        "Закарпатський обласний ТЦК та СП": {
            "open_name": "Закарпатський обласний ТЦК та СП",
            "cipher": "Закарпатський ОТЦК",
        },
    }
    text = (
        "§ 1\n"
        "1. Військовослужбовця Закарпатського обласного територіального центру "
        "комплектування та соціальної підтримки призначити до цього самого центру."
    )

    result = map_military_units(text, mapping=mapping)

    assert result["unmatched_items"] == []
    assert len(result["unit_paragraphs"]) == 1
    assert "підтверджено названим адресатом" in result["routing_audit"][0]["applied_rules"]


def test_unrouted_management_change_is_excluded_not_reported_as_missing():
    from nodeautomationtoolkit.builtin_nodes.recipient_mapping import map_military_units

    result = map_military_units(
        "§ 1\n1. Призначити на посаду до відділу управління.", mapping={}
    )

    assert result["unmatched_items"] == []
    assert len(result["skipped_items"]) == 1
    assert "виключено із загального переліку" in result["routing_audit"][0]["applied_rules"]


def test_recruiting_center_matches_when_number_order_differs_from_excel():
    from nodeautomationtoolkit.builtin_nodes.recipient_mapping import map_military_units

    mapping = {
        "Центр рекрутингу № 7": {
            "open_name": "Центр рекрутингу № 7",
            "cipher": "А0007",
        },
    }
    result = map_military_units(
        "§ 1\n8. Призначити до 7 центру рекрутингу.", mapping=mapping
    )

    assert result["unmatched_items"] == []
    assert len(result["unit_paragraphs"]) == 1


def test_tck_in_assignment_paragraph_is_not_dropped_as_biographical_text():
    from nodeautomationtoolkit.builtin_nodes.recipient_mapping import map_military_units

    mapping = {
        "Рівненський обласний ТЦК та СП": {
            "open_name": "Рівненський обласний ТЦК та СП",
            "cipher": "Рівненський ОТЦК",
        },
    }
    text = (
        "§ 1\n"
        "7. Старшого офіцера призначити на посаду.\n"
        "З Рівненського РТЦК та СП — до Рівненського обласного ТЦК та СП."
    )

    result = map_military_units(text, mapping=mapping)

    assert result["unmatched_items"] == []
    assert len(result["unit_paragraphs"]) == 1


def test_recruiting_center_is_not_searched_by_column_c_abbreviation():
    from nodeautomationtoolkit.builtin_nodes.recipient_mapping import map_military_units

    mapping = {
        "Інший запис стовпця A": {
            "open_name": "Інший запис стовпця A",
            "abbreviation": "7 ЦР",
            "cipher": "А0007",
        },
    }
    result = map_military_units(
        "§ 1\n2. Перевести з 7 центру рекрутингу.", mapping=mapping
    )

    assert len(result["unmatched_items"]) == 1
    assert result["unit_paragraphs"] == {}


def test_tck_is_not_searched_by_column_e_recipient():
    from nodeautomationtoolkit.builtin_nodes.recipient_mapping import map_military_units

    recipient = "Начальнику Рівненського обласного ТЦК та СП"
    mapping = {
        "Інший запис стовпця A": {
            "open_name": "Інший запис стовпця A",
            "cipher": "Рівненський ОТЦК",
            "abbreviation": "",
            "recipient_to": recipient,
        },
    }
    result = map_military_units(
        "§ 1\n7. Перевести з Рівненського районного територіального центру "
        "комплектування та соціальної підтримки Рівненської області.",
        mapping=mapping,
    )

    assert len(result["unmatched_items"]) == 1
    assert result["unit_paragraphs"] == {}


def test_recruiting_center_is_not_searched_by_column_e_recipient():
    from nodeautomationtoolkit.builtin_nodes.recipient_mapping import map_military_units

    recipient = "Начальнику 7 центру рекрутингу"
    mapping = {
        "Інший запис стовпця A": {
            "open_name": "Інший запис стовпця A",
            "cipher": "А0007",
            "abbreviation": "",
            "recipient_to": recipient,
        },
    }
    result = map_military_units(
        "§ 1\n8. Перевести з 7 центру рекрутингу.", mapping=mapping
    )

    assert len(result["unmatched_items"]) == 1
    assert result["unit_paragraphs"] == {}


def test_routing_keeps_column_a_names_before_biographical_suffix_in_same_paragraph():
    from nodeautomationtoolkit.builtin_nodes.recipient_mapping import map_military_units

    mapping = {
        "7 центр рекрутингу": {
            "open_name": "7 центр рекрутингу",
            "cipher": "А0007",
            "abbreviation": "7 ЦР",
        },
        "Центр підготовки підрозділів": {
            "open_name": "Центр підготовки підрозділів",
            "cipher": "А0008",
            "abbreviation": "ЦПП",
        },
    }
    text = (
        "§ 1\n"
        "8. Лейтенанта, начальника групи 7 центру рекрутингу – "
        "НАЧАЛЬНИКОМ ГРУПИ ЦЕНТРУ ПІДГОТОВКИ ПІДРОЗДІЛІВ, "
        "народився 10 січня 1990 року, освіта вища."
    )

    result = map_military_units(text, mapping=mapping)

    assert result["unmatched_items"] == []
    assert set(result["unit_paragraphs"]) == {"7 ЦР А0007", "ЦПП А0008"}


def test_oblast_tck_routes_when_biography_follows_in_same_paragraph():
    from nodeautomationtoolkit.builtin_nodes.recipient_mapping import map_military_units

    mapping = {
        "Івано-Франківський обласний територіальний центр комплектування та соціальної підтримки": {
            "open_name": "Івано-Франківський обласний територіальний центр комплектування та соціальної підтримки",
            "cipher": "Івано-Франківський ОТЦК та СП",
        },
    }
    text = (
        "§ 1\n"
        "4. Капітана, старшого офіцера Івано-Франківського районного "
        "територіального центру комплектування та соціальної підтримки "
        "Івано-Франківської області – НАЧАЛЬНИКОМ ВІДДІЛЕННЯ, "
        "р.н. 1990, підлягає направленню на військовий облік."
    )

    result = map_military_units(text, mapping=mapping)

    assert result["unmatched_items"] == []
    assert list(result["unit_paragraphs"]) == ["Івано-Франківський ОТЦК та СП"]


def test_ukrainian_typography_non_breaking_spaces():
    from generate_extracts import apply_ukrainian_typography

    raw = "з наказом та звіт до штабу в частину із списком на посаду і призначити по 5 за серпень від 15.08.2026 р.н. у ЗС із 2022 шпк майор ВОС-0210003"
    formatted = apply_ukrainian_typography(raw)

    assert "з\u00A0наказом" in formatted
    assert "та\u00A0звіт" in formatted
    assert "до\u00A0штабу" in formatted
    assert "в\u00A0частину" in formatted
    assert "із\u00A0списком" in formatted
    assert "на\u00A0посаду" in formatted
    assert "і\u00A0призначити" in formatted
    assert "по\u00A05" in formatted
    assert "за\u00A0серпень" in formatted
    assert "від\u00A015.08.2026" in formatted
    assert "р.н.\u00A0у" in formatted
    assert "із\u00A02022" in formatted
    assert "шпк\u00A0майор" in formatted


def test_clean_duplicated_units():
    from generate_extracts import clean_duplicated_units

    raw = "військової частини військової частини А1234 військової частини А1234 — ВІЙСЬКОВОЇ ЧАСТИНИ ВІЙСЬКОВОЇ ЧАСТИНИ А5678"
    cleaned = clean_duplicated_units(raw)

    assert cleaned == "військової частини А1234 — ВІЙСЬКОВОЇ ЧАСТИНИ А5678"


def test_is_biographical_paragraph():
    from generate_extracts import is_biographical_paragraph

    assert is_biographical_paragraph("1995 р.н., освіта: НТУ у 2017 р., у ЗС із 01.01.2022.") is True
    assert is_biographical_paragraph("12345678.") is True
    assert is_biographical_paragraph("РНОКПП 1234567890") is True
    assert is_biographical_paragraph("1. Старшого лейтенанта Іванова І.І.") is False
    assert is_biographical_paragraph("Відповідно до пунктів 82, 83... ЗВІЛЬНИТИ і ПРИЗНАЧИТИ:") is False
    assert is_biographical_paragraph("Призначається на рівнозначну посаду для більш доцільного використання.") is False


def test_unmatched_open_units_ignores_internal_battalions():
    from generate_extracts import find_unmatched_open_unit_spans

    # Лінійні батальйони, прикріплені до закритих шифрів ВЧ, не повинні підсвічуватися
    text_with_encrypted_units = (
        "командира гранатометного взводу штурмового батальйону військової частини А1111 "
        "— КОМАНДИРОМ ГРАНАТОМЕТНОГО ВЗВОДУ МЕХАНІЗОВАНОГО БАТАЛЬЙОНУ ВІЙСЬКОВОЇ ЧАСТИНИ А2222."
    )
    spans = find_unmatched_open_unit_spans(text_with_encrypted_units)
    assert len(spans) == 0

    # Незашифрована відкрита окрема бригада має підсвічуватися
    text_with_unmatched = "командира 100 окремої механізованої бригади призначити на посаду."
    spans_unmatched = find_unmatched_open_unit_spans(text_with_unmatched)
    assert len(spans_unmatched) > 0


def test_ensure_blank_line_before_items():
    from generate_extracts import ensure_blank_line_before_items

    # 1. Пункти без ентеру -> додається рівно 1 ентер
    raw = (
        "11. Старшого лейтенанта Іванова...\n"
        "1995 р.н., освіта...\n"
        "12. Капітана Петрова...\n"
        "1990 р.н., освіта...\n"
        "13. Майора Сидорова..."
    )
    result = ensure_blank_line_before_items(raw)
    assert (
        result
        == "11. Старшого лейтенанта Іванова...\n"
        "1995 р.н., освіта...\n\n"
        "12. Капітана Петрова...\n"
        "1990 р.н., освіта...\n\n"
        "13. Майора Сидорова..."
    )

    # 2. Якщо ентер вже є -> додатковий НЕ вставляється (лишається рівно 1)
    already_spaced = (
        "11. Старшого лейтенанта Іванова...\n"
        "1995 р.н., освіта...\n\n"
        "12. Капітана Петрова...\n"
        "1990 р.н., освіта...\n\n"
        "13. Майора Сидорова..."
    )
    assert ensure_blank_line_before_items(already_spaced) == already_spaced

    # 3. Якщо було кілька зайвих порожніх рядків -> згортається до рівно 1
    multiple_spaces = (
        "11. Старшого лейтенанта Іванова...\n"
        "1995 р.н., освіта...\n\n\n\n"
        "12. Капітана Петрова...\n\n\n"
        "13. Майора Сидорова..."
    )
    assert ensure_blank_line_before_items(multiple_spaces) == (
        "11. Старшого лейтенанта Іванова...\n"
        "1995 р.н., освіта...\n\n"
        "12. Капітана Петрова...\n\n"
        "13. Майора Сидорова..."
    )

    # 4. Перед підписантом залишається рівно 2 порожні рядки
    with_signer = (
        "11. Старшого лейтенанта Іванова...\n\n"
        "12. Капітана Петрова...\n\n\n\n\n"
        "Командир 10 армійського корпусу\n"
        "генерал-майор Іван ПЕТРЕНКО"
    )
    assert ensure_blank_line_before_items(with_signer) == (
        "11. Старшого лейтенанта Іванова...\n\n"
        "12. Капітана Петрова...\n\n\n"
        "Командир 10 армійського корпусу\n"
        "генерал-майор Іван ПЕТРЕНКО"
    )
