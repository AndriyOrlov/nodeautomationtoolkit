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


def test_back_page_tags_use_only_filename_metadata_and_copy_two_label():
    assert back_page_tag_values("396", "06.08.2026") == {
        "{{згідно_з_оригіналом}}": "Згідно з оригіналом",
        "{{примірник}}": "Примірник № 2",
        "{{номер_наказу}}": "396",
        "{{дата_наказу}}": "“06” серпня 2026 року",
    }
    assert back_page_tag_values("", "") == {
        "{{згідно_з_оригіналом}}": "Згідно з оригіналом",
        "{{примірник}}": "Примірник № 2",
    }


def test_copy_two_filename_does_not_invent_missing_order_metadata():
    assert build_copy_two_filename("396", "06.08.2026", "source.docx") == "прим_2_06.08.2026_396.docx"
    assert build_copy_two_filename("", "", "source.docx") == "прим_2_source.docx"


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


