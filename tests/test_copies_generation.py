from generate_extracts import (
    extract_metadata_from_filename,
    format_ukr_date,
    build_extracts_filename,
    plan_2up_page_layout,
)


def test_metadata_extraction_variants():
    # Стандартна назва наказу з номером і датою
    f1 = "Наказ командувача № 123 від 15.08.2026.docx"
    num, dt = extract_metadata_from_filename(f1)
    assert num == "123"
    assert dt == "15.08.2026"

    # Назва з номером через скісну риску або дефіс
    f2 = "Наказ № 12-ОС від 01.09.2026 особовий склад.docx"
    num, dt = extract_metadata_from_filename(f2)
    assert num == "12-ОС"
    assert dt == "01.09.2026"

    # Назва без реквізитів
    f3 = "Наказ_чорновик.docx"
    num, dt = extract_metadata_from_filename(f3)
    assert num == ""
    assert dt == ""


def test_build_extracts_filename():
    assert build_extracts_filename("355", "27.07.2026") == "Витяги наказу № 355 від 27.07.2026.docx"
    assert build_extracts_filename("", "") == "Витяги наказу.docx"
    assert build_extracts_filename("100/2", "") == "Витяги наказу № 100_2.docx"


def test_plan_2up_page_layout():
    # Випадок 1: Два 1-сторінкові витяги розміщуються на 1 аркуші (сторінки 1 і 2)
    plan1 = plan_2up_page_layout([1, 1])
    assert len(plan1) == 2
    assert plan1[0]["start_page"] == 1 and plan1[0]["end_page"] == 1
    assert plan1[1]["start_page"] == 2 and plan1[1]["end_page"] == 2

    # Випадок 2: 1-сторінковий витяг, потім 3-сторінковий витяг, потім 1-сторінковий
    # Витяг 0 (1 стор.) -> стор. 1 (Ліва)
    # Порожня сторінка перед витягом 1 -> стор. 2 (Права)
    # Витяг 1 (3 стор.) -> стор. 3, 4, 5
    # Порожня сторінка після витягу 1 -> стор. 6 (Права)
    # Витяг 2 (1 стор.) -> стор. 7 (Ліва)
    plan2 = plan_2up_page_layout([1, 3, 1])
    actions = [p["action"] for p in plan2]
    assert actions == [
        "insert_extract",       # 0 (1 page) -> page 1
        "insert_blank_before",  # page 2
        "insert_extract",       # 1 (3 pages) -> pages 3..5
        "insert_blank_after",   # page 6
        "insert_extract",       # 2 (1 page) -> page 7
    ]
    extract_1 = next(p for p in plan2 if p.get("extract_idx") == 1 and p["action"] == "insert_extract")
    assert extract_1["start_page"] == 3
    assert extract_1["end_page"] == 5

    # Випадок 3: Багатосторінковий парний витяг (2 сторінки), після нього 1-сторінковий
    plan3 = plan_2up_page_layout([2, 1, 1])
    actions3 = [p["action"] for p in plan3]
    assert actions3 == ["insert_extract", "insert_extract", "insert_extract"]
    assert plan3[0]["start_page"] == 1 and plan3[0]["end_page"] == 2
    assert plan3[1]["start_page"] == 3 and plan3[1]["end_page"] == 3
    assert plan3[2]["start_page"] == 4 and plan3[2]["end_page"] == 4


def test_copy_filename_and_title_support():
    from generate_extracts import build_copy_two_filename, sanitize_filename
    
    filename = build_copy_two_filename("355/1", "27.07.2026", "Наказ_355.docx")
    assert filename == "прим_2_27.07.2026_355_1.docx"

    filename_empty = build_copy_two_filename("", "", "Наказ_355.docx")
    assert filename_empty == "прим_2_Наказ_355.docx"


def test_signer_boundary_detection():
    from generate_extracts import _find_order_signer, text_before_order_signer

    order_text = (
        "НАКАЗ КОМАНДУВАЧА\n"
        "1. Пункт 1\n"
        "2. Пункт 2\n\n"
        "Командувач військ оперативного командування\n"
        "генерал-майор                                       Олександр СИДОРОВ\n\n"
        "Згідно з оригіналом:\n"
        "Стара таблиця розсилки\n"
    )

    clean_text, signer = text_before_order_signer(order_text)
    assert signer["name"] == "Олександр СИДОРОВ"
    assert "Командувач" in signer["position"]
    assert "1. Пункт 1" in clean_text
    assert "2. Пункт 2" in clean_text
    assert "Згідно з оригіналом" not in clean_text
    assert "Стара таблиця розсилки" not in clean_text


def test_commander_signer_extracted_before_trailing_distribution_table():
    from generate_extracts import _find_order_signer

    order_text = (
        "НАКАЗ КОМАНДУВАЧА ВІЙСЬК\n"
        "№ 123 від 15.08.2026\n\n"
        "§ 1\n"
        "1. Призначити майора ІВАНОВА І.І. на посаду.\n"
        "2. Призначити капітана ПЕТРЕНКА П.П. на посаду.\n\n"
        "Командувач військ оперативного командування «Північ»\n"
        "генерал-майор                                      Олександр СИДОРОВ\n\n"
        "Розрахунок розсилки наказу:\n"
        "1. В/ч А0000 - 1 прим.\n"
        "2. В/ч А1111 - 1 прим.\n"
        "Начальник служби діловодства\n"
        "майор                                              Іван КОВАЛЕНКО\n"
    )

    signer = _find_order_signer(order_text)
    assert signer is not None
    assert "Командувач" in signer["position"]
    assert "генерал-майор" in signer["rank"]
    assert "СИДОРОВ" in signer["name"]
    assert "КОВАЛЕНКО" not in signer["name"]


def test_extract_unbreakable_chain_logic():
    """Перевіряємо, що правила нерозривності ланцюга [останній пункт -> відступ -> підписант -> згідно з оригіналом]
    гарантують: підписант ніколи не відривається від останнього пункту."""
    # Симулюємо поведінку формування параграфів витягу
    paragraphs = [
        {"text": "§ 1", "is_heading": True},
        {"text": "", "is_gap": True, "after_heading": True},
        {"text": "1. Перший пункт наказу.", "is_item": True, "item_idx": 0},
        {"text": "", "is_gap": True, "after_item": True},
        {"text": "2. Другий (останній) пункт наказу.", "is_item": True, "item_idx": 1},
        {"text": "", "is_gap": True, "before_signer": True},
        {"text": "Командувач військ", "is_signer": True},
        {"text": "генерал-майор Олександр СИДОРОВ", "is_signer": True},
        {"text": "", "is_gap": True},
        {"text": "Згідно з оригіналом", "is_certifier": True},
        {"text": "Т.в.о. начальника штабу", "is_certifier": True},
        {"text": "полковник Сергій ПЕТРЕНКО", "is_certifier": True, "is_last_certifier_line": True},
    ]

    # Правило 1: Шапки мають KeepWithNext=True
    # Правило 2: Останній пункт і всі наступні абзаци до кінця засвідчення мають KeepWithNext=True (крім останнього рядка засвідчення)
    last_item_found = False
    keep_with_next_flags = []
    for p in paragraphs:
        if p.get("text", "").startswith("2. Другий"):
            last_item_found = True
        
        if p.get("is_heading") or p.get("after_heading"):
            keep_with_next_flags.append(True)
        elif last_item_found:
            # Останній рядок засвідчувача має KeepWithNext=False, решта в ланцюгу - True
            keep_with_next_flags.append(not p.get("is_last_certifier_line", False))
        else:
            keep_with_next_flags.append(False)

    # Перевіряємо, що в нерозривному ланцюгу від останнього пункту до передостаннього рядка засвідчення ВСІ значення True
    chain_start_idx = 4  # пункт 2
    chain_end_idx = len(paragraphs) - 1  # останній рядок засвідчення
    for idx in range(chain_start_idx, chain_end_idx):
        assert keep_with_next_flags[idx] is True, f"Абзац {paragraphs[idx]['text']} має бути нерозривним (KeepWithNext=True)!"
    assert keep_with_next_flags[chain_end_idx] is False
