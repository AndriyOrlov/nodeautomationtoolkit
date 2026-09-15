"""Пошук підписанта має витримувати ДОВГИЙ службовий хвіст наказу.

Відколи текст збирається з абзаців (`read_document_text`), кожна комірка
таблиці стала окремим рядком, і службова частина займає значно більше рядків.
З вузьким вікном пошуку маркер лишався поза ним: підписант не визначався,
і останній пункт «затягував» увесь службовий хвіст у витяг.

Тепер вікна немає взагалі: маркер шукається по всьому документу (з кінця, тож
перемагає останній), а підписант — між маркером і ОСТАННІМ пронумерованим
пунктом. Довжина хвоста більше ні на що не впливає, а рядок тіла наказу, що
починається з «Начальник…», не може бути прийнятий за підписанта.
"""

import importlib.util
import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))


def _load_generator_module():
    spec = importlib.util.spec_from_file_location(
        "generate_extracts_signer_window_tests", PROJECT_ROOT / "generate_extracts.py"
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


generator = _load_generator_module()

BODY = ["§ 1", "1. Перший пункт.", "", "2. Останній пункт.", ""]
SIGNER = [
    "Тимчасово виконуючий обов'язки",
    "командувача військ",
    "полковник   І. ПЕТРЕНКО",
    "",
]


def _order(tail_lines: int) -> str:
    tail = ["Розрахунок розсилки витягів із наказу:"]
    tail += [f"комірка {index}" for index in range(tail_lines)]
    return "\n".join(BODY + SIGNER + tail)


@pytest.mark.parametrize("tail_lines", [0, 50, 200, 350, 1000])
def test_cutoff_is_found_however_long_the_tail(tail_lines):
    text = _order(tail_lines)
    assert generator.find_distribution_cutoff_line(text) == len(BODY) + len(SIGNER)


@pytest.mark.parametrize("tail_lines", [0, 50, 200, 350, 1000])
def test_signer_is_found_however_long_the_tail(tail_lines):
    signer = generator._find_order_signer(_order(tail_lines))
    assert signer is not None
    assert signer["start_line"] == len(BODY)


def test_service_tail_never_leaks_into_the_extract():
    """Головний симптом: службовий хвіст опинявся в останньому пункті."""
    body_text, _ = generator.text_before_order_signer(_order(200))

    assert "Розрахунок розсилки" not in body_text
    assert "комірка" not in body_text
    assert "2. Останній пункт." in body_text


def test_signer_stays_out_of_the_body_text():
    body_text, signer = generator.text_before_order_signer(_order(50))

    assert "ПЕТРЕНКО" not in body_text
    assert signer["name"]


def test_first_valid_signer_after_last_item_is_the_cutoff():
    second_signer = [
        "Тимчасово виконуючий обов'язки командувача військ оперативного командування",
        "полковник   О. СИДОРЕНКО",
    ]
    text = "\n".join(
        BODY
        + SIGNER
        + ["СЛУЖБОВИЙ ТЕКСТ МІЖ ПІДПИСНИМИ БЛОКАМИ"]
        + second_signer
        + ["Згідно з оригіналом"]
    )

    signer = generator._find_order_signer(text)
    body_text, _ = generator.text_before_order_signer(text)

    assert signer is not None
    assert signer["start_line"] == len(BODY)
    assert "ПЕТРЕНКО" in signer["name"]
    assert "СЛУЖБОВИЙ ТЕКСТ" not in body_text
    assert "СИДОРЕНКО" not in body_text


def test_unparsed_signer_like_block_is_still_removed_completely():
    text = "\n".join(
        BODY
        + [
            "Тимчасово виконуючий обов'язки",
            "командувача військ",
            "НЕРОЗІБРАНИЙ ПІДПИСНИЙ БЛОК",
            "Тимчасово виконуючий обов'язки командувача військ",
            "полковник   О. СИДОРЕНКО",
            "Згідно з оригіналом",
        ]
    )

    body_text, signer = generator.text_before_order_signer(text)

    assert "2. Останній пункт." in body_text
    assert "Тимчасово виконуючий" not in body_text
    assert "НЕРОЗІБРАНИЙ" not in body_text
    # Реквізити можуть бути розібрані окремо, але межа тіла завжди стоїть
    # перед першим підписоподібним блоком.


BODY_WITH_SIGNER_LIKE_LINE = [
    "§ 1",
    "Начальник відділу кадрів",
    "полковник О. КОВАЛЕНКО",
    "1. Перший пункт.",
    "",
    "2. Останній пункт.",
    "",
]


def test_signer_like_line_inside_the_body_is_not_taken_for_the_signer():
    """Межа пошуку — останній пункт, а не лічильник рядків.

    Без такої межі рядок тіла наказу «Начальник відділу кадрів / полковник …»
    ставав «підписантом», і весь наказ обрізався на ньому.
    """
    text = "\n".join(
        BODY_WITH_SIGNER_LIKE_LINE
        + ["Розрахунок розсилки витягів із наказу:"]
        + [f"комірка {index}" for index in range(30)]
    )

    assert generator._find_order_signer(text) is None

    body_text, _ = generator.text_before_order_signer(text)
    assert "1. Перший пункт." in body_text
    assert "2. Останній пункт." in body_text


def test_real_signer_wins_over_signer_like_line_in_the_body():
    text = "\n".join(
        BODY_WITH_SIGNER_LIKE_LINE
        + SIGNER
        + ["Розрахунок розсилки витягів із наказу:"]
        + [f"комірка {index}" for index in range(30)]
    )

    signer = generator._find_order_signer(text)
    assert signer is not None
    assert signer["start_line"] == len(BODY_WITH_SIGNER_LIKE_LINE)
    assert "ПЕТРЕНКО" in signer["name"]


def test_last_item_line_marks_the_end_of_the_order_body():
    lines = BODY_WITH_SIGNER_LIKE_LINE + SIGNER
    assert generator._last_item_line(lines) == 5


def test_last_item_line_falls_back_to_zero_without_numbered_items():
    assert generator._last_item_line(["§ 1", "Текст без пунктів."]) == 0


def test_numbered_rows_of_the_distribution_table_do_not_move_the_body_edge():
    """Рядки таблиці розсилки теж пронумеровані: «1. в/ч А0000 — 1 прим.».

    Якщо рахувати їх пунктами наказу, межа пошуку заїжджає за підписанта —
    і він перестає знаходитися взагалі.
    """
    text = "\n".join(
        BODY
        + SIGNER
        + [
            "Розрахунок розсилки витягів із наказу:",
            "1. в/ч А0000 - 1 прим.",
            "2. в/ч А1111 - 1 прим.",
        ]
    )

    signer = generator._find_order_signer(text)
    assert signer is not None
    assert signer["start_line"] == len(BODY)


# Форма звороту останнього аркуша наказу за офіційним зразком (додаток 43):
# службових блоків ДВА, і рядки кожного пронумеровані так само, як пункти.
OFFICIAL_BACK_PAGE = [
    "Розрахунок розсилки витягів із наказу:",
    "1. Військова частина А0000\tп. 1.\t3. Військова частина А1111\tп. 2.",
    "2. Військова частина А2222\tп. 3.",
    "Розрахунок розсилки електронних повідомлень:",
    "1. Військова частина А0000\tп. 1.\t2. Військова частина А1111\tп. 2.",
    "Витяги з наказу відправлено згідно з розрахунком розсилки.",
    "Надр. 2 прим.",
    "Прим. № 1 — перший адресат.",
    "Прим. № 2 — другий адресат.",
]


def test_two_service_blocks_do_not_hide_the_signer():
    """Зворот останнього аркуша має ДВА «Розрахунок розсилки …».

    Маркер шукається з кінця, тож опорним стає другий блок. Рядки першого
    блоку пронумеровані («1. Військова частина А0000 п. 1.»), і якщо рахувати
    їх пунктами наказу, межа пошуку заїжджає за підписанта: він не знаходиться
    зовсім, а весь службовий хвіст протікає у витяг.
    """
    text = "\n".join(BODY + SIGNER + OFFICIAL_BACK_PAGE)

    signer = generator._find_order_signer(text)
    assert signer is not None
    assert signer["start_line"] == len(BODY)

    body_text, _ = generator.text_before_order_signer(text)
    assert "2. Останній пункт." in body_text
    for service_marker in ("Розрахунок розсилки", "Надр.", "Прим. №"):
        assert service_marker not in body_text


# ── Підписний блок, розсунутий порожніми абзацами ─────────────────────────────
# Підписанта в наказі часто відсувають до низу сторінки порожніми абзацами,
# тож між посадою та званням буває більше восьми рядків. Вікно у 8 рядків
# через це не знаходило підписанта ЗОВСІМ: тіло наказу не обрізалося, і блок
# протікав у {{зміст}}.
SERVICE_TAIL = ["Розрахунок розсилки витягів із наказу:", "1. Військова частина А0002\tп. 1."]


@pytest.mark.parametrize(
    "position_block",
    [
        ["Тимчасово виконуючий обов'язки", "командувача військ"],
        ["Тимчасово виконуючий обов'язки", "командувача військ", "", ""],
        ["Тимчасово виконуючий обов'язки", "командувача військ"] + [""] * 8,
        ["Тимчасово виконуючий", "обов'язки командувача військ"],
    ],
    ids=["суцільно", "порожні-комірки", "звання-далеко", "розрив-у-посаді"],
)
def test_signer_is_found_however_far_the_rank_line_sits(position_block):
    text = "\n".join(
        BODY + position_block + ["полковник   І. ПЕТРЕНКО", ""] + SERVICE_TAIL
    )

    signer = generator._find_order_signer(text)

    assert signer is not None
    assert signer["start_line"] == len(BODY)
    assert signer["rank"] == "полковник"


def test_signer_block_never_leaks_into_the_extract_content():
    """Головний симптом: підписант опинявся у {{зміст}} і дублювався."""
    text = "\n".join(
        BODY
        + ["Тимчасово виконуючий обов'язки", "командувача військ"]
        + [""] * 8
        + ["полковник   І. ПЕТРЕНКО", ""]
        + SERVICE_TAIL
    )

    body_text, signer = generator.text_before_order_signer(text)

    assert signer["name"]
    for leaked in ("Тимчасово виконуючий", "полковник", "Розрахунок розсилки"):
        assert leaked not in body_text
