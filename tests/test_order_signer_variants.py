"""Форми підписного блока наказу, які мають розпізнаватись.

Наказ підписує не лише «Т.в.о.»: буває «Командувач …», «В.о. командувача …»,
а подекуди звання стоїть у тому самому рядку, що й посада. Якщо блок не
впізнано, тіло наказу не обрізається і службовий хвіст протікає у витяг.

Окремо закріплено дворядковий блок: його другий рядок починається з родового
відмінка («командувача військ»), і цей рядок НЕ має вважатися початком блока —
`_analyze_order` шукає підписанта з кінця й узяв би саме його.

Прізвища вигадані.
"""

import importlib.util
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))


def _load_generator():
    spec = importlib.util.spec_from_file_location(
        "generate_extracts_signer_variants", PROJECT_ROOT / "generate_extracts.py"
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


generator = _load_generator()

BODY = (
    "§ 1\n\n"
    "1. Капітана ТЕСТЕНКА Андрія, командира роти 3 окремого тестового загону.\n\n"
)
TAIL = "\nРозрахунок розсилки витягів із наказу:\n1. Військова частина А0003 п. 1.\n"


def _signer(block: str) -> dict:
    return generator._find_order_signer(BODY + block + TAIL) or {}


def test_commander_two_line_block():
    signer = _signer("Командувач військ оперативного командування «Захід»\n\n"
                     "генерал-лейтенант     Іван ТЕСТЕНКО\n")

    assert signer.get("rank") == "генерал-лейтенант"
    assert signer.get("name") == "Іван ТЕСТЕНКО"
    assert signer.get("position", "").startswith("Командувач")


def test_acting_commander_with_wrapped_position():
    """Дворядкова посада: початком блока має бути ПЕРШИЙ рядок."""
    signer = _signer("Тимчасово виконуючий обов'язки\nкомандувача військ ОК «Захід»\n\n"
                     "полковник     Петро ЗАСТУПНИК\n")

    assert signer.get("name") == "Петро ЗАСТУПНИК"
    assert signer.get("position", "").startswith("Тимчасово виконуючий")


def test_acting_form_without_leading_t():
    signer = _signer("В.о. командувача військ ОК «Захід»\n\nполковник     Петро ВИКОНУВАЧ\n")

    assert signer.get("name") == "Петро ВИКОНУВАЧ"


def test_rank_on_the_same_line_as_position():
    signer = _signer("Командувач військ ОК «Захід» генерал-майор Степан ОДНОРЯДКОВ\n")

    assert signer.get("rank") == "генерал-майор"
    assert signer.get("name") == "Степан ОДНОРЯДКОВ"


def test_continuation_line_is_not_a_block_start():
    """«командувача військ» саме по собі початком підписанта не є."""
    assert not generator._ORDER_SIGNER_START_RE.match("командувача військ")
    assert not generator._ORDER_SIGNER_START_RE.match("начальника штабу")
    assert generator._ORDER_SIGNER_START_RE.match("Командувач військ ОК «Захід»")
