"""Межі тіла наказу для примірника.

Примірник має містити пункти РАЗОМ ІЗ ПІДПИСАНТОМ, але БЕЗ службової
таблиці розсилки. Підписант визначається тією самою логікою, що й у витягах.
"""

import importlib.util
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))


def _load_generator_module():
    spec = importlib.util.spec_from_file_location(
        "generate_extracts_span_tests", PROJECT_ROOT / "generate_extracts.py"
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


generator = _load_generator_module()
find_content_start_line = generator.find_content_start_line
find_distribution_cutoff_line = generator.find_distribution_cutoff_line
find_order_signer = generator._find_order_signer

ORDER = "\n".join(
    [
        "МІНІСТЕРСТВО ОБОРОНИ УКРАЇНИ",              # 0
        "НАКАЗ",                                      # 1
        "",                                           # 2
        "§ 1",                                        # 3  ← початок тіла
        "1. Солдата Петренка Петра призначити.",      # 4
        "1980 р.н.",                                  # 5
        "",                                           # 6
        "Командир військової частини А0000",          # 7  ← підписант
        "полковник            Іван ПЕТРЕНКО",         # 8
        "",                                           # 9
        "Розрахунок розсилки",                        # 10 ← відсікається
        "1. Архів",                                   # 11
    ]
)


def test_body_starts_at_paragraph_sign():
    assert find_content_start_line(ORDER) == 3


def test_distribution_table_is_cut_off():
    assert find_distribution_cutoff_line(ORDER) == 10


def test_signer_is_inside_copied_span():
    """Підписант має потрапити у примірник, а не бути відсіченим."""
    signer = find_order_signer(ORDER)
    assert signer is not None
    assert find_content_start_line(ORDER) <= signer["start_line"]
    assert signer["start_line"] < find_distribution_cutoff_line(ORDER)


def test_span_excludes_distribution_but_keeps_signer():
    start = find_content_start_line(ORDER)
    cutoff = find_distribution_cutoff_line(ORDER)
    copied = ORDER.splitlines()[start:cutoff]

    assert any("Командир військової частини" in line for line in copied)
    assert any("ПЕТРЕНКО" in line for line in copied)
    assert not any("Розрахунок розсилки" in line for line in copied)
    assert not any("Архів" in line for line in copied)


def test_order_without_distribution_table_keeps_everything():
    text = "\n".join(["§ 1", "1. Пункт.", "", "Командир", "полковник   Іван ПЕТРЕНКО"])
    assert find_distribution_cutoff_line(text) == len(text.splitlines())
