"""Тести формату дати наказу в повідомленнях.

У повідомленнях дата НЕ розкривається словами — на відміну від витягів,
де діє військовий стандарт «“20” травня 2025 року».
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
        "generate_extracts_date_tests", PROJECT_ROOT / "generate_extracts.py"
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


generator = _load_generator_module()
format_message_date = generator.format_message_date
format_ukr_date = generator.format_ukr_date


@pytest.mark.parametrize(
    "raw, expected",
    [
        ("20.05.2025", "20.05.2025 року"),
        ("01.01.2026", "01.01.2026 року"),
        ("31.12.2024", "31.12.2024 року"),
    ],
)
def test_message_date_stays_numeric(raw, expected):
    assert format_message_date(raw) == expected


@pytest.mark.parametrize("raw", ["", "   ", "не дата", "2025-05-20", None])
def test_message_date_returns_empty_on_bad_input(raw):
    assert format_message_date(raw) == ""


def test_message_date_keeps_leading_zeros():
    assert format_message_date("05.06.2025") == "05.06.2025 року"


def test_extracts_date_format_is_unchanged():
    """Витяги мають зберегти військовий стандарт із назвою місяця."""
    result = format_ukr_date("20.05.2025")
    assert "травня" in result
    assert "2025 року" in result
    assert "20.05.2025" not in result


def test_message_and_extract_formats_differ():
    assert format_message_date("20.05.2025") != format_ukr_date("20.05.2025")
