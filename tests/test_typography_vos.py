"""Тести нерозривних пробілів: код ВОС не повинен розриватись між рядками."""

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
        "generate_extracts_typography_tests", PROJECT_ROOT / "generate_extracts.py"
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


generator = _load_generator_module()
apply_ukrainian_typography = generator.apply_ukrainian_typography

NBSP = " "


@pytest.mark.parametrize(
    "source",
    [
        "ВОС - 0602002.",
        "ВОС – 030400.",
        "ВОС — 0930002.",
        "ВОС-2905001.",
        "вос - 0602002.",
    ],
)
def test_vos_code_never_breaks(source):
    """Між «ВОС», тире та номером мають бути нерозривні пробіли."""
    result = apply_ukrainian_typography(source)
    assert " -" not in result.replace(NBSP, " "), result
    assert NBSP in result


def test_vos_keeps_the_number_attached():
    result = apply_ukrainian_typography("ВОС - 0602002.")
    assert result.startswith(f"ВОС{NBSP}-{NBSP}0602002")


def test_vos_rule_does_not_touch_other_dashes():
    result = apply_ukrainian_typography("Строк - 5 років.")
    assert f"Строк{NBSP}-{NBSP}5" not in result


def test_short_conjunctions_still_get_nbsp():
    """Наявне правило для «та», «із», «і» має продовжувати працювати."""
    result = apply_ukrainian_typography("та подання із документів і копій")
    assert f"та{NBSP}" in result
    assert f"із{NBSP}" in result
    assert f"і{NBSP}" in result


def test_empty_input_is_safe():
    assert apply_ukrainian_typography("") == ""
    assert apply_ukrainian_typography(None) == ""
