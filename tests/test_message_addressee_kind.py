"""Тести тегу {{тцк чі вч}} — типу адресата повідомлення про кадрові рішення."""

import importlib.util
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))


def _load_generator_module():
    """Імпортує generate_extracts без запуску Tkinter-вікна."""
    spec = importlib.util.spec_from_file_location(
        "generate_extracts_for_tests", PROJECT_ROOT / "generate_extracts.py"
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


generator = _load_generator_module()
build_addressee_kind_text = generator.build_addressee_kind_text


def test_only_units_gives_commanders():
    groups = {"corps": [], "units": ["Командиру військової частини А1111"], "tck": []}
    assert build_addressee_kind_text(groups) == "Командирам військових частин"


def test_only_corps_still_counts_as_units():
    groups = {"corps": ["Командиру 10 армійського корпусу"], "units": [], "tck": []}
    assert build_addressee_kind_text(groups) == "Командирам військових частин"


def test_only_tck_gives_heads():
    groups = {"corps": [], "units": [], "tck": ["Начальнику Львівського ОТЦК та СП"]}
    assert build_addressee_kind_text(groups) == "Начальникам ТЦК"


def test_units_and_tck_gives_both():
    groups = {
        "corps": [],
        "units": ["Командиру військової частини А1111"],
        "tck": ["Начальнику Львівського ОТЦК та СП"],
    }
    assert (
        build_addressee_kind_text(groups)
        == "Командирам військових частин та Начальникам ТЦК"
    )


def test_corps_units_and_tck_gives_both():
    groups = {
        "corps": ["Командиру 10 армійського корпусу"],
        "units": ["Командиру військової частини А1111"],
        "tck": ["Начальнику Львівського ОТЦК та СП"],
    }
    assert (
        build_addressee_kind_text(groups)
        == "Командирам військових частин та Начальникам ТЦК"
    )


def test_no_recipients_gives_empty_text():
    assert build_addressee_kind_text({"corps": [], "units": [], "tck": []}) == ""


def test_missing_keys_do_not_raise():
    assert build_addressee_kind_text({}) == ""


def test_recipient_list_keeps_corps_units_tck_order():
    """Плоский список має лишатися сумісним із {{кому_список}}."""
    groups = {"corps": ["К"], "units": ["Ч"], "tck": ["Т"]}
    flat = groups["corps"] + groups["units"] + groups["tck"]
    assert flat == ["К", "Ч", "Т"]


def test_addressee_kind_tags_cover_both_spellings():
    assert "{{тцк чі вч}}" in generator._MESSAGE_ADDRESSEE_KIND_TAGS
    assert "{{тцк чи вч}}" in generator._MESSAGE_ADDRESSEE_KIND_TAGS
