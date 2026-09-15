"""Тести іменування папок пакетної генерації примірників.

Номер наказу може містити «/» («б/н», «123/45»). ОС трактує його як
роздільник шляху, тож без очищення замість однієї папки створювались
вкладені — і примірник опинявся не там, де очікувалось.
"""

import importlib.util
import os
import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))


def _load_generator_module():
    spec = importlib.util.spec_from_file_location(
        "generate_extracts_batch_tests", PROJECT_ROOT / "generate_extracts.py"
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


generator = _load_generator_module()
sanitize_filename = generator.sanitize_filename
extract_metadata_from_filename = generator.extract_metadata_from_filename
build_copy_two_filename = generator.build_copy_two_filename


def _sub_folder_name(order_num: str, fname: str) -> str:
    """Повторює правило іменування підпапки з run_generate_copies."""
    if order_num:
        return f"Наказ № {sanitize_filename(order_num)}"
    return sanitize_filename(os.path.splitext(fname)[0])


@pytest.mark.parametrize(
    "order_num",
    ["б/н", "123/45", "12\\34", "б/н:1"],
)
def test_folder_name_never_contains_path_separators(order_num):
    folder = _sub_folder_name(order_num, "Наказ.docx")
    assert "/" not in folder
    assert "\\" not in folder


def test_slash_number_stays_single_folder():
    """«б/н» має дати ОДНУ папку, а не вкладені «Наказ № б» / «н»."""
    folder = _sub_folder_name("б/н", "Наказ.docx")
    joined = os.path.join("out", folder)
    assert os.path.dirname(joined) == "out"


def test_plain_number_is_unchanged():
    assert _sub_folder_name("418", "Наказ.docx") == "Наказ № 418"


def test_folder_and_file_use_same_sanitisation():
    fname = "Наказ_№ б/н від 20.05.2025.docx"
    order_num, order_date = extract_metadata_from_filename(fname)

    folder = _sub_folder_name(order_num, fname)
    file_name = build_copy_two_filename(order_num, order_date, fname)

    for forbidden in ("/", "\\", ":", "*", "?", '"', "<", ">", "|"):
        assert forbidden not in folder
        assert forbidden not in file_name


def test_fallback_folder_from_filename_is_sanitised():
    """Без номера папка береться з назви файлу — її теж треба чистити."""
    folder = _sub_folder_name("", "Наказ 1/2 без номера.docx")
    assert "/" not in folder


def test_sanitize_keeps_readable_number():
    assert sanitize_filename("б/н") == "б_н"
    assert sanitize_filename("123/45") == "123_45"
