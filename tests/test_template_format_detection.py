"""Тести визначення справжнього формату шаблонів Word.

Word відмовляється відкривати файл, якщо його вміст не відповідає розширенню
(«формат і розширення файлу не збігаються»). Робоча копія має отримувати
розширення, що відповідає РЕАЛЬНОМУ вмісту файлу.
"""

import importlib.util
import sys

import pytest
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))


def _load_generator_module():
    spec = importlib.util.spec_from_file_location(
        "generate_extracts_format_tests", PROJECT_ROOT / "generate_extracts.py"
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


generator = _load_generator_module()
detect_word_extension = generator.detect_word_extension
copy_template_for_editing = generator.copy_template_for_editing
is_path_writable = generator.is_path_writable

_OOXML_SIGNATURE = b"PK\x03\x04" + b"\x00" * 32
_OLE2_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 32
_RTF_SIGNATURE = b"{\\rtf1\\ansi" + b" " * 32


def test_detects_ooxml_regardless_of_extension(tmp_path):
    path = tmp_path / "template.doc"
    path.write_bytes(_OOXML_SIGNATURE)
    assert detect_word_extension(str(path)) == ".docx"


def test_detects_legacy_doc_named_as_docx(tmp_path):
    """Головний випадок помилки: вміст Word 97-2003 під розширенням .docx."""
    path = tmp_path / "template.docx"
    path.write_bytes(_OLE2_SIGNATURE)
    assert detect_word_extension(str(path)) == ".doc"


def test_detects_rtf(tmp_path):
    path = tmp_path / "template.docx"
    path.write_bytes(_RTF_SIGNATURE)
    assert detect_word_extension(str(path)) == ".rtf"


def test_unknown_content_falls_back_to_extension(tmp_path):
    path = tmp_path / "template.docx"
    path.write_bytes(b"just some bytes")
    assert detect_word_extension(str(path)) == ".docx"


def test_missing_file_falls_back_to_extension(tmp_path):
    assert detect_word_extension(str(tmp_path / "absent.docx")) == ".docx"


def test_copy_keeps_output_path_when_format_matches(tmp_path):
    template = tmp_path / "template.docx"
    template.write_bytes(_OOXML_SIGNATURE)
    output = tmp_path / "result.docx"

    working = copy_template_for_editing(str(template), str(output))

    assert working == str(output)
    assert Path(working).read_bytes() == _OOXML_SIGNATURE


def test_copy_uses_real_extension_when_format_differs(tmp_path):
    """Шаблон Word 97-2003 не можна копіювати одразу в .docx."""
    template = tmp_path / "template.docx"
    template.write_bytes(_OLE2_SIGNATURE)
    output = tmp_path / "result.docx"

    working = copy_template_for_editing(str(template), str(output))

    assert working == str(tmp_path / "result.doc")
    assert Path(working).exists()
    assert Path(working).read_bytes() == _OLE2_SIGNATURE


def test_missing_file_is_writable(tmp_path):
    """Неіснуючий файл можна створити — генерація не має блокуватись."""
    assert is_path_writable(str(tmp_path / "ще-не-створений.docx")) is True


def test_existing_free_file_is_writable(tmp_path):
    path = tmp_path / "result.docx"
    path.write_bytes(_OOXML_SIGNATURE)
    assert is_path_writable(str(path)) is True


def test_locked_file_is_detected(tmp_path):
    """Файл, який тримає інший процес, має розпізнаватись як недоступний.

    Word блокує документ режимом спільного доступу Windows (а не діапазонним
    блокуванням), тому імітуємо саме його: CreateFile із dwShareMode = 0.
    """
    win32con = pytest.importorskip("win32con")
    win32file = pytest.importorskip("win32file")

    path = tmp_path / "locked.docx"
    path.write_bytes(_OOXML_SIGNATURE)
    handle = win32file.CreateFile(
        str(path),
        win32con.GENERIC_READ | win32con.GENERIC_WRITE,
        0,  # жодного спільного доступу — так само тримає файл Word
        None,
        win32con.OPEN_EXISTING,
        0,
        None,
    )
    try:
        assert is_path_writable(str(path)) is False
    finally:
        handle.Close()


def test_copy_does_not_modify_original_template(tmp_path):
    template = tmp_path / "template.docx"
    template.write_bytes(_OLE2_SIGNATURE)
    output = tmp_path / "result.docx"

    copy_template_for_editing(str(template), str(output))

    assert template.read_bytes() == _OLE2_SIGNATURE
