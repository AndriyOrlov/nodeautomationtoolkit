"""Тести людських пояснень помилок.

Користувач бачить не код збою, а три речі: що сталося, з яким файлом і що
робити далі. Технічний текст лишається в кінці — для розробника.
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
        "generate_extracts_error_tests", PROJECT_ROOT / "generate_extracts.py"
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


generator = _load_generator_module()
explain_error = generator.explain_error
UserError = generator.UserError


def test_user_error_reads_as_what_and_what_to_do():
    error = UserError("немає файлу зразка", "виберіть його заново")
    text = str(error)
    assert text.startswith("Що сталося: немає файлу зразка")
    assert "Що зробити: виберіть його заново" in text


def test_missing_file_without_name_is_still_explained():
    """Головний випадок: Windows не додає назви файлу до «[WinError 2]»."""
    error = FileNotFoundError(2, "The system cannot find the file specified")

    text = explain_error(error)

    assert "не знайшов" in text
    assert "WinError" not in text.split("технічні подробиці")[0]
    assert "Що зробити:" in text


def test_missing_file_with_name_names_it():
    error = FileNotFoundError(2, "The system cannot find the file specified")
    error.filename = r"C:\Робочий стіл\зразок витягу.docx"

    text = explain_error(error)

    assert "«зразок витягу.docx»" in text
    assert r"C:\Робочий стіл\зразок витягу.docx" in text


def test_busy_file_tells_to_close_word():
    error = PermissionError(13, "Permission denied")
    error.filename = r"C:\Вихід\Витяги.docx"

    text = explain_error(error)

    assert "зайнятий" in text
    assert "Word" in text


def test_disconnected_drive_is_named_as_such():
    error = OSError(21, "The device is not ready")
    error.winerror = 21

    text = explain_error(error)

    assert "диск" in text.casefold()


def test_word_com_failure_is_explained_in_plain_words():
    class com_error(Exception):  # ім'я класу — те саме, що в pywin32
        pass

    text = explain_error(com_error("(-2147418111, 'Call was rejected by callee.')"))

    assert "Word" in text
    assert "Що зробити:" in text


def test_unknown_error_keeps_technical_text_for_the_developer():
    text = explain_error(ZeroDivisionError("division by zero"))

    assert "несподівана помилка" in text
    assert "ZeroDivisionError: division by zero" in text


def test_user_error_passes_through_unchanged():
    error = UserError("немає зразка", "виберіть заново")
    assert explain_error(error) == str(error)
