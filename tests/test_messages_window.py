"""Окрема програма «Повідомлення»: ті самі поля, що й у великій програмі.

Word не запускається, конфіг користувача не читається; дані вигадані.
"""

import importlib.util
import os
import sys
from pathlib import Path

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
pytest.importorskip("PySide6")

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT / "src") not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT / "src"))

from PySide6.QtWidgets import QApplication, QCheckBox, QLineEdit  # noqa: E402

from nodeautomationtoolkit.generator_qt.main_window import install_qt_bridge  # noqa: E402
from nodeautomationtoolkit.generator_qt.messages_window import (  # noqa: E402
    SettingsDialog,
    create_messages_app_class,
)


@pytest.fixture(scope="module")
def messages_app():
    qt_app = QApplication.instance() or QApplication([])
    spec = importlib.util.spec_from_file_location(
        "generate_extracts_messages_window_tests", PROJECT_ROOT / "generate_extracts.py"
    )
    legacy = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(legacy)
    install_qt_bridge(legacy)

    class Isolated(create_messages_app_class(legacy)):
        def load_config(self):
            return None

        def save_config(self):
            return None

    app = Isolated()
    yield app
    app.main_window.close()
    qt_app.processEvents()


def _edits(widget, text):
    return [edit for edit in widget.findChildren(QLineEdit) if edit.text() == text]


def test_main_window_has_the_same_run_fields_as_the_big_program(messages_app):
    app = messages_app
    for variable, value in (
        (app.message_executor, "Виконавець-тест"),
        (app.order_signer_position, "Командир-тест"),
        (app.order_signer_rank, "полковник-тест"),
        (app.order_signer_name, "Підписант ТЕСТОВИЙ"),
        (app.certifier_position, "Засвідчувач-посада"),
        (app.certifier_rank, "майор-тест"),
        (app.certifier_name, "Засвідчувач ТЕСТОВИЙ"),
        (app.message_out_folder, "C:/Тест/Повідомлення"),
    ):
        variable.set(value)
        assert _edits(app.main_window, value), f"у головному вікні немає поля зі значенням {value!r}"
    assert any(
        "внутрішнє переміщення в управлінні" in box.text()
        for box in app.main_window.findChildren(QCheckBox)
    )


def test_settings_dialog_keeps_only_table_and_samples(messages_app):
    app = messages_app
    app.message_executor.set("лише-в-головному-вікні")
    app.excel_path.set("C:/Тест/таблиця.xlsx")
    dialog = SettingsDialog(app, app.main_window)
    try:
        assert _edits(dialog, "C:/Тест/таблиця.xlsx")
        assert not _edits(dialog, "лише-в-головному-вікні")
    finally:
        dialog.deleteLater()


def test_new_order_clears_the_previous_signer(messages_app, monkeypatch, tmp_path):
    app = messages_app
    app.order_signer_name.set("Підписант ПОПЕРЕДНЬОГО наказу")
    order = tmp_path / "Наказ № 1 від 01.10.2026.docx"
    order.write_bytes(b"")
    # Word не запускаємо: у тексті «наказу» підписанта немає.
    monkeypatch.setattr(type(app), "_read_word_text", lambda self, path, normalize=True: "1. Пункт.")
    app._set_order(str(order))
    assert app.order_signer_name.get() == ""
