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


def test_extended_chain_checkbox_is_off_by_default_and_drives_the_option(messages_app):
    app = messages_app
    boxes = [box for box in app.main_window.findChildren(QCheckBox) if box.text() == "Подовжена версія"]
    assert len(boxes) == 1
    assert boxes[0].isChecked() is False
    assert app._message_chain_options() == {"full_chain": False, "high_commands": None}
    boxes[0].setChecked(True)
    QApplication.processEvents()
    assert app.message_full_chain.get() is True
    assert app._message_chain_options()["full_chain"] is True
    boxes[0].setChecked(False)


def test_management_extracts_go_to_the_messages_folder(messages_app, monkeypatch, tmp_path):
    app = messages_app
    template = tmp_path / "zrazok.docx"
    template.write_bytes(b"")
    calls = []

    def fake_messages(self):
        self.last_messages_created = outcome["ok"]

    monkeypatch.setattr(app.legacy.App, "run_generate_messages", fake_messages)
    monkeypatch.setattr(
        app, "run_extracts",
        lambda scope, notify_empty=True: calls.append(
            (scope, notify_empty, app.out_folder.get(), app.executor.get())
        ),
    )
    monkeypatch.setattr(app, "_settings_missing", lambda: [])
    # Розрахунок розсилки тут — лише Word-файл повідомлень (>10 адресатів), не XLSX витягів.
    monkeypatch.setattr(app, "run_rozrahunok", lambda: calls.append(("xlsx-calc",)))
    for variable, value in (
        (app.template_path, str(template)),
        (app.message_out_folder, str(tmp_path / "Повідомлення")),
        (app.out_folder, "C:/Готово-тест"),
        (app.executor, ""),
        (app.message_executor, "Виконавець-тест"),
    ):
        variable.set(value)

    outcome = {"ok": False}
    app.run_generate_messages()
    assert calls == []  # повідомлення не створено — витягів теж немає

    outcome["ok"] = True
    app.run_generate_messages()
    # Обидві галочки (за замовчуванням) — усі витяги: адресні й до управління.
    assert calls == [("all", False, str(tmp_path / "Повідомлення"), "Виконавець-тест")]
    # Папка й виконавець великої програми не змінилися.
    assert app.out_folder.get() == "C:/Готово-тест"
    assert app.executor.get() == ""

    app.template_path.set("")
    calls.clear()
    app.run_generate_messages()
    assert calls == []  # без зразка витягу — лише запис у журналі

    app.template_path.set(str(template))
    for general, management, expected in (
        (True, False, ["general"]),
        (False, True, ["management"]),
        (False, False, []),  # обидві зняті — лише повідомлення
    ):
        calls.clear()
        app.message_general_extracts.set(general)
        app.message_management_extracts.set(management)
        app.run_generate_messages()
        assert [call[0] for call in calls] == expected
    app.message_general_extracts.set(True)
    app.message_management_extracts.set(True)
    app.template_path.set("")


def test_extracts_checkboxes_are_on_by_default(messages_app):
    app = messages_app
    for caption, variable in (
        ("Звичайні витяги", app.message_general_extracts),
        ("Витяги до управління", app.message_management_extracts),
    ):
        boxes = [box for box in app.main_window.findChildren(QCheckBox) if box.text() == caption]
        assert len(boxes) == 1 and boxes[0].isChecked(), caption
        assert variable.get() is True


def test_settings_have_the_extract_template_row(messages_app):
    app = messages_app
    dialog = SettingsDialog(app, app.main_window)
    app.template_path.set("C:/zrazok-vytiahu-test.docx")
    assert _edits(dialog, "C:/zrazok-vytiahu-test.docx")
    dialog.close()
    app.template_path.set("")


def test_signer_cut_keeps_the_signer_fields(messages_app):
    app = messages_app
    app.order_signer_name.set("Поле-тест")
    usable, _signer = app._refresh_order_signer("НАКАЗУЮ:\n1. Пункт.\n")
    assert "1. Пункт." in usable
    assert app.order_signer_name.get() == "Поле-тест"
    app.order_signer_name.set("")
