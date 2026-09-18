"""Вікно «Накази»: замок паролем, поля, чотири кроки.

Генератор наказів схований за кнопкою «🔒 Накази» й тимчасовим паролем, тож
тести спершу відмикають вікно. Word не запускається, конфіг користувача не
читається — усі файли вигадані, у тимчасовій теці.
"""

from __future__ import annotations

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

from PySide6.QtWidgets import QApplication, QLineEdit  # noqa: E402

from nodeautomationtoolkit.generator_qt import compat  # noqa: E402
from nodeautomationtoolkit.generator_qt.main_window import (  # noqa: E402
    create_qt_app_class,
    install_qt_bridge,
)

from test_order_pipeline import PLAN_HEADER, PLAN_ROWS  # noqa: E402


@pytest.fixture(scope="module")
def qt_app():
    return QApplication.instance() or QApplication([])


@pytest.fixture(scope="module")
def legacy(qt_app):
    spec = importlib.util.spec_from_file_location(
        "generate_extracts_orders_tab_tests", PROJECT_ROOT / "generate_extracts.py"
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    install_qt_bridge(module)
    return module


@pytest.fixture
def shell(legacy, monkeypatch):
    shown: list[tuple] = []
    for name in ("showinfo", "showwarning", "showerror"):
        monkeypatch.setattr(
            compat.MessageBoxBridge,
            name,
            staticmethod(lambda *args, _n=name, **_k: shown.append((_n, args))),
        )
    base = create_qt_app_class(legacy)

    class IsolatedShell(base):
        def load_config(self):
            return None

        def save_config(self):
            return None

    app = IsolatedShell()
    app._refresh_order_signer = lambda text=None: ("", {"position": "", "rank": "", "name": ""})
    app.shown = shown
    yield app
    app.main_window.close()
    app.main_window.deleteLater()
    QApplication.processEvents()


@pytest.fixture
def unlock(shell, monkeypatch):
    """Відмикає вікно наказів тимчасовим паролем."""
    from nodeautomationtoolkit.generator_qt.orders_window import ORDERS_PASSWORD

    monkeypatch.setattr(type(shell), "ask_orders_password", lambda self: ORDERS_PASSWORD)
    return ORDERS_PASSWORD


def _edits_with_text(container, text):
    return [edit for edit in container.findChildren(QLineEdit) if edit.text() == text]


def _plan_file(tmp_path) -> Path:
    from openpyxl import Workbook

    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["ПЛАН переміщення військовослужбовців"])
    sheet.append(PLAN_HEADER)
    for row in PLAN_ROWS:
        sheet.append(row)
    path = tmp_path / "План переміщення.xlsx"
    workbook.save(path)
    return path


def test_orders_are_not_a_tab_and_need_a_password(shell, monkeypatch):
    assert shell._tabs.count() == 3
    assert all("наказ" not in shell._tabs.tabText(i).casefold() for i in range(3))

    monkeypatch.setattr(type(shell), "ask_orders_password", lambda self: "1234")
    assert shell.open_orders_window() is None
    assert getattr(shell, "_orders_window", None) is None


def test_the_right_password_opens_the_window(shell, unlock):
    window = shell.open_orders_window()
    assert window is not None and window.isVisible()
    # Другого разу пароль уже не питають.
    shell.ask_orders_password = lambda: ""
    assert shell.open_orders_window() is window


def test_order_fields_are_in_the_window(shell, unlock):
    page = shell.open_orders_window()
    shell.new_order_template_path.set("C:/Зразки/Шаблон наказу.docx")
    shell.new_order_executor.set("Виконавець наказу")
    shell.new_order_out_folder.set("C:/Накази")
    shell.new_order_archive_folder.set("C:/Архів наказів")

    assert _edits_with_text(page, "C:/Зразки/Шаблон наказу.docx")
    assert _edits_with_text(page, "Виконавець наказу")
    assert _edits_with_text(page, "C:/Накази")
    assert _edits_with_text(page, "C:/Архів наказів")


def test_orders_folder_can_be_picked(shell, unlock, tmp_path, monkeypatch):
    shell.open_orders_window()
    monkeypatch.setattr(
        compat.FileDialogBridge, "askdirectory", staticmethod(lambda **_k: str(tmp_path))
    )
    shell.select_order_archive_folder()
    assert shell.new_order_archive_folder.get() == str(tmp_path)
    # Тека індексу пропонується сама, щоб її не шукали вручну.
    assert shell.new_order_index_folder.get().endswith("order_index")


def test_dropping_a_plan_into_the_window(shell, unlock, tmp_path):
    shell.open_orders_window()
    plan = _plan_file(tmp_path)
    shell.accept_order_source_path(str(plan))

    assert shell.new_order_plan_path.get() == str(plan)
    assert "план" in shell.new_order_source_summary.get().casefold()


def test_four_steps_make_a_document(shell, unlock, tmp_path):
    shell.open_orders_window()
    shell.new_order_plan_path.set(str(_plan_file(tmp_path)))
    shell.new_order_points.set("пункту 45")
    shell.new_order_number.set("525")
    shell.new_order_date.set("17.09.2026")
    shell.new_order_section.set("§ 1")
    shell.new_order_out_folder.set(str(tmp_path / "готове"))
    shell.new_order_signer_position.set("Командувач Сухопутних військ Збройних Сил України")
    shell.new_order_signer_rank.set("генерал-лейтенант")
    shell.new_order_signer_name.set("І. ПЕТРЕНКО")

    shell.run_order_compose()
    assert shell.order_items_table.topLevelItemCount() == 2
    assert shell.order_items_table.topLevelItem(0).text(2) == "БОНДАРЕНКО Сергій Петрович"

    shell.run_order_check()
    assert shell._order_check is not None and shell._order_check.ready

    shell.run_order_assemble()
    made = list((tmp_path / "готове").glob("*.docx"))
    assert [path.name for path in made] == ["Наказ № 525 від 17.09.2026.docx"]


def test_assembly_without_a_draft_says_so(shell, unlock, tmp_path):
    shell.open_orders_window()
    shell.new_order_out_folder.set(str(tmp_path))
    shell.run_order_assemble()
    assert not list(Path(tmp_path).glob("*.docx"))
