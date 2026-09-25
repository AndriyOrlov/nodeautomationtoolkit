"""Головне вікно Qt-оболонки поверх справжньої логіки `generate_extracts.App`.

Word тут не запускається й конфіг користувача не читається: у ньому робочі
шляхи. Усі файли — вигадані, у тимчасовій теці.
"""

import importlib.util
import os
import sys
from pathlib import Path
from types import SimpleNamespace

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
pytest.importorskip("PySide6")

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from PySide6.QtCore import Qt  # noqa: E402
from PySide6.QtGui import QGuiApplication  # noqa: E402
from PySide6.QtWidgets import QApplication, QLineEdit  # noqa: E402

from nodeautomationtoolkit.generator_qt import compat  # noqa: E402
from nodeautomationtoolkit.generator_qt.main_window import (  # noqa: E402
    create_qt_app_class,
    install_qt_bridge,
)
from nodeautomationtoolkit.generator_qt.widgets import LogConsole  # noqa: E402


@pytest.fixture(scope="module")
def qt_app():
    return QApplication.instance() or QApplication([])


@pytest.fixture(scope="module")
def legacy(qt_app):
    spec = importlib.util.spec_from_file_location(
        "generate_extracts_qt_window_tests", PROJECT_ROOT / "generate_extracts.py"
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    install_qt_bridge(module)
    return module


@pytest.fixture
def dialogs(monkeypatch):
    """Модальні вікна в offscreen чекали б натискання вічно — записуємо їх."""
    shown = []
    for name in ("showinfo", "showwarning", "showerror"):
        monkeypatch.setattr(
            compat.MessageBoxBridge, name, staticmethod(lambda *args, _n=name, **_k: shown.append((_n, args)))
        )
    return shown


@pytest.fixture
def shell(legacy, dialogs):
    base = create_qt_app_class(legacy)

    class IsolatedShell(base):
        saved = 0

        def load_config(self):
            return None

        def save_config(self):
            self.saved += 1

    app = IsolatedShell()
    # Підписанта з наказу читає Word — у тестах його не запускаємо.
    app._refresh_order_signer = lambda text=None: ("", {"position": "", "rank": "", "name": ""})
    yield app
    app.main_window.close()
    app.main_window.deleteLater()
    QApplication.processEvents()


def _edits_with_text(container, text):
    return [edit for edit in container.findChildren(QLineEdit) if edit.text() == text]


def _order(tmp_path, name):
    path = tmp_path / name
    path.write_bytes(b"")
    return os.path.abspath(str(path))


def test_logic_surface_exists_and_tab_order_matches_tk(shell):
    for name in (
        "btn_calc", "btn_extracts", "btn_management_extracts", "btn_full_cycle",
        "btn_run_p2", "btn_generate_messages",
        "p2_tree", "copy_two_tree", "orders_tree", "p2_orders_tree",
        "log_text", "p2_log_text", "results_notebook", "result_views", "result_tabs",
        "notebook", "tab_copies_page", "tab_extracts_page", "tab_messages_page",
    ):
        assert hasattr(shell, name), name
    # На цей порядок спирається handle_drag_and_drop (0 — примірники, 1 — витяги).
    assert shell.notebook.index(shell.tab_copies_page) == 0
    assert shell.notebook.index(shell.tab_extracts_page) == 1
    assert shell.notebook.index(shell.tab_messages_page) == 2
    assert "всі витяги" in shell.btn_extracts.text().lower()
    assert "управління" in shell.btn_management_extracts.text().lower()


def test_run_fields_live_on_every_generation_tab(shell):
    shell.executor.set("Виконавець витягів")
    shell.p2_executor.set("Виконавець примірників")
    shell.message_executor.set("Виконавець повідомлень")
    shell.out_folder.set("C:/Витяги")
    shell.p2_out_folder.set("C:/Примірники")
    shell.message_out_folder.set("C:/Повідомлення")

    assert _edits_with_text(shell.tab_extracts_page, "Виконавець витягів")
    assert _edits_with_text(shell.tab_extracts_page, "C:/Витяги")
    assert _edits_with_text(shell.tab_copies_page, "Виконавець примірників")
    assert _edits_with_text(shell.tab_copies_page, "C:/Примірники")
    assert _edits_with_text(shell.tab_messages_page, "Виконавець повідомлень")
    assert _edits_with_text(shell.tab_messages_page, "C:/Повідомлення")

    edit = _edits_with_text(shell.tab_extracts_page, "Виконавець витягів")[0]
    edit.textEdited.emit("Новий виконавець")
    assert shell.executor.get() == "Новий виконавець"


def test_certifiers_are_separate_for_copies_and_extracts(shell):
    shell.certifier_name.set("Засвідчувач Витягів")
    shell.p2_certifier_name.set("Засвідчувач Примірників")
    assert _edits_with_text(shell.tab_extracts_page, "Засвідчувач Витягів")
    assert not _edits_with_text(shell.tab_extracts_page, "Засвідчувач Примірників")
    assert _edits_with_text(shell.tab_copies_page, "Засвідчувач Примірників")


def test_samples_dialog_holds_only_samples(shell, qt_app):
    shell.template_path.set("C:/Зразки/витяг.docx")
    shell.executor.set("унікальний-виконавець")
    shell.open_samples_window()
    dialog = shell._samples_dialog
    assert dialog is not None
    shell.open_samples_window()
    assert shell._samples_dialog is dialog

    assert _edits_with_text(dialog, "C:/Зразки/витяг.docx")
    assert not _edits_with_text(dialog, "унікальний-виконавець")

    saved_before = shell.saved
    dialog.reject()
    qt_app.processEvents()
    assert shell.saved == saved_before + 1
    assert shell._samples_dialog is None


def test_selected_orders_fill_table_badges_and_status(shell, tmp_path):
    first = _order(tmp_path, "Наказ № 445 від 02.09.2026.docx")
    second = _order(tmp_path, "Наказ № 446 від 03.09.2026.docx")
    shell.notebook.select(shell.tab_extracts_page)
    shell._set_orders([first, second])

    widget = shell.orders_tree.widget
    assert widget.topLevelItemCount() == 2
    assert widget.topLevelItem(0).text(3) == "02.09.2026 / № 445"
    assert widget.topLevelItem(0).text(5) == "✓ Готовий"
    assert "Наказів в обробці: 2" in shell.source_summary.get()
    assert shell._tab_badges[1].text() == "2 у черзі"
    assert shell.status_strip.selected.text() == "Вибрано: 2/2"

    widget.topLevelItem(1).setCheckState(0, Qt.CheckState.Unchecked)
    assert shell._selected_order_paths() == ([first], False)
    assert "Наказів в обробці: 1" in shell.source_summary.get()
    assert shell.status_strip.selected.text() == "Вибрано: 1/2"


def test_selecting_order_loads_its_saved_analysis_without_regeneration(shell, legacy, qt_app, tmp_path):
    first = _order(tmp_path, "Наказ № 71 від 15.09.2026.docx")
    second = _order(tmp_path, "Наказ № 72 від 15.09.2026.docx")

    def save_reports(order_path, sender):
        output = Path(order_path).parent / "Extracts_Output"
        output.mkdir(exist_ok=True)
        base = legacy.sanitize_filename(Path(order_path).stem)
        legacy._save_table_to_excel(
            str(output / f"Розрахунок_розсилки_{base}.xlsx"),
            ["Військова частина / Відправник", "Номери пунктів витягу", "Кількість пунктів"],
            [(sender, "Пункт 1.", 1)],
        )
        legacy._save_table_to_excel(
            str(output / f"Контроль_пропущених_пунктів_{base}.xlsx"),
            ["Пункт", "Текст пункту", "Причина"],
            [],
        )
        legacy._save_table_to_excel(
            str(output / f"Контроль_маршрутизації_{base}.xlsx"),
            [
                "Пункт", "Збіги з таблиці", "Застосовані правила",
                "Адресати з пункту", "Адресати з контексту", "Підсумкові адресати",
            ],
            [("Пункт 1.", sender, "адресат знайдено", sender, "—", sender)],
        )

    save_reports(first, "Перша тестова частина")
    save_reports(second, "Друга тестова частина")
    shell._set_orders([first, second])
    assert shell.result_views["calculation"].widget.topLevelItem(0).text(0) == "Перша тестова частина"

    shell.orders_tree.widget.setCurrentItem(shell.orders_tree.widget.topLevelItem(1))
    qt_app.processEvents()

    assert shell.doc_path.get() == second
    assert shell.result_views["calculation"].widget.topLevelItem(0).text(0) == "Друга тестова частина"
    assert "повторний розрахунок не запускався" in shell.log_model.entries[-1].message


def test_drop_routes_orders_by_active_tab(shell, tmp_path):
    order = _order(tmp_path, "Наказ № 7 від 01.09.2026.docx")
    shell.notebook.select(shell.tab_copies_page)
    shell.accept_dropped_paths([order])
    assert shell.p2_manual_order_paths == [order]

    shell.notebook.select(shell.tab_extracts_page)
    shell.accept_dropped_paths([order])
    assert shell.manual_order_paths == [order]


def test_messages_accept_only_one_order(shell, tmp_path):
    first = _order(tmp_path, "Наказ № 20 від 01.09.2026.docx")
    second = _order(tmp_path, "Наказ № 21 від 01.09.2026.docx")
    shell.notebook.select(shell.tab_messages_page)
    shell.accept_dropped_paths([first, second])
    assert shell.manual_order_paths == [first]
    assert shell.doc_path.get() == first
    assert "Один наказ:" in shell.message_source_summary.get()


def test_messages_take_only_first_order_from_dropped_folder(shell, tmp_path):
    folder = tmp_path / "повідомлення"
    folder.mkdir()
    first = _order(folder, "Наказ № 30 від 01.09.2026.docx")
    _order(folder, "Наказ № 31 від 01.09.2026.docx")
    shell.notebook.select(shell.tab_messages_page)
    shell.accept_dropped_paths([str(folder)])
    assert shell.manual_order_paths == [first]


def test_drop_is_ignored_while_a_batch_runs(shell, tmp_path):
    order = _order(tmp_path, "Наказ № 8 від 01.09.2026.docx")
    shell.notebook.select(shell.tab_extracts_page)
    shell._busy_depth = 1
    try:
        shell.accept_dropped_paths([order])
    finally:
        shell._busy_depth = 0
    assert shell.manual_order_paths == []
    assert "проігноровано" in shell.log_model.entries[-1].message


def test_analysis_results_fill_result_tables(shell, legacy):
    shell.show_analysis_results(
        {
            "units_table": SimpleNamespace(rows=[("55 АК А0055", "—", "Пункт 1.", 1)]),
            "management_paragraphs": {
                "Управління — Пункт 9.": {"items": [{"label": "Пункт 9."}]}
            },
            "unmatched_items": [{"label": "Пункт 2.", "text": "текст", "reason": "немає адресата"}],
            "routing_audit": [
                {"label": "Пункт 1.", "matched_entries": "55 АК", "applied_rules": "—", "final_recipients": "55 АК"}
            ],
        }
    )
    calculation = shell.result_views["calculation"]
    assert len(calculation.get_children()) == 2
    assert calculation.widget.topLevelItem(0).text(0) == "55 АК А0055"
    management = calculation.widget.topLevelItem(1)
    assert management.text(0) == legacy.MANAGEMENT_RESULT_LABEL
    assert management.background(0).color().name() == "#052e2b"
    assert shell.result_views["unmatched"].widget.topLevelItem(0).text(2) == "немає адресата"

    shell.show_layout_warnings(["А0055: шапка відірвана від першого пункту."])
    assert shell._results_tabs.currentWidget() is shell.result_tabs["layout"]
    shell.show_layout_warnings([])
    assert shell._results_tabs.currentWidget() is shell.result_tabs["calculation"]


def test_busy_guard_locks_buttons_through_nested_runs(shell, legacy, monkeypatch):
    seen = []

    def fake_extracts(self):
        seen.append(("extracts", self._busy_depth, self.btn_calc.isEnabled()))
        # Логіка сама вмикає свої кнопки наприкінці — блокування пакета це не знімає.
        self.btn_calc.config(state=compat.NORMAL)
        seen.append(("after-logic", self.btn_calc.isEnabled()))

    def fake_full_cycle(self):
        seen.append(("cycle", self._busy_depth, self.btn_full_cycle.isEnabled()))
        self.run_extracts_action()

    monkeypatch.setattr(legacy.App, "run_extracts_action", fake_extracts)
    monkeypatch.setattr(legacy.App, "run_full_cycle", fake_full_cycle)
    shell.run_full_cycle()

    assert seen == [("cycle", 1, False), ("extracts", 2, False), ("after-logic", False)]
    assert shell._busy_depth == 0
    assert shell.btn_calc.isEnabled() and shell.btn_full_cycle.isEnabled()
    assert shell.status_strip.state.text() == "Готово"


def test_extract_scope_and_management_action_pass_through_qt_guard(shell, legacy, monkeypatch):
    seen = []

    def fake_extracts(self, scope="general"):
        seen.append(("extracts", scope, self._busy_depth))

    def fake_management(self):
        seen.append(("management", self._busy_depth))

    monkeypatch.setattr(legacy.App, "run_extracts_action", fake_extracts)
    monkeypatch.setattr(legacy.App, "run_management_extracts_action", fake_management)

    shell.run_extracts_action(scope="all")
    shell.run_management_extracts_action()

    assert seen == [("extracts", "all", 1), ("management", 1)]
    assert shell._busy_depth == 0


def test_log_is_shown_with_levels_and_mirrored_on_messages_tab(shell):
    shell.log("УВАГА: тестовий рядок")
    assert shell.log_model.entries[-1].level == "warning"
    consoles = [
        console
        for page in (shell.tab_extracts_page, shell.tab_messages_page)
        for console in page.findChildren(LogConsole)
    ]
    assert len(consoles) == 2
    assert all("тестовий рядок" in console.toPlainText() for console in consoles)

    shell.log_p2("рядок примірників")
    assert all("рядок примірників" not in entry.message for entry in shell.log_model.entries)


def test_redacted_copy_uses_plain_messages(shell):
    shell.clear_log()
    shell.log("[1/1] Генеруємо витяг для: 199омбр А9998")
    shell.copy_log(shell.log_text, redacted=True)
    copied = QGuiApplication.clipboard().text()
    assert "А9998" not in copied
    assert "[ІНФО]" not in copied and "[ГОТОВО]" not in copied
    assert "Генеруємо витяг для:" in copied


def test_copy_results_keep_raw_values_for_logic(shell, tmp_path):
    path = _order(tmp_path, "2,3_№445 від 02.09.2026.docx")
    for child in shell.p2_tree.get_children():
        shell.p2_tree.delete(child)
    iid = shell.p2_tree.insert("", compat.END, values=(1, "2,3_№445 від 02.09.2026.docx", "445", "02.09.2026", 3, 2, path))
    assert shell.p2_tree.item(iid, "values")[-1] == path
    row = shell.p2_tree.widget.topLevelItem(0)
    assert row.text(2) == "№ 445"
    assert row.text(4) == "3 стор."
    assert shell._p2_results_count.text() == "Результатів: 1"

    shell.select_all_copies()
    assert shell.p2_tree.selection() == (iid,)


def test_copies_source_mode_switch_is_visible_and_honoured(shell, tmp_path):
    separate = tmp_path / "окремо"
    separate.mkdir()
    single = _order(separate, "Наказ № 9 від 01.09.2026.docx")
    folder = tmp_path / "тека"
    folder.mkdir()
    in_folder = [_order(folder, f"Наказ № {n} від 02.09.2026.docx") for n in (10, 11)]

    shell.notebook.select(shell.tab_copies_page)
    shell._set_p2_orders([single])
    assert shell.p2_source_mode.get() == "file"
    assert shell._p2_mode_files.isChecked() and not shell._p2_mode_folder.isChecked()
    assert shell._selected_p2_order_paths() == [single]
    assert "обрані файли" in shell.p2_source_summary.get()

    shell.p2_orders_folder.set(str(folder))
    shell._p2_mode_folder.click()
    assert shell.p2_source_mode.get() == "folder"
    assert shell._p2_mode_folder.isChecked() and not shell._p2_mode_files.isChecked()
    assert sorted(shell._selected_p2_order_paths()) == sorted(in_folder)

    # Назад до окремого наказу — без діалогу, бо список уже є.
    shell._p2_mode_files.click()
    assert shell.p2_source_mode.get() == "file"
    assert shell._selected_p2_order_paths() == [single]


def test_copies_source_mode_survives_restart(shell, legacy, tmp_path):
    order = _order(tmp_path, "Наказ № 12 від 03.09.2026.docx")
    shell.config_file = str(tmp_path / "config.json")
    shell._set_p2_orders([order])
    legacy.App.save_config(shell)

    shell.p2_source_mode.set("folder")
    shell.p2_manual_order_paths = []
    legacy.App.load_config(shell)
    assert shell.p2_source_mode.get() == "file"
    assert shell.p2_manual_order_paths == [order]


def test_old_config_with_single_file_opens_in_single_mode(shell, legacy, tmp_path):
    import json

    order = _order(tmp_path, "Наказ № 13 від 04.09.2026.docx")
    config = tmp_path / "old_config.json"
    config.write_text(json.dumps({"p2_single_file": order}, ensure_ascii=False), encoding="utf-8")
    shell.config_file = str(config)
    shell.p2_source_mode.set("folder")
    shell.p2_manual_order_paths = []
    legacy.App.load_config(shell)
    assert shell.p2_source_mode.get() == "file"
    assert shell._selected_p2_order_paths() == [order]


# ── Інструкція, папки результату, рамка перетягування наказу ────────────────


def _find_buttons(widget, text_part):
    from nodeautomationtoolkit.generator_qt.compat import TkButton

    return [child for child in widget.findChildren(TkButton) if text_part in child.text()]


def test_instruction_button_opens_dialog_with_text(shell, qt_app):
    from nodeautomationtoolkit.generator_qt.instruction_dialog import INSTRUCTION_FILE

    assert INSTRUCTION_FILE.is_file()
    [instruction_button] = _find_buttons(shell.main_window, "Інструкція")
    instruction_button.click()
    qt_app.processEvents()
    dialog = shell._instruction_dialog
    assert dialog is not None and dialog.isVisible()
    text = dialog.viewer.toPlainText()
    assert "Таблиця частин" in text and "Повідомлення" in text
    dialog.close()


def test_result_folder_buttons_on_extracts_and_messages(shell, tmp_path, monkeypatch, dialogs):
    opened = []
    monkeypatch.setattr(os, "startfile", lambda path: opened.append(path), raising=False)
    assert _find_buttons(shell.tab_extracts_page, "Папка результату")
    assert _find_buttons(shell.tab_messages_page, "Папка результату")

    shell.out_folder.set(str(tmp_path))
    shell.open_extracts_output_folder()
    assert opened == [str(tmp_path)]

    order = _order(tmp_path, "Наказ № 40 від 01.09.2026.docx")
    shell.message_out_folder.set("")
    shell.doc_path.set(order)
    shell.open_message_output_folder()  # Messages_Output ще немає — попередження
    assert dialogs and dialogs[-1][0] == "showwarning"

    (tmp_path / "Messages_Output").mkdir()
    shell.open_message_output_folder()
    assert opened[-1] == str(tmp_path / "Messages_Output")


def test_order_dropped_into_zone_is_selected_and_messages_start(shell, tmp_path, monkeypatch):
    from PySide6.QtCore import QMimeData, QPointF, QUrl
    from PySide6.QtGui import QDropEvent

    from nodeautomationtoolkit.generator_qt.widgets import OrderDropZone

    started = []
    monkeypatch.setattr(type(shell), "run_generate_messages", lambda self: started.append(self.doc_path.get()))
    order = _order(tmp_path, "Наказ № 41 від 01.09.2026.docx")
    [zone] = shell.tab_messages_page.findChildren(OrderDropZone)

    mime = QMimeData()
    mime.setUrls([QUrl.fromLocalFile(str(tmp_path / "~$Наказ.docx")), QUrl.fromLocalFile(order)])
    event = QDropEvent(
        QPointF(5, 5), Qt.DropAction.CopyAction, mime, Qt.MouseButton.LeftButton, Qt.KeyboardModifier.NoModifier
    )
    zone.dropEvent(event)

    assert shell.manual_order_paths == [os.path.abspath(order)]
    assert started == [os.path.abspath(order)]


def test_zone_ignores_non_orders_and_busy_state(shell, tmp_path, monkeypatch):
    from PySide6.QtCore import QUrl

    from nodeautomationtoolkit.generator_qt.widgets import dropped_order_paths

    table = tmp_path / "словник.xlsx"
    table.write_bytes(b"")
    assert dropped_order_paths([QUrl.fromLocalFile(str(table))]) == []

    started = []
    monkeypatch.setattr(type(shell), "run_generate_messages", lambda self: started.append(True))
    order = _order(tmp_path, "Наказ № 42 від 01.09.2026.docx")
    shell._busy_depth = 1
    try:
        shell.generate_messages_for_dropped_order(order)
    finally:
        shell._busy_depth = 0
    assert started == [] and shell.manual_order_paths == []


def test_messages_tab_has_internal_management_checkbox(shell, legacy, tmp_path):
    from PySide6.QtWidgets import QCheckBox

    boxes = [
        box for box in shell.tab_messages_page.findChildren(QCheckBox)
        if "внутрішнє переміщення в управлінні" in box.text()
    ]
    assert len(boxes) == 1
    assert shell.message_skip_internal_management.get() is True
    assert boxes[0].isChecked()

    boxes[0].click()
    assert shell.message_skip_internal_management.get() is False

    shell.config_file = str(tmp_path / "config.json")
    legacy.App.save_config(shell)
    shell.message_skip_internal_management.set(True)
    legacy.App.load_config(shell)
    assert shell.message_skip_internal_management.get() is False


def test_compare_window_ignores_blank_paragraphs_by_default(shell, tmp_path, monkeypatch):
    from nodeautomationtoolkit.generator_qt import compare_window

    calls = []
    monkeypatch.setattr(
        compare_window,
        "compare_docx_documents",
        lambda *args, **kwargs: calls.append(kwargs) or (_ for _ in ()).throw(RuntimeError("stop")),
    )
    monkeypatch.setattr(compare_window.QMessageBox, "critical", lambda *args, **kwargs: None)
    reference = _order(tmp_path, "еталон.docx")
    checked = _order(tmp_path, "студент.docx")

    window = shell._open_compare("витяги", checked)
    assert window.ignore_blank_box.isChecked()
    window.reference_edit.setText(reference)
    window.run_comparison()
    window.ignore_blank_box.setChecked(False)
    window.run_comparison()
    assert [call["ignore_blank_paragraphs"] for call in calls] == [True, False]
    window.close()


def test_order_review_button_writes_report_without_touching_the_order(
    shell, legacy, tmp_path, monkeypatch, dialogs
):
    import openpyxl

    order = tmp_path / "Наказ № 5 від 01.09.2026.docx"
    order.write_bytes(b"original")
    text = "\n".join([
        "НАКАЗ",
        "§ 1",
        "Відповідно до пункту 1 Положення ПРИЗНАЧИТИ:",
        "1. Капітана ТЕСТЕНКА Олега Васильовича, командира роти - КОМАНДИРОМ БАТАЛЬЙОНУ.",
        "3. Капітана ПРИКЛАДЕНКА Івана Петровича, командира роти - КОМАНДИРОМ БАТАЛЬЙОНУ.",
        "",
        "Командир військової частини А0001",
        "полковник                    Петро ТЕСТОВИЙ",
    ])
    marked = []

    class FakeWord:
        Visible = True
        DisplayAlerts = 1

    monkeypatch.setattr(shell, "_read_word_text", lambda path, normalize=True: text)
    monkeypatch.setattr(legacy, "default_index_folder", lambda configured="": "")
    monkeypatch.setattr(legacy.win32com.client, "DispatchEx", lambda *_args: FakeWord())
    monkeypatch.setattr(legacy, "force_quit_word", lambda word: None)
    monkeypatch.setattr(
        legacy, "save_marked_copy", lambda word, source, output, findings: marked.append(output) or (1, 0)
    )
    shell.excel_path.set("")
    shell._set_orders([str(order)])

    [review_button] = _find_buttons(shell.tab_extracts_page, "Перевірити наказ")
    review_button.click()

    report = tmp_path / "Перевірка" / "Перевірка_Наказ № 5 від 01.09.2026.xlsx"
    rows = list(openpyxl.load_workbook(report).active.iter_rows(values_only=True))
    assert rows[0] == ("Серйозність", "Що перевірялось", "Де", "Що не так", "Як виправити")
    assert ("помилка", "Нумерація пунктів", "Пункт 3", "після пункту 1 йде пункт 3") == rows[1][:4]
    assert marked == [str(tmp_path / "Перевірка" / "Наказ № 5 від 01.09.2026 — перевірка.docx")]
    assert order.read_bytes() == b"original"
    assert any(name == "showinfo" and "помилок 1" in args[1] for name, args in dialogs)
