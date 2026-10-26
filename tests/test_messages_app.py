"""Окрема програма повідомлень: та сама сторінка й логіка, без витягів і примірників.

Word тут не запускається, конфіг користувача не читається.
"""

import importlib.util
import os
import sys
from pathlib import Path

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
pytest.importorskip("PySide6")

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from PySide6.QtWidgets import QApplication, QTabWidget  # noqa: E402

from nodeautomationtoolkit.generator_qt import compat, messages_window  # noqa: E402
from nodeautomationtoolkit.generator_qt.main_window import install_qt_bridge  # noqa: E402

ICONS = SRC_PATH / "nodeautomationtoolkit" / "generator_qt" / "icons"


@pytest.fixture(scope="module")
def qt_app():
    return QApplication.instance() or QApplication([])


@pytest.fixture(scope="module")
def legacy(qt_app):
    spec = importlib.util.spec_from_file_location(
        "generate_extracts_messages_tests", PROJECT_ROOT / "generate_extracts.py"
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    install_qt_bridge(module)
    return module


@pytest.fixture
def shell(legacy, monkeypatch):
    for name in ("showinfo", "showwarning", "showerror"):
        monkeypatch.setattr(
            compat.MessageBoxBridge, name, staticmethod(lambda *args, **kwargs: None)
        )
    base = messages_window.create_messages_app_class(legacy)

    class IsolatedShell(base):
        def load_config(self):
            return None

        def save_config(self):
            return None

    app = IsolatedShell()
    app._refresh_order_signer = lambda text=None: ("", {"position": "", "rank": "", "name": ""})
    yield app
    app.main_window.close()
    app.main_window.deleteLater()


def test_only_the_messages_page_is_built(shell):
    assert messages_window.APP_TITLE in shell.main_window.windowTitle()
    assert shell.btn_generate_messages is not None
    for generator_part in ("orders_tree", "p2_tree", "copy_two_tree", "btn_extracts", "btn_run_p2"):
        assert not hasattr(shell, generator_part), generator_part
    assert not shell.main_window.windowIcon().isNull()


def test_own_config_next_to_the_program(shell, legacy, monkeypatch, tmp_path):
    assert shell.config_file == os.path.join(
        os.path.dirname(os.path.abspath(legacy.__file__)), "config_messages.json"
    )
    # Однофайловий EXE розпаковується в тимчасову теку — конфіг має лежати поруч із EXE.
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "executable", str(tmp_path / "PovidomlenniaKadrovi.exe"))
    assert shell._config_path() == str(tmp_path / "config_messages.json")


def test_dropped_files_go_to_message_fields(shell, tmp_path):
    order = tmp_path / "Наказ № 12 від 01.10.2026.docx"
    content = tmp_path / "зразок зміст.docx"
    table = tmp_path / "список частин.xlsx"
    for path in (order, content, table):
        path.write_bytes(b"")
    shell.accept_dropped_paths([str(order), str(content), str(table)])
    assert shell.doc_path.get() == str(order)
    assert shell.message_content_template_path.get() == str(content)
    assert shell.excel_path.get() == str(table)
    assert order.name in shell.message_source_summary.get()
    assert order.name in shell.status_strip.selected.text()


def test_samples_window_has_table_and_two_message_samples(shell):
    shell.open_samples_window()
    tabs = shell._samples_dialog.findChild(QTabWidget)
    assert tabs.count() == 1
    assert "Повідомлення" in tabs.tabText(0)
    shell._samples_dialog.close()


def test_logo_files_for_window_and_exe():
    from PIL import Image

    assert (ICONS / "messages_logo.svg").is_file()
    assert (ICONS / "messages_logo.png").is_file()
    with Image.open(ICONS / "messages_logo.ico") as icon:
        assert (256, 256) in icon.info["sizes"] and (16, 16) in icon.info["sizes"]


def test_one_signed_exe_build_script():
    script = (PROJECT_ROOT / "build_messages_qt.bat").read_text(encoding="utf-8")
    for needed in (
        "--onefile",
        "--noupx",
        "messages_logo.ico",
        "--version-file",
        "generate_messages_qt.py",
        "sign_exe.ps1",
    ):
        assert needed in script, needed
    # .bat із кирилицею в коді ламається: команди й коментарі — лише ASCII.
    for line in script.splitlines():
        if line.startswith(("rem ", "set ", '"%PY%"', "  --")):
            assert line.isascii(), line


def test_powershell_scripts_have_bom():
    # Windows PowerShell 5.1 читає .ps1 без BOM як ANSI — кирилиця зламалась би.
    for name in ("sign_exe.ps1", "new_signing_cert.ps1", "install_signing_cert.ps1"):
        assert (
            (PROJECT_ROOT / "scripts" / "signing" / name).read_bytes().startswith(b"\xef\xbb\xbf")
        ), name


def test_version_resource_matches_package():
    sys.path.insert(0, str(PROJECT_ROOT / "scripts" / "messages_exe"))
    try:
        from make_version_info import version_info_text
    finally:
        sys.path.pop(0)
    import nodeautomationtoolkit

    text = version_info_text(nodeautomationtoolkit.__version__)
    assert f"'FileVersion', '{nodeautomationtoolkit.__version__}'" in text
    assert "PovidomlenniaKadrovi.exe" in text
