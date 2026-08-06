from nodeautomationtoolkit.builtin_nodes import windows_dialogs
from nodeautomationtoolkit.core.registry import NodeRegistry


def test_windows_dialog_nodes_are_available_to_graph_planner():
    registry = NodeRegistry()
    registry.reload()

    assert registry.get("builtin.windows.open_files").outputs[0].data_type in ("list", "List")
    assert registry.get("builtin.windows.select_folder").name == "Вибрати папку"
    assert registry.get("builtin.windows.save_file").name == "Зберегти файл як"


def test_open_file_dialog_reuses_selected_path_without_dialog(monkeypatch):
    def fail_if_called():
        raise AssertionError("діалог не повинен відкриватися")

    monkeypatch.setattr(windows_dialogs, "_qt_widgets", fail_if_called)

    result = windows_dialogs.open_file_dialog(selected_path="C:/orders/order.docx")

    assert result == "C:/orders/order.docx"
