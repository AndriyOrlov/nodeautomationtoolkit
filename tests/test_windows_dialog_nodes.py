from nodeautomationtoolkit.core.registry import NodeRegistry


def test_windows_dialog_nodes_are_available_to_graph_planner():
    registry = NodeRegistry()
    registry.reload()

    assert registry.get("builtin.windows.open_file").name == "Відкрити файл"
    assert registry.get("builtin.windows.open_files").outputs[0].data_type == "List"
    assert registry.get("builtin.windows.select_folder").name == "Вибрати папку"
    assert registry.get("builtin.windows.save_file").name == "Зберегти файл як"

