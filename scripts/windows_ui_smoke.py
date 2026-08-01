from __future__ import annotations

import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication  # noqa: E402

from nodeautomationtoolkit.core.registry import NodeRegistry  # noqa: E402
from nodeautomationtoolkit.ui.nodegraphqt_editor import NodeGraphQtEditor  # noqa: E402


def main() -> int:
    application = QApplication.instance() or QApplication([])
    registry = NodeRegistry()
    registry.reload()
    editor = NodeGraphQtEditor(registry)

    text_node = editor.graph.create_node("builtin.text.value")
    text_node.set_selected(True)
    assert text_node in editor.graph.selected_nodes()
    editor.delete_selected_nodes()
    assert editor.graph.get_node_by_id(text_node.id) is None

    file_node = editor.graph.create_node("builtin.windows.open_file")
    assert file_node.get_widget("_nat_pick_file") is not None
    assert file_node.get_widget("_nat_run") is not None
    assert file_node.get_widget("_nat_preview") is not None
    assert file_node.get_widget("_nat_help") is not None
    assert not editor.properties.isVisible()

    words_node = editor.graph.create_node("builtin.text.split_words_dynamic")
    editor._sync_dynamic_outputs(
        words_node.id,
        {"слова": ["Альфа"], "1. Альфа": "Альфа"},
    )
    assert "1. Альфа" in words_node.outputs()

    placeholder_node = editor.graph.create_node("builtin.word_batch.fill_placeholder")
    replacement_widget = placeholder_node.get_widget("replacement_text")
    assert replacement_widget is not None
    assert replacement_widget.get_custom_widget().minimumHeight() >= 80
    application.processEvents()
    editor.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
