from __future__ import annotations

from .models import ConnectionModel, GraphModel, NodeModel


def build_word_smoke_graph() -> GraphModel:
    return GraphModel(
        name="Word — швидкий тест",
        nodes=[
            NodeModel(id="start", type_id="builtin.flow.start", x=-720, y=-40),
            NodeModel(
                id="pick-word-file",
                type_id="builtin.windows.open_file",
                x=-720,
                y=220,
                parameters={
                    "title": "Виберіть Word-документ",
                    "file_filter": "Документи Word (*.docx)",
                    "initial_folder": "",
                },
            ),
            NodeModel(
                id="read-word-file",
                type_id="builtin.word.read_docx",
                x=-380,
                y=180,
            ),
            NodeModel(
                id="show-word-text",
                type_id="builtin.output.show_result",
                x=20,
                y=-40,
                parameters={"title": "Текст Word-документа"},
            ),
        ],
        connections=[
            ConnectionModel(
                "start",
                "then",
                "show-word-text",
                "exec",
                kind="execution",
            ),
            ConnectionModel(
                "pick-word-file",
                "result",
                "read-word-file",
                "path",
            ),
            ConnectionModel(
                "read-word-file",
                "text",
                "show-word-text",
                "value",
            ),
        ],
    )
