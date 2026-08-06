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


def build_order_senders_graph() -> GraphModel:
    """Готовий сценарій: Аналіз відправників/ВЧ у відкритому наказі та розподіл пунктів."""
    return GraphModel(
        name="Аналіз відправників та пунктів наказу",
        nodes=[
            NodeModel(id="start", type_id="builtin.flow.start", x=-800, y=0),
            NodeModel(
                id="pick-order-file",
                type_id="builtin.windows.open_file",
                x=-800,
                y=200,
                parameters={
                    "title": "Виберіть відкритий наказ (.docx)",
                    "file_filter": "Документи Word (*.docx)",
                    "initial_folder": "",
                },
            ),
            NodeModel(
                id="read-order",
                type_id="builtin.word.read_docx",
                x=-540,
                y=160,
            ),
            NodeModel(
                id="pick-mapping-file",
                type_id="builtin.windows.open_file",
                x=-540,
                y=380,
                parameters={
                    "title": "Виберіть таблицю ВЧ/відправників (.xlsx/.csv)",
                    "file_filter": "Таблиці (*.xlsx *.csv)",
                    "initial_folder": "",
                },
            ),
            NodeModel(
                id="read-mapping",
                type_id="builtin.order.read_recipient_mapping",
                x=-260,
                y=360,
                parameters={
                    "open_name_column": "Відкрите найменування",
                    "cipher_column": "Шифр",
                    "destination_column": "Куди направляється",
                },
            ),
            NodeModel(
                id="map-units",
                type_id="builtin.order.map_military_units",
                x=-20,
                y=160,
            ),
            NodeModel(
                id="groups-ciphers",
                type_id="builtin.order.groups_to_ciphers",
                x=260,
                y=240,
            ),
            NodeModel(
                id="show-report",
                type_id="builtin.output.show_table",
                x=540,
                y=0,
                parameters={"title": "Аналіз відправників та відповідних пунктів"},
            ),
        ],
        connections=[
            ConnectionModel("start", "then", "read-order", "exec", kind="execution"),
            ConnectionModel("read-order", "then", "map-units", "exec", kind="execution"),
            ConnectionModel("map-units", "then", "groups-ciphers", "exec", kind="execution"),
            ConnectionModel("groups-ciphers", "then", "show-report", "exec", kind="execution"),
            ConnectionModel("pick-order-file", "result", "read-order", "path"),
            ConnectionModel("read-order", "text", "map-units", "text"),
            ConnectionModel("pick-mapping-file", "result", "read-mapping", "path"),
            ConnectionModel("read-mapping", "mapping", "map-units", "mapping"),
            ConnectionModel("read-mapping", "mapping", "groups-ciphers", "mapping"),
            ConnectionModel("map-units", "unit_paragraphs", "groups-ciphers", "groups"),
            ConnectionModel("groups-ciphers", "report", "show-report", "table"),
        ],
    )
