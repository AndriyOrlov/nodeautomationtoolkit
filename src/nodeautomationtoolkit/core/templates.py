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


def build_order_blocks_constructor_graph() -> GraphModel:
    """Готовий сценарій: Блочний конструктор наказу (розбірка на блоки, заміна назв і зворотів та збирання закритими)."""
    return GraphModel(
        name="Конструктор блоків наказу (закритий наказ)",
        nodes=[
            NodeModel(id="start", type_id="builtin.flow.start", x=-900, y=0),
            NodeModel(
                id="pick-order-file",
                type_id="builtin.windows.open_file",
                x=-900,
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
                x=-620,
                y=160,
            ),
            NodeModel(
                id="pick-mapping-file",
                type_id="builtin.windows.open_file",
                x=-620,
                y=380,
                parameters={
                    "title": "Виберіть таблицю ВЧ (.xlsx/.csv)",
                    "file_filter": "Таблиці (*.xlsx *.csv)",
                    "initial_folder": "",
                },
            ),
            NodeModel(
                id="read-mapping",
                type_id="builtin.order.read_recipient_mapping",
                x=-340,
                y=360,
            ),
            NodeModel(
                id="parse-blocks",
                type_id="builtin.order.parse_to_blocks",
                x=-340,
                y=160,
            ),
            NodeModel(
                id="transform-blocks",
                type_id="builtin.order.filter_transform_blocks",
                x=-60,
                y=160,
                parameters={"replace_unit_phrases": True},
            ),
            NodeModel(
                id="assemble-order",
                type_id="builtin.order.assemble_from_blocks",
                x=220,
                y=160,
                parameters={
                    "new_header": "НАКАЗ командира військової частини А0000 (по стройовій частині)",
                    "separator": "\n\n",
                },
            ),
            NodeModel(
                id="show-result",
                type_id="builtin.output.show_result",
                x=500,
                y=160,
                parameters={"title": "Сформований закритий наказ про прийняття рішень"},
            ),
        ],
        connections=[
            ConnectionModel("start", "then", "read-order", "exec", kind="execution"),
            ConnectionModel("read-order", "then", "parse-blocks", "exec", kind="execution"),
            ConnectionModel("parse-blocks", "then", "transform-blocks", "exec", kind="execution"),
            ConnectionModel("transform-blocks", "then", "assemble-order", "exec", kind="execution"),
            ConnectionModel("assemble-order", "then", "show-result", "exec", kind="execution"),
            ConnectionModel("pick-order-file", "result", "read-order", "path"),
            ConnectionModel("read-order", "text", "parse-blocks", "text"),
            ConnectionModel("pick-mapping-file", "result", "read-mapping", "path"),
            ConnectionModel("read-mapping", "mapping", "transform-blocks", "mapping"),
            ConnectionModel("parse-blocks", "blocks", "transform-blocks", "blocks"),
            ConnectionModel("transform-blocks", "blocks", "assemble-order", "blocks"),
            ConnectionModel("assemble-order", "text", "show-result", "value"),
        ],
    )


def build_order_extracts_generator_graph() -> GraphModel:
    """Повний автоматизований сценарій: генерація витягів та закритого наказу про прийняття рішень."""
    return GraphModel(
        name="Генерація витягів з наказу та закритих наказів",
        nodes=[
            NodeModel(id="start", type_id="builtin.flow.start", x=-900, y=0),
            NodeModel(
                id="read-mapping",
                type_id="builtin.order.read_recipient_mapping",
                x=-900,
                y=180,
                parameters={"path": "samples/sample_recipient_mapping.xlsx"},
            ),
            NodeModel(
                id="read-order",
                type_id="builtin.word.read_docx",
                x=-600,
                y=0,
                parameters={"path": "samples/sample_open_order.docx"},
            ),
            NodeModel(
                id="map-senders",
                type_id="builtin.order.map_military_units",
                x=-320,
                y=0,
            ),
            NodeModel(
                id="generate-extracts",
                type_id="builtin.order_batch.create_unit_extracts",
                x=-40,
                y=-100,
                parameters={"output_folder": "output/extracts"},
            ),
            NodeModel(
                id="generate-decision",
                type_id="builtin.word.generate_decision_order_docx",
                x=-40,
                y=100,
                parameters={
                    "output_path": "output/decision_order.docx",
                    "new_header": "НАКАЗ командира військової частини А0000 (по стройовій частині)",
                    "replace_unit_phrases": True,
                },
            ),
            NodeModel(
                id="group-ciphers",
                type_id="builtin.order.groups_to_ciphers",
                x=240,
                y=-100,
            ),
            NodeModel(
                id="show-report",
                type_id="builtin.output.show_table",
                x=520,
                y=-100,
            ),
        ],
        connections=[
            ConnectionModel("start", "then", "read-order", "exec", kind="execution"),
            ConnectionModel("read-order", "then", "map-senders", "exec", kind="execution"),
            ConnectionModel("map-senders", "then", "generate-extracts", "exec", kind="execution"),
            ConnectionModel("map-senders", "then", "generate-decision", "exec", kind="execution"),
            ConnectionModel("generate-extracts", "then", "group-ciphers", "exec", kind="execution"),
            ConnectionModel("group-ciphers", "then", "show-report", "exec", kind="execution"),
            ConnectionModel("read-mapping", "mapping", "map-senders", "mapping"),
            ConnectionModel("read-mapping", "mapping", "generate-decision", "mapping"),
            ConnectionModel("read-mapping", "mapping", "group-ciphers", "mapping"),
            ConnectionModel("read-order", "text", "map-senders", "text"),
            ConnectionModel("read-order", "path", "generate-decision", "path"),
            ConnectionModel("map-senders", "unit_paragraphs", "generate-extracts", "unit_paragraphs"),
            ConnectionModel("map-senders", "unit_paragraphs", "group-ciphers", "groups"),
            ConnectionModel("group-ciphers", "report", "show-report", "table"),
        ],
    )
