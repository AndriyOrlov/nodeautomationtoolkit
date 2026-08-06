import pytest

from nodeautomationtoolkit.core.automation_assistant import (
    AddNodeAction,
    AutomationAssistant,
    AutomationPlan,
    ConnectAction,
    SetParameterAction,
)
from nodeautomationtoolkit.core.models import GraphModel, NodeModel
from nodeautomationtoolkit.core.registry import NodeRegistry


class UnusedClient:
    pass


def registry() -> NodeRegistry:
    result = NodeRegistry()
    result.reload()
    return result


def test_applies_valid_plan_atomically():
    node_registry = registry()
    assistant = AutomationAssistant(UnusedClient(), node_registry)
    plan = AutomationPlan(
        title="Назва файла",
        actions=[
            AddNodeAction(
                action="add_node",
                alias="source",
                type_id="builtin.files.file_path",
                parameters={"path": "C:/test.docx"},
            ),
            AddNodeAction(
                action="add_node",
                alias="name",
                type_id="builtin.files.filename",
                x=280,
            ),
            ConnectAction(
                action="connect",
                source_alias="source",
                source_port="result",
                target_alias="name",
                target_port="path",
            ),
        ],
    )
    original = GraphModel()

    updated = assistant.apply_plan(original, plan)

    assert original.nodes == []
    assert len(updated.nodes) == 2
    assert len(updated.connections) == 1


def test_rejects_unknown_node_without_changing_graph():
    assistant = AutomationAssistant(UnusedClient(), registry())
    plan = AutomationPlan(
        title="Помилка",
        actions=[
            AddNodeAction(
                action="add_node",
                alias="missing",
                type_id="does.not.exist",
            )
        ],
    )
    graph = GraphModel()

    with pytest.raises(KeyError, match="Нода не встановлена"):
        assistant.apply_plan(graph, plan)

    assert graph.nodes == []


def test_can_edit_existing_node_by_safe_alias():
    assistant = AutomationAssistant(UnusedClient(), registry())
    graph = GraphModel(
        nodes=[
            NodeModel(
                id="real-secret-id",
                type_id="builtin.text.value",
                parameters={"value": "C:/secret/order.docx"},
            )
        ]
    )
    plan = AutomationPlan(
        title="Змінити текст",
        actions=[
            SetParameterAction(
                action="set_parameter",
                alias="node_1",
                parameter="value",
                value="Готово",
            )
        ],
    )

    updated = assistant.apply_plan(graph, plan)

    assert updated.nodes[0].parameters["value"] == "Готово"
    summary = assistant.safe_graph_summary(graph)
    assert summary["nodes"][0]["alias"] == "node_1"
    assert summary["nodes"][0]["type_id"] == "builtin.text.value"
