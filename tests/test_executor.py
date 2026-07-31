import pytest

from nodeautomationtoolkit.core.definition import node
from nodeautomationtoolkit.core.executor import GraphExecutor
from nodeautomationtoolkit.core.models import ConnectionModel, GraphModel, NodeModel
from nodeautomationtoolkit.core.registry import NodeRegistry


def build_registry() -> NodeRegistry:
    @node(type_id="test.value")
    def value(number: int = 1) -> int:
        return number

    @node(type_id="test.add")
    def add(first: int, second: int) -> int:
        return first + second

    registry = NodeRegistry()
    registry.register(value.__nat_node_definition__)
    registry.register(add.__nat_node_definition__)
    return registry


def test_executes_connected_graph():
    registry = build_registry()
    first = NodeModel(id="first", type_id="test.value", parameters={"number": 4})
    second = NodeModel(id="second", type_id="test.value", parameters={"number": 7})
    addition = NodeModel(id="addition", type_id="test.add")
    graph = GraphModel(
        nodes=[first, second, addition],
        connections=[
            ConnectionModel("first", "result", "addition", "first"),
            ConnectionModel("second", "result", "addition", "second"),
        ],
    )

    result = GraphExecutor(registry).execute(graph)

    assert result.values["addition"]["result"] == 11
    assert result.order[-1] == "addition"


def test_rejects_cycle():
    registry = build_registry()
    first = NodeModel(id="first", type_id="test.add", parameters={"second": 1})
    second = NodeModel(id="second", type_id="test.add", parameters={"second": 1})
    graph = GraphModel(
        nodes=[first, second],
        connections=[
            ConnectionModel("first", "result", "second", "first"),
            ConnectionModel("second", "result", "first", "first"),
        ],
    )

    with pytest.raises(ValueError, match="цикл"):
        GraphExecutor(registry).execute(graph)


def test_execution_graph_follows_selected_branch():
    node_registry = NodeRegistry()
    node_registry.reload()
    start = NodeModel(id="start", type_id="builtin.flow.start")
    condition = NodeModel(
        id="condition",
        type_id="builtin.text.value",
        parameters={"value": "yes"},
    )
    branch = NodeModel(id="branch", type_id="builtin.flow.branch")
    true_sequence = NodeModel(id="true", type_id="builtin.flow.sequence")
    false_sequence = NodeModel(id="false", type_id="builtin.flow.sequence")
    graph = GraphModel(
        nodes=[start, condition, branch, true_sequence, false_sequence],
        connections=[
            ConnectionModel("start", "then", "branch", "exec", kind="execution"),
            ConnectionModel("condition", "result", "branch", "condition"),
            ConnectionModel("branch", "true", "true", "exec", kind="execution"),
            ConnectionModel("branch", "false", "false", "exec", kind="execution"),
        ],
    )

    result = GraphExecutor(node_registry).execute(graph)

    assert "true" in result.order
    assert "false" not in result.order


def test_execute_target_runs_only_required_dependencies():
    registry = build_registry()
    first = NodeModel(id="first", type_id="test.value", parameters={"number": 4})
    unused = NodeModel(id="unused", type_id="test.value", parameters={"number": 99})
    addition = NodeModel(
        id="addition",
        type_id="test.add",
        parameters={"second": 3},
    )
    graph = GraphModel(
        nodes=[first, unused, addition],
        connections=[ConnectionModel("first", "result", "addition", "first")],
    )

    result = GraphExecutor(registry).execute_target(graph, "addition")

    assert result.order == ["first", "addition"]
    assert result.values["addition"]["result"] == 7


def test_preview_skips_manual_node_until_triggered():
    @node(type_id="test.manual", preview_policy="manual")
    def manual(value: str = "ready") -> str:
        return value

    registry = NodeRegistry()
    registry.register(manual.__nat_node_definition__)
    graph = GraphModel(nodes=[NodeModel(id="manual", type_id="test.manual")])

    skipped = GraphExecutor(registry).preview(graph)
    executed = GraphExecutor(registry).preview(graph, trigger_node_id="manual")

    assert skipped.order == []
    assert executed.values["manual"]["result"] == "ready"
