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

