from nodeautomationtoolkit.core.models import ConnectionModel, GraphModel, NodeModel


def test_graph_json_roundtrip():
    graph = GraphModel(
        name="Тест",
        nodes=[NodeModel(id="one", type_id="example", x=10, y=20, parameters={"x": "ї"})],
        connections=[ConnectionModel("one", "result", "two", "value")],
    )

    restored = GraphModel.from_dict(graph.to_dict())

    assert restored.to_dict() == graph.to_dict()

