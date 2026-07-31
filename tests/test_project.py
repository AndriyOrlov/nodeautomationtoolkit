from nodeautomationtoolkit.core.models import ConnectionModel, GraphModel, NodeModel


def test_graph_json_roundtrip():
    graph = GraphModel(
        name="Тест",
        nodes=[NodeModel(id="one", type_id="example", x=10, y=20, parameters={"x": "ї"})],
        connections=[ConnectionModel("one", "result", "two", "value")],
    )

    restored = GraphModel.from_dict(graph.to_dict())

    assert restored.to_dict() == graph.to_dict()


def test_old_connections_default_to_data_kind():
    connection = ConnectionModel.from_dict(
        {
            "source_node": "a",
            "source_port": "result",
            "target_node": "b",
            "target_port": "value",
        }
    )
    assert connection.kind == "data"


def test_save_graph_creates_missing_parent_directory(tmp_path):
    target = tmp_path / "persistent" / "graphs" / "test.nat.json"
    graph = GraphModel(name="Постійний граф")

    from nodeautomationtoolkit.core.project import load_graph, save_graph

    save_graph(graph, target)

    assert load_graph(target).name == "Постійний граф"
