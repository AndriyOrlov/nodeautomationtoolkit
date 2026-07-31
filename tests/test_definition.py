from nodeautomationtoolkit.core.definition import node


def test_decorator_builds_definition_from_type_hints():
    @node(name="Сума", category="Тест")
    def add(first: int, second: int = 2) -> int:
        return first + second

    definition = add.__nat_node_definition__
    assert definition.name == "Сума"
    assert definition.category == "Тест"
    assert [(port.name, port.data_type, port.required) for port in definition.inputs] == [
        ("first", "int", True),
        ("second", "int", False),
    ]
    assert definition.outputs[0].data_type == "int"

