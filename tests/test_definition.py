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


def test_decorator_supports_blueprint_execution_ports():
    @node(
        name="Branch",
        execution_inputs=("exec",),
        execution_outputs=("true", "false"),
        execution_router="boolean",
    )
    def branch(condition: bool) -> bool:
        return condition

    definition = branch.__nat_node_definition__
    assert [port.name for port in definition.execution_inputs] == ["exec"]
    assert [port.name for port in definition.execution_outputs] == ["true", "false"]
    assert definition.execution_router == "boolean"
