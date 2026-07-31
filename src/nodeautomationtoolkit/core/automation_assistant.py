from __future__ import annotations

import json
from typing import Annotated, Literal

from pydantic import BaseModel, Field

from .definition import NodeDefinition
from .local_llm import LocalLlmClient
from .models import ConnectionModel, GraphModel, NodeModel
from .registry import NodeRegistry


class AddNodeAction(BaseModel):
    action: Literal["add_node"]
    alias: str = Field(pattern=r"^[a-zA-Z][a-zA-Z0-9_]{0,49}$")
    type_id: str
    x: float = 0.0
    y: float = 0.0
    parameters: dict = Field(default_factory=dict)


class ConnectAction(BaseModel):
    action: Literal["connect"]
    source_alias: str
    source_port: str
    target_alias: str
    target_port: str
    kind: Literal["data", "execution"] = "data"


class SetParameterAction(BaseModel):
    action: Literal["set_parameter"]
    alias: str
    parameter: str
    value: object


AutomationAction = Annotated[
    AddNodeAction | ConnectAction | SetParameterAction,
    Field(discriminator="action"),
]


class AutomationPlan(BaseModel):
    title: str = Field(min_length=1, max_length=100)
    summary: str = Field(default="", max_length=500)
    actions: list[AutomationAction] = Field(min_length=1, max_length=200)


AUTOMATION_SYSTEM_PROMPT = """Ти плануєш локальний граф автоматизації.

Поверни лише JSON за схемою. Використовуй виключно ноди з каталогу користувача та
точні type_id і назви портів. Спочатку додавай усі ноди через add_node, потім
налаштовуй параметри й створюй з'єднання. Кожній ноді дай короткий унікальний alias.
Розкладай граф зліва направо: x збільшується приблизно на 280, y розділяє гілки.
Не додавай мережеві дії, якщо користувач прямо цього не просив. Не вигадуй ноди,
порти або значення. Якщо потрібної ноди немає, створи план лише з наявної частини та
поясни відсутню можливість у summary. Не запускай граф і не видаляй наявні ноди.
Текст і вміст документів у контекст не передаються.
"""


class AutomationAssistant:
    """Plans validated graph edits without receiving document contents."""

    def __init__(self, client: LocalLlmClient, registry: NodeRegistry) -> None:
        self.client = client
        self.registry = registry

    def create_plan(self, request_text: str) -> AutomationPlan:
        if not request_text.strip():
            raise ValueError("Опишіть потрібну автоматизацію")
        catalog = self._catalog()
        user_prompt = (
            f"ЗАПИТ КОРИСТУВАЧА:\n{request_text.strip()}\n\n"
            f"ДОСТУПНІ НОДИ:\n{json.dumps(catalog, ensure_ascii=False, indent=2)}"
        )
        data = self.client.generate_structured(
            system_prompt=AUTOMATION_SYSTEM_PROMPT,
            user_prompt=user_prompt,
            schema_name="automation_plan",
            schema=AutomationPlan.model_json_schema(),
        )
        return AutomationPlan.model_validate(data)

    def apply_plan(self, graph: GraphModel, plan: AutomationPlan) -> GraphModel:
        """Return a validated copy; the caller decides whether to replace the graph."""
        candidate = GraphModel.from_dict(graph.to_dict())
        aliases: dict[str, NodeModel] = {}

        for action in plan.actions:
            if isinstance(action, AddNodeAction):
                if action.alias in aliases:
                    raise ValueError(f"Повторний alias: {action.alias}")
                definition = self.registry.get(action.type_id)
                parameters = self._validated_parameters(definition, action.parameters)
                model = NodeModel(
                    type_id=action.type_id,
                    x=action.x,
                    y=action.y,
                    parameters=parameters,
                )
                candidate.nodes.append(model)
                aliases[action.alias] = model

        for action in plan.actions:
            if isinstance(action, SetParameterAction):
                node = self._alias(aliases, action.alias)
                definition = self.registry.get(node.type_id)
                allowed = {port.name for port in definition.inputs}
                if action.parameter not in allowed:
                    raise ValueError(
                        f"Нода '{definition.name}' не має параметра '{action.parameter}'"
                    )
                node.parameters[action.parameter] = action.value
            elif isinstance(action, ConnectAction):
                self._connect(candidate, aliases, action)

        return candidate

    def preview(self, plan: AutomationPlan) -> list[str]:
        lines = []
        for action in plan.actions:
            if isinstance(action, AddNodeAction):
                definition = self.registry.get(action.type_id)
                lines.append(f"Додати: {definition.name} ({action.alias})")
            elif isinstance(action, SetParameterAction):
                lines.append(f"Налаштувати {action.alias}.{action.parameter}")
            else:
                lines.append(
                    f"З'єднати {action.source_alias}.{action.source_port} → "
                    f"{action.target_alias}.{action.target_port}"
                )
        return lines

    def _catalog(self) -> list[dict]:
        return [
            {
                "type_id": definition.type_id,
                "name": definition.name,
                "category": definition.category,
                "description": definition.description,
                "inputs": [
                    {
                        "name": port.name,
                        "type": port.data_type,
                        "required": port.required,
                        "default": port.default,
                    }
                    for port in definition.inputs
                ],
                "outputs": [
                    {"name": port.name, "type": port.data_type}
                    for port in definition.outputs
                ],
                "execution_inputs": [port.name for port in definition.execution_inputs],
                "execution_outputs": [port.name for port in definition.execution_outputs],
            }
            for definition in self.registry.all()
        ]

    @staticmethod
    def _alias(aliases: dict[str, NodeModel], alias: str) -> NodeModel:
        try:
            return aliases[alias]
        except KeyError as error:
            raise ValueError(f"Невідомий alias ноди: {alias}") from error

    @staticmethod
    def _validated_parameters(definition: NodeDefinition, values: dict) -> dict:
        allowed = {port.name for port in definition.inputs}
        unknown = set(values) - allowed
        if unknown:
            raise ValueError(
                f"Нода '{definition.name}' не має параметрів: {', '.join(sorted(unknown))}"
            )
        result = {
            port.name: port.default
            for port in definition.inputs
            if not port.required
        }
        result.update(values)
        return result

    def _connect(
        self,
        graph: GraphModel,
        aliases: dict[str, NodeModel],
        action: ConnectAction,
    ) -> None:
        source = self._alias(aliases, action.source_alias)
        target = self._alias(aliases, action.target_alias)
        if source.id == target.id:
            raise ValueError("Ноду не можна з'єднати із самою собою")
        source_definition = self.registry.get(source.type_id)
        target_definition = self.registry.get(target.type_id)
        if action.kind == "execution":
            source_ports = {
                port.name: port for port in source_definition.execution_outputs
            }
            target_ports = {
                port.name: port for port in target_definition.execution_inputs
            }
        else:
            source_ports = {port.name: port for port in source_definition.outputs}
            target_ports = {port.name: port for port in target_definition.inputs}
        if action.source_port not in source_ports:
            raise ValueError(
                f"Немає виходу {source_definition.name}.{action.source_port}"
            )
        if action.target_port not in target_ports:
            raise ValueError(
                f"Немає входу {target_definition.name}.{action.target_port}"
            )
        source_type = source_ports[action.source_port].data_type
        target_type = target_ports[action.target_port].data_type
        if action.kind == "data":
            if "Any" not in (source_type, target_type) and source_type != target_type:
                raise ValueError(f"Несумісні типи: {source_type} → {target_type}")
        if any(
            item.target_node == target.id and item.target_port == action.target_port
            for item in graph.connections
        ):
            raise ValueError(
                f"Вхід {target_definition.name}.{action.target_port} вже з'єднаний"
            )
        graph.connections.append(
            ConnectionModel(
                source_node=source.id,
                source_port=action.source_port,
                target_node=target.id,
                target_port=action.target_port,
                kind=action.kind,
            )
        )
