from __future__ import annotations

import inspect
from collections import defaultdict, deque
from collections.abc import Callable
from dataclasses import dataclass, field
from typing import Any

from .models import GraphModel
from .registry import NodeRegistry


@dataclass(slots=True)
class ExecutionResult:
    values: dict[str, dict[str, Any]] = field(default_factory=dict)
    order: list[str] = field(default_factory=list)


@dataclass(slots=True)
class PreviewResult:
    values: dict[str, dict[str, Any]] = field(default_factory=dict)
    errors: dict[str, str] = field(default_factory=dict)
    order: list[str] = field(default_factory=list)


class GraphExecutor:
    def __init__(self, registry: NodeRegistry) -> None:
        self.registry = registry

    def execute(
        self,
        graph: GraphModel,
        on_node_started: Callable[[str], None] | None = None,
        on_node_finished: Callable[[str, dict[str, Any]], None] | None = None,
    ) -> ExecutionResult:
        if any(connection.kind == "execution" for connection in graph.connections):
            return self._execute_control_flow(
                graph,
                on_node_started=on_node_started,
                on_node_finished=on_node_finished,
            )
        order = self._topological_order(graph)
        results: dict[str, dict[str, Any]] = {}

        incoming = defaultdict(list)
        for connection in graph.connections:
            if connection.kind == "data":
                incoming[connection.target_node].append(connection)

        for node_id in order:
            model = graph.node_by_id(node_id)
            definition = self.registry.get(model.type_id)
            if on_node_started:
                on_node_started(node_id)

            kwargs = dict(model.parameters)
            for connection in incoming[node_id]:
                kwargs[connection.target_port] = results[connection.source_node][
                    connection.source_port
                ]

            for port in definition.inputs:
                if port.name not in kwargs and not port.required:
                    kwargs[port.name] = port.default
                if port.required and port.name not in kwargs:
                    raise ValueError(
                        f"Нода '{definition.name}': відсутній обов'язковий вхід '{port.name}'"
                    )

            value = definition.function(**kwargs)
            if inspect.isawaitable(value):
                raise RuntimeError("Async-ноди будуть додані в наступній версії")

            outputs = self._map_outputs(definition.outputs, value)
            results[node_id] = outputs
            if on_node_finished:
                on_node_finished(node_id, outputs)

        return ExecutionResult(values=results, order=order)

    def execute_target(
        self,
        graph: GraphModel,
        node_id: str,
        on_node_started: Callable[[str], None] | None = None,
        on_node_finished: Callable[[str, dict[str, Any]], None] | None = None,
    ) -> ExecutionResult:
        graph.node_by_id(node_id)
        results: dict[str, dict[str, Any]] = {}
        order: list[str] = []
        incoming = defaultdict(list)
        for connection in graph.connections:
            if connection.kind == "data":
                incoming[connection.target_node].append(connection)
        self._execute_one_node(
            graph,
            node_id,
            incoming,
            results,
            order,
            on_node_started,
            on_node_finished,
            allow_flow_dependency=True,
        )
        return ExecutionResult(values=results, order=order)

    def preview(
        self,
        graph: GraphModel,
        *,
        trigger_node_id: str | None = None,
        initial_values: dict[str, dict[str, Any]] | None = None,
        on_node_started: Callable[[str], None] | None = None,
        on_node_finished: Callable[[str, dict[str, Any]], None] | None = None,
        on_node_failed: Callable[[str, str], None] | None = None,
    ) -> PreviewResult:
        results = dict(initial_values or {})
        errors: dict[str, str] = {}
        executed_order: list[str] = []
        incoming = defaultdict(list)
        for connection in graph.connections:
            if connection.kind == "data":
                incoming[connection.target_node].append(connection)

        for node_id in self._topological_order(graph):
            if node_id in results:
                continue
            model = graph.node_by_id(node_id)
            definition = self.registry.get(model.type_id)
            if definition.execution_inputs or definition.preview_policy == "never":
                continue
            if definition.preview_policy == "manual" and node_id != trigger_node_id:
                has_saved_selection = any(
                    key.startswith("selected_") and bool(value)
                    for key, value in model.parameters.items()
                )
                if not has_saved_selection:
                    continue

            kwargs = dict(model.parameters)
            unresolved = False
            for connection in incoming[node_id]:
                source_values = results.get(connection.source_node)
                if source_values is None or connection.source_port not in source_values:
                    unresolved = True
                    break
                kwargs[connection.target_port] = source_values[connection.source_port]
            if unresolved:
                continue
            if any(port.required and port.name not in kwargs for port in definition.inputs):
                continue
            for port in definition.inputs:
                if port.name not in kwargs and not port.required:
                    kwargs[port.name] = port.default

            if on_node_started:
                on_node_started(node_id)
            try:
                value = definition.function(**kwargs)
                if inspect.isawaitable(value):
                    raise RuntimeError("Async-ноди не підтримують live-прев'ю")
                outputs = self._map_outputs(definition.outputs, value)
                results[node_id] = outputs
                executed_order.append(node_id)
                if on_node_finished:
                    on_node_finished(node_id, outputs)
            except Exception as error:  # noqa: BLE001 - live preview boundary
                message = str(error)
                errors[node_id] = message
                if on_node_failed:
                    on_node_failed(node_id, message)

        return PreviewResult(values=results, errors=errors, order=executed_order)

    def _execute_control_flow(
        self,
        graph: GraphModel,
        *,
        on_node_started: Callable[[str], None] | None,
        on_node_finished: Callable[[str, dict[str, Any]], None] | None,
    ) -> ExecutionResult:
        results: dict[str, dict[str, Any]] = {}
        order: list[str] = []
        data_incoming = defaultdict(list)
        flow_outgoing = defaultdict(list)
        flow_incoming_count = dict.fromkeys((node.id for node in graph.nodes), 0)

        for connection in graph.connections:
            if connection.kind == "execution":
                flow_outgoing[connection.source_node].append(connection)
                flow_incoming_count[connection.target_node] += 1
            else:
                data_incoming[connection.target_node].append(connection)

        start_nodes = []
        for model in graph.nodes:
            definition = self.registry.get(model.type_id)
            if definition.execution_outputs and flow_incoming_count[model.id] == 0:
                start_nodes.append(model.id)
        if not start_nodes:
            raise ValueError("Execution-граф не має початкової ноди")

        queue = deque(start_nodes)
        executed_flow_nodes: set[str] = set()
        while queue:
            node_id = queue.popleft()
            if node_id in executed_flow_nodes:
                continue
            executed_flow_nodes.add(node_id)
            raw_value = self._execute_one_node(
                graph,
                node_id,
                data_incoming,
                results,
                order,
                on_node_started,
                on_node_finished,
                allow_flow_dependency=False,
            )
            definition = self.registry.get(graph.node_by_id(node_id).type_id)
            if definition.execution_router == "boolean":
                selected_ports = ["true" if bool(raw_value) else "false"]
            else:
                selected_ports = [port.name for port in definition.execution_outputs]
            for port_name in selected_ports:
                for connection in flow_outgoing[node_id]:
                    if connection.source_port == port_name:
                        queue.append(connection.target_node)

        return ExecutionResult(values=results, order=order)

    def _execute_one_node(
        self,
        graph: GraphModel,
        node_id: str,
        data_incoming: dict,
        results: dict[str, dict[str, Any]],
        order: list[str],
        on_node_started: Callable[[str], None] | None,
        on_node_finished: Callable[[str, dict[str, Any]], None] | None,
        *,
        allow_flow_dependency: bool,
    ) -> Any:
        if node_id in results:
            return results[node_id].get("result")
        model = graph.node_by_id(node_id)
        definition = self.registry.get(model.type_id)
        kwargs = dict(model.parameters)
        for connection in data_incoming[node_id]:
            if connection.source_node not in results:
                source_model = graph.node_by_id(connection.source_node)
                source_definition = self.registry.get(source_model.type_id)
                if source_definition.execution_inputs and not allow_flow_dependency:
                    raise ValueError(
                        f"Дані з '{source_definition.name}' ще не були виконані потоком"
                    )
                self._execute_one_node(
                    graph,
                    connection.source_node,
                    data_incoming,
                    results,
                    order,
                    on_node_started,
                    on_node_finished,
                    allow_flow_dependency=True,
                )
            kwargs[connection.target_port] = results[connection.source_node][
                connection.source_port
            ]
        for port in definition.inputs:
            if port.name not in kwargs and not port.required:
                kwargs[port.name] = port.default
            if port.required and port.name not in kwargs:
                raise ValueError(
                    f"Нода '{definition.name}': відсутній обов'язковий вхід '{port.name}'"
                )
        if on_node_started:
            on_node_started(node_id)
        value = definition.function(**kwargs)
        if inspect.isawaitable(value):
            raise RuntimeError("Async-ноди будуть додані в наступній версії")
        outputs = self._map_outputs(definition.outputs, value)
        results[node_id] = outputs
        order.append(node_id)
        if on_node_finished:
            on_node_finished(node_id, outputs)
        return value

    @staticmethod
    def _map_outputs(output_ports: list, value: Any) -> dict[str, Any]:
        if len(output_ports) == 1:
            return {output_ports[0].name: value}
        if isinstance(value, dict):
            return {port.name: value[port.name] for port in output_ports}
        if isinstance(value, (tuple, list)) and len(value) == len(output_ports):
            return {port.name: item for port, item in zip(output_ports, value, strict=True)}
        return {port.name: getattr(value, port.name) for port in output_ports}

    @staticmethod
    def _topological_order(graph: GraphModel) -> list[str]:
        node_ids = {node.id for node in graph.nodes}
        indegree = dict.fromkeys(node_ids, 0)
        outgoing: dict[str, list[str]] = defaultdict(list)

        for connection in graph.connections:
            if connection.kind != "data":
                continue
            if connection.source_node not in node_ids or connection.target_node not in node_ids:
                raise ValueError("Зв'язок посилається на відсутню ноду")
            indegree[connection.target_node] += 1
            outgoing[connection.source_node].append(connection.target_node)

        ready = deque(node_id for node_id, degree in indegree.items() if degree == 0)
        order: list[str] = []
        while ready:
            node_id = ready.popleft()
            order.append(node_id)
            for target_id in outgoing[node_id]:
                indegree[target_id] -= 1
                if indegree[target_id] == 0:
                    ready.append(target_id)

        if len(order) != len(node_ids):
            raise ValueError("Сценарій містить цикл")
        return order
