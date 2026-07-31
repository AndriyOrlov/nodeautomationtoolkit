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


class GraphExecutor:
    def __init__(self, registry: NodeRegistry) -> None:
        self.registry = registry

    def execute(
        self,
        graph: GraphModel,
        on_node_started: Callable[[str], None] | None = None,
        on_node_finished: Callable[[str, dict[str, Any]], None] | None = None,
    ) -> ExecutionResult:
        order = self._topological_order(graph)
        results: dict[str, dict[str, Any]] = {}

        incoming = defaultdict(list)
        for connection in graph.connections:
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
