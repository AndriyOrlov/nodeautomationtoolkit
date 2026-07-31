from __future__ import annotations

import inspect
from collections.abc import Callable
from dataclasses import dataclass, field
from enum import StrEnum
from types import UnionType
from typing import Any, Union, get_args, get_origin, get_type_hints


def type_name(annotation: Any) -> str:
    if annotation in (inspect.Signature.empty, Any, None):
        return "Any"
    origin = get_origin(annotation)
    if origin in (list, tuple, set, dict):
        return origin.__name__.title()
    if origin in (Union, UnionType):
        names = [type_name(item) for item in get_args(annotation) if item is not type(None)]
        return names[0] if len(names) == 1 else "Any"
    return getattr(annotation, "__name__", str(annotation).replace("typing.", ""))


class PortKind(StrEnum):
    DATA = "data"
    EXECUTION = "execution"


@dataclass(frozen=True, slots=True)
class PortDefinition:
    name: str
    data_type: str = "Any"
    required: bool = True
    default: Any = None
    kind: PortKind = PortKind.DATA


@dataclass(slots=True)
class NodeDefinition:
    type_id: str
    name: str
    category: str
    description: str
    function: Callable[..., Any]
    inputs: list[PortDefinition] = field(default_factory=list)
    outputs: list[PortDefinition] = field(default_factory=list)
    execution_inputs: list[PortDefinition] = field(default_factory=list)
    execution_outputs: list[PortDefinition] = field(default_factory=list)
    execution_router: str = "all"
    preview_policy: str = "auto"


def _definition_from_function(
    function: Callable[..., Any],
    *,
    name: str | None,
    category: str,
    description: str,
    type_id: str | None,
    outputs: dict[str, str] | None,
    execution_inputs: tuple[str, ...],
    execution_outputs: tuple[str, ...],
    execution_router: str,
    preview_policy: str,
) -> NodeDefinition:
    signature = inspect.signature(function)
    try:
        hints = get_type_hints(function)
    except (NameError, TypeError):
        hints = {}

    input_ports: list[PortDefinition] = []
    for parameter in signature.parameters.values():
        default = None if parameter.default is inspect.Signature.empty else parameter.default
        input_ports.append(
            PortDefinition(
                name=parameter.name,
                data_type=type_name(hints.get(parameter.name, parameter.annotation)),
                required=parameter.default is inspect.Signature.empty,
                default=default,
            )
        )

    if outputs is not None:
        output_ports = [
            PortDefinition(name=output_name, data_type=output_type, required=False)
            for output_name, output_type in outputs.items()
        ]
    else:
        output_ports = [
            PortDefinition(
                name="result",
                data_type=type_name(hints.get("return", signature.return_annotation)),
                required=False,
            )
        ]

    resolved_type_id = type_id or f"{function.__module__}.{function.__name__}"
    return NodeDefinition(
        type_id=resolved_type_id,
        name=name or function.__name__.replace("_", " ").title(),
        category=category,
        description=description or inspect.getdoc(function) or "",
        function=function,
        inputs=input_ports,
        outputs=output_ports,
        execution_inputs=[
            PortDefinition(
                name=port_name,
                data_type="Execution",
                required=False,
                kind=PortKind.EXECUTION,
            )
            for port_name in execution_inputs
        ],
        execution_outputs=[
            PortDefinition(
                name=port_name,
                data_type="Execution",
                required=False,
                kind=PortKind.EXECUTION,
            )
            for port_name in execution_outputs
        ],
        execution_router=execution_router,
        preview_policy=preview_policy,
    )


def node(
    *,
    name: str | None = None,
    category: str = "Інше",
    description: str = "",
    type_id: str | None = None,
    outputs: dict[str, str] | None = None,
    execution_inputs: tuple[str, ...] = (),
    execution_outputs: tuple[str, ...] = (),
    execution_router: str = "all",
    preview_policy: str = "auto",
) -> Callable[[Callable[..., Any]], Callable[..., Any]]:
    """Expose a regular Python function as a visual node."""

    def decorator(function: Callable[..., Any]) -> Callable[..., Any]:
        definition = _definition_from_function(
            function,
            name=name,
            category=category,
            description=description,
            type_id=type_id,
            outputs=outputs,
            execution_inputs=execution_inputs,
            execution_outputs=execution_outputs,
            execution_router=execution_router,
            preview_policy=preview_policy,
        )
        function.__nat_node_definition__ = definition
        return function

    return decorator
