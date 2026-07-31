from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any
from uuid import uuid4


@dataclass(slots=True)
class NodeModel:
    type_id: str
    id: str = field(default_factory=lambda: str(uuid4()))
    x: float = 0.0
    y: float = 0.0
    parameters: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        return {
            "id": self.id,
            "type_id": self.type_id,
            "x": self.x,
            "y": self.y,
            "parameters": self.parameters,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> NodeModel:
        return cls(
            id=str(data["id"]),
            type_id=str(data["type_id"]),
            x=float(data.get("x", 0.0)),
            y=float(data.get("y", 0.0)),
            parameters=dict(data.get("parameters", {})),
        )


@dataclass(frozen=True, slots=True)
class ConnectionModel:
    source_node: str
    source_port: str
    target_node: str
    target_port: str

    def to_dict(self) -> dict[str, str]:
        return {
            "source_node": self.source_node,
            "source_port": self.source_port,
            "target_node": self.target_node,
            "target_port": self.target_port,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> ConnectionModel:
        return cls(
            source_node=str(data["source_node"]),
            source_port=str(data["source_port"]),
            target_node=str(data["target_node"]),
            target_port=str(data["target_port"]),
        )


@dataclass(slots=True)
class GraphModel:
    name: str = "Новий сценарій"
    version: int = 1
    nodes: list[NodeModel] = field(default_factory=list)
    connections: list[ConnectionModel] = field(default_factory=list)

    def node_by_id(self, node_id: str) -> NodeModel:
        return next(node for node in self.nodes if node.id == node_id)

    def to_dict(self) -> dict[str, Any]:
        return {
            "format": "nodeautomationtoolkit",
            "version": self.version,
            "name": self.name,
            "nodes": [node.to_dict() for node in self.nodes],
            "connections": [connection.to_dict() for connection in self.connections],
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> GraphModel:
        if data.get("format") != "nodeautomationtoolkit":
            raise ValueError("Це не сценарій Node Automation Toolkit")
        return cls(
            name=str(data.get("name", "Сценарій")),
            version=int(data.get("version", 1)),
            nodes=[NodeModel.from_dict(item) for item in data.get("nodes", [])],
            connections=[
                ConnectionModel.from_dict(item) for item in data.get("connections", [])
            ],
        )
