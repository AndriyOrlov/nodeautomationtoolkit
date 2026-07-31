from __future__ import annotations

import hashlib
import re
from typing import Any

from PySide6.QtCore import Signal
from PySide6.QtWidgets import QSplitter, QVBoxLayout, QWidget

from nodeautomationtoolkit.core.definition import NodeDefinition, PortDefinition, PortKind
from nodeautomationtoolkit.core.models import ConnectionModel, GraphModel, NodeModel
from nodeautomationtoolkit.core.registry import NodeRegistry

TYPE_COLORS = {
    "Any": (148, 163, 184),
    "bool": (220, 70, 90),
    "dict": (245, 158, 11),
    "Dictionary": (245, 158, 11),
    "float": (74, 222, 128),
    "int": (34, 197, 94),
    "List": (168, 85, 247),
    "str": (236, 72, 153),
}
EXECUTION_COLOR = (245, 245, 245)


def port_color(port: PortDefinition) -> tuple[int, int, int]:
    if port.kind == PortKind.EXECUTION:
        return EXECUTION_COLOR
    return TYPE_COLORS.get(port.data_type, (45, 212, 191))


def _safe_class_name(type_id: str) -> str:
    stem = re.sub(r"[^a-zA-Z0-9]+", "_", type_id).strip("_") or "Node"
    digest = hashlib.sha1(type_id.encode("utf-8"), usedforsecurity=False).hexdigest()[:8]
    return f"Nat_{stem}_{digest}"


def create_nodegraphqt_class(definition: NodeDefinition):
    """Build a NodeGraphQt class lazily so core tests need no display server."""
    from NodeGraphQt import BaseNode

    def init(self) -> None:
        BaseNode.__init__(self)
        self.set_color(15, 118, 110)
        for port in definition.execution_inputs + definition.inputs:
            self.add_input(port.name, multi_input=False, color=port_color(port))
        for port in definition.execution_outputs + definition.outputs:
            self.add_output(port.name, multi_output=True, color=port_color(port))
        for port in definition.inputs:
            if not port.required:
                _add_parameter_widget(self, port)

    category = re.sub(r"[^a-zA-Z0-9]+", "_", definition.category).strip("_")
    attributes: dict[str, Any] = {
        "__identifier__": f"nodeautomationtoolkit.{category or 'Other'}",
        "NODE_NAME": definition.name,
        "NAT_TYPE_ID": definition.type_id,
        "NAT_DEFINITION": definition,
        "__init__": init,
    }
    return type(_safe_class_name(definition.type_id), (BaseNode,), attributes)


def _add_parameter_widget(node, port: PortDefinition) -> None:
    value = port.default
    if port.data_type == "bool" or isinstance(value, bool):
        node.add_checkbox(port.name, state=bool(value))
    elif port.data_type == "int" or isinstance(value, int):
        node.add_spinbox(
            port.name,
            value=int(value or 0),
            min_value=-1_000_000_000,
            max_value=1_000_000_000,
        )
    elif port.data_type == "float" or isinstance(value, float):
        node.add_spinbox(
            port.name,
            value=float(value or 0),
            min_value=-1_000_000_000,
            max_value=1_000_000_000,
            double=True,
        )
    else:
        node.add_text_input(port.name, text="" if value is None else str(value))


class NodeGraphQtEditor(QWidget):
    graph_changed = Signal()

    def __init__(self, registry: NodeRegistry, parent=None) -> None:
        super().__init__(parent)
        from NodeGraphQt import NodeGraph, NodesPaletteWidget, PropertiesBinWidget
        from NodeGraphQt.constants import ViewerEnum

        self.registry = registry
        self.graph = NodeGraph()
        self.graph.set_acyclic(True)
        self.graph.set_background_color(11, 17, 32)
        self.graph.set_grid_color(30, 41, 59)
        self.graph.set_grid_mode(ViewerEnum.GRID_DISPLAY_DOTS.value)
        self._register_definitions()

        self.palette = NodesPaletteWidget(node_graph=self.graph)
        self.properties = PropertiesBinWidget(node_graph=self.graph)
        splitter = QSplitter()
        splitter.addWidget(self.palette)
        splitter.addWidget(self.graph.widget)
        splitter.addWidget(self.properties)
        splitter.setSizes([260, 900, 300])

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(splitter)

        self.graph.node_created.connect(lambda _node: self.graph_changed.emit())
        self.graph.nodes_deleted.connect(lambda _ids: self.graph_changed.emit())
        self.graph.port_connected.connect(lambda _a, _b: self.graph_changed.emit())
        self.graph.port_disconnected.connect(lambda _a, _b: self.graph_changed.emit())
        self.graph.property_changed.connect(
            lambda _node, _name, _value: self.graph_changed.emit()
        )

    def _register_definitions(self) -> None:
        for definition in self.registry.all():
            self.graph.register_node(
                create_nodegraphqt_class(definition),
                alias=definition.type_id,
            )

    def set_graph_model(self, model: GraphModel) -> None:
        self.graph.clear_session()
        created = {}
        for item in model.nodes:
            node = self.graph.create_node(
                item.type_id,
                pos=(item.x, item.y),
                push_undo=False,
            )
            if node is None:
                raise ValueError(f"Не вдалося створити ноду {item.type_id}")
            for name, value in item.parameters.items():
                if node.has_property(name):
                    node.set_property(name, value, push_undo=False)
            created[item.id] = node
        for connection in model.connections:
            source = created[connection.source_node]
            target = created[connection.target_node]
            source.get_output(connection.source_port).connect_to(
                target.get_input(connection.target_port),
                push_undo=False,
            )
        self.graph.clear_undo_stack()
        self.graph_changed.emit()

    def graph_model(self) -> GraphModel:
        model = GraphModel(name="Blueprint")
        node_ids = {}
        for node in self.graph.all_nodes():
            definition: NodeDefinition | None = getattr(node, "NAT_DEFINITION", None)
            if definition is None:
                continue
            x, y = node.pos()
            parameters = {}
            for port in definition.inputs:
                if node.has_property(port.name):
                    parameters[port.name] = node.get_property(port.name)
            item = NodeModel(
                id=node.id,
                type_id=definition.type_id,
                x=x,
                y=y,
                parameters=parameters,
            )
            model.nodes.append(item)
            node_ids[node.id] = item

        seen = set()
        for node in self.graph.all_nodes():
            definition: NodeDefinition | None = getattr(node, "NAT_DEFINITION", None)
            if definition is None or node.id not in node_ids:
                continue
            execution_names = {port.name for port in definition.execution_outputs}
            for output_name, output in node.outputs().items():
                for target_port in output.connected_ports():
                    key = (node.id, output_name, target_port.node().id, target_port.name())
                    if key in seen:
                        continue
                    seen.add(key)
                    model.connections.append(
                        ConnectionModel(
                            source_node=node.id,
                            source_port=output_name,
                            target_node=target_port.node().id,
                            target_port=target_port.name(),
                            kind=(
                                "execution" if output_name in execution_names else "data"
                            ),
                        )
                    )
        return model
