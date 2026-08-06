from __future__ import annotations

from collections.abc import Callable
from typing import Any

from PySide6.QtCore import QLineF, QMimeData, QPointF, QRectF, Qt, Signal
from PySide6.QtGui import (
    QBrush,
    QColor,
    QDrag,
    QPainter,
    QPainterPath,
    QPen,
)
from PySide6.QtWidgets import (
    QGraphicsEllipseItem,
    QGraphicsItem,
    QGraphicsPathItem,
    QGraphicsRectItem,
    QGraphicsScene,
    QGraphicsSimpleTextItem,
    QGraphicsView,
    QListWidget,
    QListWidgetItem,
)

from nodeautomationtoolkit.core.definition import (
    NodeDefinition,
    PortDefinition,
    PortKind,
    are_types_compatible,
)
from nodeautomationtoolkit.core.models import ConnectionModel, GraphModel, NodeModel
from nodeautomationtoolkit.core.registry import NodeRegistry

NODE_MIME = "application/x-nodeautomationtoolkit-node"


class NodePalette(QListWidget):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setDragEnabled(True)
        self.setAlternatingRowColors(True)
        self.setStyleSheet(
            "QListWidget { background: #0f172a; color: #f8fafc; border: 1px solid #334155; font-size: 12px; padding: 4px; }"
            "QListWidget::item { padding: 8px 10px; border-bottom: 1px solid #1e293b; border-radius: 4px; margin-bottom: 2px; }"
            "QListWidget::item:hover { background: #1e293b; color: #38bdf8; }"
            "QListWidget::item:selected { background: #0284c7; color: white; font-weight: bold; }"
        )
        self.setWordWrap(True)

    def populate(self, registry: NodeRegistry, query: str = "") -> None:
        self.clear()
        query = query.casefold().strip()
        for definition in registry.all():
            haystack = (
                f"{definition.category} {definition.name} {definition.description}"
            ).casefold()
            if query and query not in haystack:
                continue
            item = QListWidgetItem(f"[{definition.category}]\n{definition.name}")
            item.setData(Qt.ItemDataRole.UserRole, definition.type_id)
            item.setToolTip(f"{definition.name}\n{definition.description}")
            self.addItem(item)

    def startDrag(self, supported_actions: Qt.DropAction) -> None:
        item = self.currentItem()
        if item is None:
            return
        mime = QMimeData()
        mime.setData(NODE_MIME, item.data(Qt.ItemDataRole.UserRole).encode("utf-8"))
        drag = QDrag(self)
        drag.setMimeData(mime)
        drag.exec(Qt.DropAction.CopyAction)


class PortItem(QGraphicsEllipseItem):
    SIZE = 12.0

    def __init__(
        self,
        node_item: NodeItem,
        definition: PortDefinition,
        is_output: bool,
    ) -> None:
        super().__init__(-self.SIZE / 2, -self.SIZE / 2, self.SIZE, self.SIZE, node_item)
        self.node_item = node_item
        self.definition = definition
        self.is_output = is_output
        if definition.kind == PortKind.EXECUTION:
            color = QColor("#f8fafc")
        elif definition.name in ("rules", "corrections", "overrides", "additional_rules"):
            color = QColor("#f472b6")  # Рожевий колір для портів правил/виправлень знизу
        else:
            color = QColor("#5eead4") if is_output else QColor("#93c5fd")
        self.setBrush(QBrush(color))
        self.setPen(QPen(QColor("#111827"), 1.5))
        self.setZValue(3)
        self.setToolTip(f"{definition.name}: {definition.data_type}")

    def scene_center(self) -> QPointF:
        return self.mapToScene(self.boundingRect().center())


class NodeItem(QGraphicsRectItem):
    HEADER = 36.0
    ROW = 26.0

    def __init__(
        self,
        model: NodeModel,
        definition: NodeDefinition,
        moved_callback: Callable[[], None],
    ) -> None:
        from PySide6.QtGui import QFont, QFontMetrics

        all_inputs = definition.execution_inputs + definition.inputs
        all_outputs = definition.execution_outputs + definition.outputs
        rows = max(len(all_inputs), len(all_outputs), 1)

        # Адаптивний розрахунок ширини ноди під назву та порти
        title_font = QFont("Segoe UI", 10, QFont.Weight.Bold)
        title_metrics = QFontMetrics(title_font)
        title_w = title_metrics.horizontalAdvance(definition.name) + 36.0

        port_font = QFont("Segoe UI", 9)
        port_metrics = QFontMetrics(port_font)
        in_max = max([port_metrics.horizontalAdvance(p.name) for p in all_inputs] + [0])
        out_max = max([port_metrics.horizontalAdvance(p.name) for p in all_outputs] + [0])
        ports_w = in_max + out_max + 54.0

        self.WIDTH = max(250.0, title_w, ports_w)

        super().__init__(0, 0, self.WIDTH, self.HEADER + rows * self.ROW + 14)
        self.model = model
        self.definition = definition
        self.moved_callback = moved_callback
        self.inputs: dict[str, PortItem] = {}
        self.outputs: dict[str, PortItem] = {}

        self.setBrush(QBrush(QColor("#1f2937")))
        self.setPen(QPen(QColor("#475569"), 1.4))
        self.setFlags(
            QGraphicsItem.GraphicsItemFlag.ItemIsMovable
            | QGraphicsItem.GraphicsItemFlag.ItemIsSelectable
            | QGraphicsItem.GraphicsItemFlag.ItemSendsGeometryChanges
        )
        self.setPos(model.x, model.y)
        self.setToolTip(definition.description)

        title_background = QGraphicsRectItem(0, 0, self.WIDTH, self.HEADER, self)
        title_background.setBrush(QBrush(QColor("#0f766e")))
        title_background.setPen(QPen(Qt.PenStyle.NoPen))
        title = QGraphicsSimpleTextItem(definition.name, self)
        title.setFont(title_font)
        title.setBrush(QBrush(QColor("#f8fafc")))
        title.setPos(12, 8)

        for index, port in enumerate(all_inputs):
            y = self.HEADER + self.ROW * index + self.ROW / 2 + 5
            port_item = PortItem(self, port, is_output=False)
            port_item.setPos(0, y)
            self.inputs[port.name] = port_item
            label = QGraphicsSimpleTextItem(port.name, self)
            label.setFont(port_font)
            label.setBrush(QBrush(QColor("#dbeafe")))
            label.setPos(10, y - 9)

        for index, port in enumerate(all_outputs):
            y = self.HEADER + self.ROW * index + self.ROW / 2 + 5
            port_item = PortItem(self, port, is_output=True)
            port_item.setPos(self.WIDTH, y)
            self.outputs[port.name] = port_item
            label = QGraphicsSimpleTextItem(port.name, self)
            label.setFont(port_font)
            label.setBrush(QBrush(QColor("#ccfbf1")))
            label_width = label.boundingRect().width()
            label.setPos(self.WIDTH - label_width - 10, y - 9)

    def itemChange(self, change: QGraphicsItem.GraphicsItemChange, value: Any) -> Any:
        if change == QGraphicsItem.GraphicsItemChange.ItemPositionHasChanged:
            position = value
            self.model.x = position.x()
            self.model.y = position.y()
            self.moved_callback()
        return super().itemChange(change, value)


class ConnectionItem(QGraphicsPathItem):
    def __init__(self, model: ConnectionModel, source: PortItem, target: PortItem) -> None:
        super().__init__()
        self.model = model
        self.source = source
        self.target = target
        self.setPen(QPen(QColor("#67e8f9"), 2.2))
        self.setZValue(-1)
        self.update_path()

    def update_path(self) -> None:
        start = self.source.scene_center()
        end = self.target.scene_center()
        distance = max(abs(end.x() - start.x()) * 0.5, 60.0)
        path = QPainterPath(start)
        path.cubicTo(
            QPointF(start.x() + distance, start.y()),
            QPointF(end.x() - distance, end.y()),
            end,
        )
        self.setPath(path)


class GraphScene(QGraphicsScene):
    node_selected = Signal(object)
    graph_changed = Signal()
    message = Signal(str)

    def __init__(self, registry: NodeRegistry, parent=None) -> None:
        super().__init__(parent)
        self.registry = registry
        self.graph = GraphModel()
        self.node_items: dict[str, NodeItem] = {}
        self.connection_items: list[ConnectionItem] = []
        self.pending_port: PortItem | None = None
        self.temporary_line = None
        self.setSceneRect(-3000, -3000, 6000, 6000)
        self.selectionChanged.connect(self._selection_changed)

    def set_graph(self, graph: GraphModel) -> None:
        self.clear()
        self.graph = graph
        self.node_items.clear()
        self.connection_items.clear()
        for model in graph.nodes:
            self._create_item(model)
        for connection in graph.connections:
            self._create_connection_item(connection)
        self.graph_changed.emit()

    def add_node(self, type_id: str, position: QPointF, parameters: dict | None = None) -> NodeItem:
        definition = self.registry.get(type_id)
        model = NodeModel(
            type_id=type_id,
            x=position.x(),
            y=position.y(),
            parameters={
                port.name: port.default
                for port in definition.inputs
                if not port.required
            },
        )
        if parameters:
            model.parameters.update(parameters)
        self.graph.nodes.append(model)
        item = self._create_item(model)
        self.graph_changed.emit()
        return item

    def _create_item(self, model: NodeModel) -> NodeItem:
        definition = self.registry.get(model.type_id)
        item = NodeItem(model, definition, self.update_connections)
        self.addItem(item)
        self.node_items[model.id] = item
        return item

    def _create_connection_item(self, model: ConnectionModel) -> ConnectionItem | None:
        source_node = self.node_items.get(model.source_node)
        target_node = self.node_items.get(model.target_node)
        if source_node is None or target_node is None:
            return None
        source_port = source_node.outputs.get(model.source_port)
        target_port = target_node.inputs.get(model.target_port)
        if source_port is None or target_port is None:
            return None
        item = ConnectionItem(
            model,
            source_port,
            target_port,
        )
        self.addItem(item)
        self.connection_items.append(item)
        return item

    def update_connections(self) -> None:
        for connection in self.connection_items:
            connection.update_path()

    def _selection_changed(self) -> None:
        selected = [item for item in self.selectedItems() if isinstance(item, NodeItem)]
        self.node_selected.emit(selected[0] if len(selected) == 1 else None)

    def mousePressEvent(self, event) -> None:
        item = self.itemAt(event.scenePos(), self.views()[0].transform()) if self.views() else None
        if isinstance(item, PortItem):
            self.pending_port = item
            self.temporary_line = self.addLine(
                QLineF(item.scene_center(), event.scenePos()),
                QPen(QColor("#facc15"), 2, Qt.PenStyle.DashLine),
            )
            event.accept()
            return
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event) -> None:
        if self.pending_port is not None and self.temporary_line is not None:
            self.temporary_line.setLine(QLineF(self.pending_port.scene_center(), event.scenePos()))
            event.accept()
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event) -> None:
        if self.pending_port is not None:
            start = self.pending_port
            candidates = self.items(
                QRectF(event.scenePos().x() - 8, event.scenePos().y() - 8, 16, 16)
            )
            end = next((item for item in candidates if isinstance(item, PortItem)), None)
            if self.temporary_line is not None:
                self.removeItem(self.temporary_line)
            self.temporary_line = None
            self.pending_port = None
            if end is not None and start is not end:
                self.try_connect(start, end)
            event.accept()
            return
        super().mouseReleaseEvent(event)

    def try_connect(self, first: PortItem, second: PortItem) -> None:
        if first.is_output == second.is_output:
            self.message.emit("З'єднайте вихід ноди з входом іншої ноди")
            return
        source, target = (first, second) if first.is_output else (second, first)
        source_type = source.definition.data_type
        target_type = target.definition.data_type
        if not are_types_compatible(source_type, target_type):
            self.message.emit(f"Несумісні типи: {source_type} → {target_type}")
            return
        if source.node_item.model.id == target.node_item.model.id:
            self.message.emit("Ноду не можна з'єднати із самою собою")
            return

        self.remove_connection_to(target.node_item.model.id, target.definition.name)
        model = ConnectionModel(
            source_node=source.node_item.model.id,
            source_port=source.definition.name,
            target_node=target.node_item.model.id,
            target_port=target.definition.name,
            kind=source.definition.kind.value,
        )
        self.graph.connections.append(model)
        self._create_connection_item(model)
        self.graph_changed.emit()

    def remove_connection_to(self, node_id: str, port_name: str) -> None:
        for item in list(self.connection_items):
            if item.model.target_node == node_id and item.model.target_port == port_name:
                self.removeItem(item)
                self.connection_items.remove(item)
                self.graph.connections.remove(item.model)

    def delete_selected(self) -> None:
        selected_nodes = [item for item in self.selectedItems() if isinstance(item, NodeItem)]
        selected_connections = [
            item for item in self.selectedItems() if isinstance(item, ConnectionItem)
        ]
        node_ids = {item.model.id for item in selected_nodes}
        for item in list(self.connection_items):
            if item in selected_connections or {
                item.model.source_node,
                item.model.target_node,
            } & node_ids:
                self.removeItem(item)
                self.connection_items.remove(item)
                self.graph.connections.remove(item.model)
        for item in selected_nodes:
            self.removeItem(item)
            self.graph.nodes.remove(item.model)
            self.node_items.pop(item.model.id, None)
        if selected_nodes or selected_connections:
            self.graph_changed.emit()


class GraphView(QGraphicsView):
    def __init__(self, scene: GraphScene, parent=None) -> None:
        super().__init__(scene, parent)
        self.setAcceptDrops(True)
        self.setRenderHint(QPainter.RenderHint.Antialiasing)
        self.setDragMode(QGraphicsView.DragMode.RubberBandDrag)
        self.setViewportUpdateMode(QGraphicsView.ViewportUpdateMode.BoundingRectViewportUpdate)
        self.setBackgroundBrush(QBrush(QColor("#0b1120")))
        self._panning = False
        self._last_pan_point = None

    @property
    def graph_scene(self) -> GraphScene:
        return self.scene()

    def drawBackground(self, painter: QPainter, rect: QRectF) -> None:
        super().drawBackground(painter, rect)
        grid = 24
        left = int(rect.left()) - (int(rect.left()) % grid)
        top = int(rect.top()) - (int(rect.top()) % grid)
        painter.setPen(QPen(QColor("#172033"), 1))
        x = left
        while x < rect.right():
            painter.drawLine(QLineF(x, rect.top(), x, rect.bottom()))
            x += grid
        y = top
        while y < rect.bottom():
            painter.drawLine(QLineF(rect.left(), y, rect.right(), y))
            y += grid

    def dragEnterEvent(self, event) -> None:
        if event.mimeData().hasFormat(NODE_MIME) or event.mimeData().hasUrls():
            event.acceptProposedAction()
            return
        super().dragEnterEvent(event)

    def dragMoveEvent(self, event) -> None:
        if event.mimeData().hasFormat(NODE_MIME) or event.mimeData().hasUrls():
            event.acceptProposedAction()
            return
        super().dragMoveEvent(event)

    def dropEvent(self, event) -> None:
        position = self.mapToScene(event.position().toPoint())
        if event.mimeData().hasFormat(NODE_MIME):
            type_id = bytes(event.mimeData().data(NODE_MIME)).decode("utf-8")
            self.graph_scene.add_node(type_id, position)
            event.acceptProposedAction()
            return
        if event.mimeData().hasUrls():
            offset = QPointF()
            for url in event.mimeData().urls():
                if url.isLocalFile():
                    self.graph_scene.add_node(
                        "builtin.files.file_path",
                        position + offset,
                        {"path": url.toLocalFile()},
                    )
                    offset += QPointF(28, 28)
            event.acceptProposedAction()
            return
        super().dropEvent(event)

    def wheelEvent(self, event) -> None:
        factor = 1.15 if event.angleDelta().y() > 0 else 1 / 1.15
        self.scale(factor, factor)

    def mousePressEvent(self, event) -> None:
        if event.button() == Qt.MouseButton.MiddleButton:
            self._panning = True
            self._last_pan_point = event.position()
            self.setCursor(Qt.CursorShape.ClosedHandCursor)
            event.accept()
            return
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event) -> None:
        if self._panning and self._last_pan_point is not None:
            delta = event.position() - self._last_pan_point
            self._last_pan_point = event.position()
            self.horizontalScrollBar().setValue(self.horizontalScrollBar().value() - int(delta.x()))
            self.verticalScrollBar().setValue(self.verticalScrollBar().value() - int(delta.y()))
            event.accept()
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event) -> None:
        if event.button() == Qt.MouseButton.MiddleButton:
            self._panning = False
            self._last_pan_point = None
            self.unsetCursor()
            event.accept()
            return
        super().mouseReleaseEvent(event)

    def keyPressEvent(self, event) -> None:
        if event.key() == Qt.Key.Key_Delete:
            self.graph_scene.delete_selected()
            event.accept()
            return
        super().keyPressEvent(event)
