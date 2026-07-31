from __future__ import annotations

from typing import Any

from PySide6.QtCore import Signal
from PySide6.QtWidgets import (
    QCheckBox,
    QDoubleSpinBox,
    QFormLayout,
    QLabel,
    QLineEdit,
    QScrollArea,
    QSpinBox,
    QWidget,
)

from .graph_view import NodeItem


class PropertiesPanel(QScrollArea):
    changed = Signal()

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setWidgetResizable(True)
        self._node: NodeItem | None = None
        self._content = QWidget()
        self._layout = QFormLayout(self._content)
        self.setWidget(self._content)
        self.show_node(None)

    def _clear(self) -> None:
        while self._layout.rowCount():
            self._layout.removeRow(0)

    def show_node(self, node: NodeItem | None) -> None:
        self._clear()
        self._node = node
        if node is None:
            self._layout.addRow(QLabel("Виберіть ноду"))
            return
        self._layout.addRow(QLabel(f"<b>{node.definition.name}</b>"))
        if node.definition.description:
            label = QLabel(node.definition.description)
            label.setWordWrap(True)
            self._layout.addRow(label)

        connected_inputs = {
            connection.target_port
            for connection in node.scene().graph.connections
            if connection.target_node == node.model.id
        }
        for port in node.definition.inputs:
            if port.name in connected_inputs:
                self._layout.addRow(port.name, QLabel(f"← з'єднано ({port.data_type})"))
                continue
            value = node.model.parameters.get(port.name, port.default)
            editor = self._editor_for(value, port.data_type)
            self._layout.addRow(f"{port.name} ({port.data_type})", editor)

    def _editor_for(self, value: Any, data_type: str):
        if data_type == "bool" or isinstance(value, bool):
            editor = QCheckBox()
            editor.setChecked(bool(value))
            editor.toggled.connect(self._update_sender(editor.isChecked))
            return editor
        if data_type == "int" or isinstance(value, int):
            editor = QSpinBox()
            editor.setRange(-1_000_000_000, 1_000_000_000)
            editor.setValue(int(value or 0))
            editor.valueChanged.connect(self._update_sender(editor.value))
            return editor
        if data_type == "float" or isinstance(value, float):
            editor = QDoubleSpinBox()
            editor.setRange(-1e12, 1e12)
            editor.setDecimals(6)
            editor.setValue(float(value or 0))
            editor.valueChanged.connect(self._update_sender(editor.value))
            return editor
        editor = QLineEdit("" if value is None else str(value))
        editor.textChanged.connect(self._update_sender(editor.text))
        return editor

    def _update_sender(self, getter):
        def update(*_args) -> None:
            if self._node is None:
                return
            sender = self.sender()
            row_label = self._layout.labelForField(sender)
            if row_label is None:
                return
            raw_name = row_label.text().split(" (", 1)[0]
            self._node.model.parameters[raw_name] = getter()
            self.changed.emit()

        return update

