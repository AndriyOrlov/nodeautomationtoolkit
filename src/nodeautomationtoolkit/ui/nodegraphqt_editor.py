from __future__ import annotations

import hashlib
import inspect
import re
from collections import deque
from typing import Any

from PySide6.QtCore import QEvent, QObject, QPointF, QRunnable, Qt, QThreadPool, QTimer, Signal
from PySide6.QtGui import QAction, QKeySequence, QPixmap
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QDialog,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QScrollArea,
    QSplitter,
    QStyle,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from nodeautomationtoolkit.core.definition import NodeDefinition, PortDefinition, PortKind
from nodeautomationtoolkit.core.executor import GraphExecutor, PreviewResult
from nodeautomationtoolkit.core.models import ConnectionModel, GraphModel, NodeModel
from nodeautomationtoolkit.core.preview import format_live_preview
from nodeautomationtoolkit.core.registry import NodeRegistry

TYPE_COLORS = {
    "Any": (148, 163, 184),
    "bool": (220, 70, 90),
    "dict": (245, 158, 11),
    "Dictionary": (245, 158, 11),
    "DataTable": (14, 165, 233),
    "float": (74, 222, 128),
    "int": (34, 197, 94),
    "List": (168, 85, 247),
    "str": (236, 72, 153),
    "WordDocument": (37, 99, 235),
    "WordDocumentBatch": (124, 58, 237),
    "WordParagraphs": (6, 182, 212),
    "WordSaveResult": (22, 163, 74),
}
EXECUTION_COLOR = (245, 245, 245)
NODE_STATE_COLORS = {
    "idle": (15, 118, 110),
    "running": (180, 83, 9),
    "success": (21, 128, 61),
    "error": (185, 28, 28),
}


class _PreviewWorkerSignals(QObject):
    finished = Signal(int, object)


class _PreviewWorker(QRunnable):
    def __init__(
        self,
        generation: int,
        registry: NodeRegistry,
        graph: GraphModel,
        trigger_node_id: str | None,
        initial_values: dict[str, dict[str, Any]],
    ) -> None:
        super().__init__()
        self.generation = generation
        self.registry = registry
        self.graph = graph
        self.trigger_node_id = trigger_node_id
        self.initial_values = initial_values
        self.signals = _PreviewWorkerSignals()

    def run(self) -> None:
        result = GraphExecutor(self.registry).preview(
            self.graph,
            trigger_node_id=self.trigger_node_id,
            initial_values=self.initial_values,
        )
        self.signals.finished.emit(self.generation, result)


class _FileDropFilter(QObject):
    def __init__(self, editor: NodeGraphQtEditor) -> None:
        super().__init__(editor)
        self.editor = editor

    def eventFilter(self, watched: QObject, event: QEvent) -> bool:
        if event.type() in (QEvent.Type.DragEnter, QEvent.Type.DragMove):
            if event.mimeData().hasUrls():
                event.acceptProposedAction()
                return True
        elif event.type() == QEvent.Type.Drop:
            if event.mimeData().hasUrls():
                event.acceptProposedAction()
                urls = [url.toLocalFile() for url in event.mimeData().urls() if url.isLocalFile()]
                if urls:
                    pos = event.position().toPoint() if hasattr(event, "position") else event.pos()
                    self.editor._handle_dropped_file_paths(urls, pos)
                return True
        return super().eventFilter(watched, event)


def port_color(port: PortDefinition) -> tuple[int, int, int]:
    if port.kind == PortKind.EXECUTION:
        return EXECUTION_COLOR
    return TYPE_COLORS.get(port.data_type, (45, 212, 191))


def _safe_class_name(type_id: str) -> str:
    stem = re.sub(r"[^a-zA-Z0-9]+", "_", type_id).strip("_") or "Node"
    digest = hashlib.sha1(type_id.encode("utf-8"), usedforsecurity=False).hexdigest()[:8]
    return f"Nat_{stem}_{digest}"


def _node_widget_classes():
    from NodeGraphQt.widgets.node_widgets import NodeBaseWidget

    class NodeStatusWidget(NodeBaseWidget):
        def __init__(self, parent=None, name="_nat_status") -> None:
            super().__init__(parent, name, "Стан")
            self._label_widget = QLabel("ОЧІКУЄ")
            self._label_widget.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self._label_widget.setMinimumWidth(220)
            self.set_custom_widget(self._label_widget)
            self.set_state("idle", "ОЧІКУЄ")

        def get_value(self):
            return self._label_widget.text()

        def set_value(self, value):
            self._label_widget.setText(str(value))

        def set_state(self, state: str, text: str) -> None:
            colors = {
                "idle": (51, 65, 85),
                "running": (180, 83, 9),
                "success": (21, 128, 61),
                "error": (185, 28, 28),
            }
            red, green, blue = colors.get(state, colors["idle"])
            self._label_widget.setText(text)
            self._label_widget.setStyleSheet(
                "QLabel {"
                f"background: rgb({red}, {green}, {blue});"
                "border-radius: 4px; color: white; font-weight: 700;"
                "padding: 4px 8px;"
                "}"
            )

    class NodePreviewWidget(NodeBaseWidget):
        def __init__(self, parent=None, name="_nat_preview") -> None:
            super().__init__(parent, name, "Прев'ю")
            self._label_widget = QLabel("Результат ще не обчислено")
            self._label_widget.setWordWrap(True)
            self._label_widget.setMinimumWidth(250)
            self._label_widget.setMaximumWidth(310)
            self._label_widget.setMaximumHeight(88)
            self._label_widget.setTextInteractionFlags(
                Qt.TextInteractionFlag.TextSelectableByMouse
            )
            self._label_widget.setStyleSheet(
                "QLabel { background: rgba(15, 23, 42, 180);"
                "border: 1px solid rgba(148, 163, 184, 90);"
                "border-radius: 4px; color: #dbeafe; padding: 6px; }"
            )
            self.set_custom_widget(self._label_widget)

        def get_value(self):
            return self._label_widget.text()

        def set_value(self, value):
            text = str(value)
            self._label_widget.setText(text)
            self._label_widget.setToolTip(text)

    class NodeMultilineWidget(NodeBaseWidget):
        def __init__(self, parent=None, name="", label="") -> None:
            super().__init__(parent, name, label)
            self._editor = QPlainTextEdit()
            self._editor.setMinimumWidth(270)
            self._editor.setMaximumWidth(330)
            self._editor.setMinimumHeight(82)
            self._editor.setMaximumHeight(125)
            self._editor.setPlaceholderText("Вставте або введіть текст…")
            self._editor.setStyleSheet(
                "QPlainTextEdit { background: rgba(15, 23, 42, 210);"
                "border: 1px solid rgba(148, 163, 184, 110);"
                "border-radius: 4px; color: #f8fafc; padding: 5px; }"
            )
            self.set_custom_widget(self._editor)
            self._editor.textChanged.connect(self.on_value_changed)

        def get_value(self):
            return self._editor.toPlainText()

        def set_value(self, value):
            text = "" if value is None else str(value)
            if self._editor.toPlainText() != text:
                self._editor.setPlainText(text)

    class NodeCodeWidget(NodeBaseWidget):
        def __init__(self, parent=None, code="") -> None:
            super().__init__(parent, "_nat_source_code", "Python-код")
            self._editor = QPlainTextEdit(code)
            self._editor.setReadOnly(True)
            self._editor.setLineWrapMode(QPlainTextEdit.LineWrapMode.NoWrap)
            self._editor.setMinimumSize(340, 180)
            self._editor.setMaximumWidth(520)
            self.set_custom_widget(self._editor)

        def get_value(self):
            return self._editor.toPlainText()

        def set_value(self, value):
            self._editor.setPlainText(str(value))

    return NodeStatusWidget, NodePreviewWidget, NodeMultilineWidget, NodeCodeWidget


class QuickNodeSearchDialog(QDialog):
    """Швидке вікно пошуку та вставки нод за кнопкою Пробіл / Tab."""

    def __init__(self, registry: NodeRegistry, parent=None) -> None:
        super().__init__(parent)
        self.registry = registry
        self.selected_type_id: str | None = None

        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.Popup)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)

        container = QFrame(self)
        container.setObjectName("quick_container")
        container.setStyleSheet(
            "#quick_container { background-color: #0f172a; border: 2px solid #38bdf8; "
            "border-radius: 10px; }"
        )

        layout = QVBoxLayout(container)
        layout.setContentsMargins(10, 10, 10, 10)

        title = QLabel("🔍 ШВИДКИЙ ПОШУК ТА ДОДАВАННЯ НОДИ (Space / Tab)")
        title.setStyleSheet("color: #38bdf8; font-weight: 700; font-size: 11px; margin-bottom: 2px;")
        layout.addWidget(title)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Пошук ноди (введіть Наказ, DOCX, Заміна)...")
        self.search_input.setStyleSheet(
            "QLineEdit { background: #1e293b; color: #f8fafc; font-size: 13px; "
            "border: 1px solid #475569; border-radius: 6px; padding: 6px 10px; }"
            "QLineEdit:focus { border-color: #38bdf8; }"
        )
        self.search_input.textChanged.connect(self._filter_list)
        layout.addWidget(self.search_input)

        self.list_widget = QListWidget()
        self.list_widget.setStyleSheet(
            "QListWidget { background: #1e293b; color: #f8fafc; border: 1px solid #334155; border-radius: 6px; padding: 4px; font-size: 12px; }"
            "QListWidget::item { padding: 8px; border-bottom: 1px solid #334155; border-radius: 4px; }"
            "QListWidget::item:selected { background: #0284c7; color: #ffffff; font-weight: bold; }"
            "QListWidget::item:hover { background: #0f766e; }"
        )
        self.list_widget.itemDoubleClicked.connect(self._accept_item)
        layout.addWidget(self.list_widget)

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.addWidget(container)

        self.setFixedSize(460, 340)
        self.search_input.installEventFilter(self)
        self._populate("")

    def _populate(self, query: str) -> None:
        self.list_widget.clear()
        query = query.casefold().strip()
        for definition in self.registry.all():
            haystack = f"{definition.category} {definition.name} {definition.description}".casefold()
            if query and query not in haystack:
                continue
            item = QListWidgetItem(f"[{definition.category}]\n{definition.name}")
            item.setData(Qt.ItemDataRole.UserRole, definition.type_id)
            item.setToolTip(f"{definition.name}\n{definition.description}")
            self.list_widget.addItem(item)
        if self.list_widget.count() > 0:
            self.list_widget.setCurrentRow(0)

    def _filter_list(self, text: str) -> None:
        self._populate(text)

    def _accept_item(self, item: QListWidgetItem | None = None) -> None:
        item = item or self.list_widget.currentItem()
        if item:
            self.selected_type_id = item.data(Qt.ItemDataRole.UserRole)
            self.accept()

    def eventFilter(self, watched, event) -> bool:
        if watched is self.search_input and event.type() == QEvent.Type.KeyPress:
            key = event.key()
            if key in (Qt.Key.Key_Down, Qt.Key.Key_Up):
                self.list_widget.keyPressEvent(event)
                return True
            elif key in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
                self._accept_item()
                return True
            elif key == Qt.Key.Key_Escape:
                self.reject()
                return True
        return super().eventFilter(watched, event)


class FullDocumentPreviewWidget(QWidget):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setMinimumWidth(280)
        self.setMaximumWidth(600)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(6, 6, 6, 6)

        self.title_label = QLabel("📄 ПОВНЕ ПРЕВ'Ю ДОКУМЕНТА")
        self.title_label.setStyleSheet("font-weight: 700; color: #38bdf8; font-size: 13px;")
        self.stats_label = QLabel("Виберіть ноду для перегляду згенерованого результату")
        self.stats_label.setStyleSheet("color: #94a3b8; font-size: 11px;")
        layout.addWidget(self.title_label)
        layout.addWidget(self.stats_label)

        self.tabs = QTabWidget()
        self.text_edit = QTextEdit()
        self.text_edit.setReadOnly(True)
        self.text_edit.setPlaceholderText("Текст згенерованого документа відображатиметься тут…")
        self.text_edit.setStyleSheet(
            "QTextEdit { background: #0f172a; color: #f8fafc; font-family: Consolas, monospace; "
            "border: 1px solid #334155; border-radius: 4px; padding: 6px; }"
        )
        self.tabs.addTab(self.text_edit, "Текст")

        self.image_scroll = QScrollArea()
        self.image_label = QLabel("Схема не згенерована")
        self.image_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.image_scroll.setWidget(self.image_label)
        self.image_scroll.setWidgetResizable(True)
        self.tabs.addTab(self.image_scroll, "Схема / Зображення")

        self.table_widget = QTableWidget()
        self.tabs.addTab(self.table_widget, "Таблиця")

        layout.addWidget(self.tabs)

        btn_layout = QHBoxLayout()
        copy_btn = QPushButton("Копіювати текст")
        copy_btn.clicked.connect(self._copy_text)
        btn_layout.addWidget(copy_btn)
        layout.addLayout(btn_layout)

    def _copy_text(self) -> None:
        from PySide6.QtGui import QGuiApplication
        text = self.text_edit.toPlainText()
        if text:
            QGuiApplication.clipboard().setText(text)

    def display_node_output(self, node_title: str, outputs: dict | None) -> None:
        from pathlib import Path
        from nodeautomationtoolkit.core.table_types import DataTable

        self.title_label.setText(f"📄 {node_title}")
        if not outputs:
            self.stats_label.setText("Результат ще не обчислено")
            self.text_edit.setPlainText("Запустіть ноду або увімкніть Live-прев'ю для обчислення")
            self.image_label.setText("Зображення не згенеровано")
            self.table_widget.setRowCount(0)
            return

        image_path = ""
        text_content = ""
        table_obj = None

        for k, v in outputs.items():
            if (k in ("image_path", "path") or k.endswith("_path")) and isinstance(v, str) and Path(v).suffix.casefold() in (".png", ".jpg", ".jpeg"):
                image_path = v
            elif isinstance(v, DataTable):
                table_obj = v
            elif isinstance(v, str) and not text_content:
                text_content = v

        if not text_content:
            from nodeautomationtoolkit.core.preview import format_live_preview
            text_content = format_live_preview(outputs, limit=5000)

        lines = text_content.splitlines()
        words = len(text_content.split())
        self.stats_label.setText(f"Рядків: {len(lines)} · Слів: {words} · {len(text_content)} симв.")
        self.text_edit.setPlainText(text_content)

        if image_path and Path(image_path).is_file():
            pixmap = QPixmap(image_path)
            self.image_label.setPixmap(pixmap.scaledToWidth(380, Qt.TransformationMode.SmoothTransformation))
            self.tabs.setCurrentIndex(1)
        else:
            self.image_label.setText("Зображення не згенеровано")
            if table_obj:
                self.tabs.setCurrentIndex(2)
            else:
                self.tabs.setCurrentIndex(0)

        if table_obj and isinstance(table_obj, DataTable):
            self.table_widget.setRowCount(len(table_obj.rows))
            self.table_widget.setColumnCount(len(table_obj.columns))
            self.table_widget.setHorizontalHeaderLabels(list(table_obj.columns))
            for r, row in enumerate(table_obj.rows):
                for c, val in enumerate(row):
                    self.table_widget.setItem(r, c, QTableWidgetItem(str(val)))
            self.table_widget.resizeColumnsToContents()


def create_nodegraphqt_class(
    definition: NodeDefinition,
    action_handler,
):
    """Build a NodeGraphQt class lazily so core tests need no display server."""
    from NodeGraphQt import BaseNode

    NodeStatusWidget, NodePreviewWidget, NodeMultilineWidget, NodeCodeWidget = (
        _node_widget_classes()
    )

    def init(self) -> None:
        BaseNode.__init__(self)
        self.set_color(*NODE_STATE_COLORS["idle"])
        for port in definition.execution_inputs + definition.inputs:
            self.add_input(port.name, multi_input=False, color=port_color(port))
        for port in definition.execution_outputs + definition.outputs:
            self.add_output(port.name, multi_output=True, color=port_color(port))

        hidden_inputs = set()
        if definition.type_id == "builtin.windows.open_file":
            hidden_inputs = {"title", "file_filter", "initial_folder"}
        for port in definition.inputs:
            if not port.required and port.name not in hidden_inputs:
                label = "Файл" if port.name == "selected_path" else ""
                _add_parameter_widget(
                    self,
                    port,
                    label=label,
                    multiline_widget_class=NodeMultilineWidget,
                )

        self.add_custom_widget(NodeStatusWidget(self.view))
        self.add_button(
            "_nat_help",
            text="Що робить ця нода?",
            tooltip="Показати пояснення входів і виходів",
        )
        self.get_widget("_nat_help").value_changed.connect(
            lambda *_args: action_handler("help", self)
        )
        self.add_button(
            "_nat_run",
            text="Виконати ноду",
            tooltip="Виконати цю ноду та потрібні їй вхідні ноди",
        )
        run_button = self.get_widget("_nat_run")
        run_control = run_button.get_custom_widget()
        run_control.setIcon(
            run_control.style().standardIcon(QStyle.StandardPixmap.SP_MediaPlay)
        )
        run_button.value_changed.connect(lambda *_args: action_handler("run", self))

        if definition.type_id == "builtin.windows.open_file":
            self.add_button(
                "_nat_pick_file",
                text="Вибрати файл…",
                tooltip="Відкрити системне вікно вибору файла",
            )
            picker = self.get_widget("_nat_pick_file")
            picker_control = picker.get_custom_widget()
            picker_control.setIcon(
                picker_control.style().standardIcon(
                    QStyle.StandardPixmap.SP_DialogOpenButton
                )
            )
            picker.value_changed.connect(
                lambda *_args: action_handler("pick_file", self)
            )

        if definition.category == "Згенеровані" or definition.function.__module__.startswith(
            "nat_user_plugin_"
        ):
            try:
                source_code = inspect.getsource(definition.function)
            except (OSError, TypeError):
                source_code = "Код цієї ноди недоступний для перегляду"
            self.add_custom_widget(NodeCodeWidget(self.view, source_code))

        self.add_custom_widget(NodePreviewWidget(self.view))

    category = re.sub(r"[^a-zA-Z0-9]+", "_", definition.category).strip("_")
    attributes: dict[str, Any] = {
        "__identifier__": f"nodeautomationtoolkit.{category or 'Other'}",
        "NODE_NAME": definition.name,
        "NAT_TYPE_ID": definition.type_id,
        "NAT_DEFINITION": definition,
        "__init__": init,
    }
    return type(_safe_class_name(definition.type_id), (BaseNode,), attributes)


def _add_parameter_widget(
    node,
    port: PortDefinition,
    *,
    label: str = "",
    multiline_widget_class=None,
) -> None:
    value = port.default
    if port.data_type == "bool" or isinstance(value, bool):
        node.add_checkbox(port.name, label=label, state=bool(value))
    elif port.data_type == "int" or isinstance(value, int):
        node.add_spinbox(
            port.name,
            label=label,
            value=int(value or 0),
            min_value=-1_000_000_000,
            max_value=1_000_000_000,
        )
    elif port.data_type == "float" or isinstance(value, float):
        node.add_spinbox(
            port.name,
            label=label,
            value=float(value or 0),
            min_value=-1_000_000_000,
            max_value=1_000_000_000,
            double=True,
        )
    elif multiline_widget_class is not None and port.name in {
        "fields_json",
        "markers_text",
        "names_text",
        "replacement_text",
        "text",
    }:
        widget = multiline_widget_class(node.view, port.name, label)
        widget.set_value(value)
        node.add_custom_widget(widget)
    else:
        node.add_text_input(
            port.name,
            label=label,
            text="" if value is None else str(value),
        )


class NodeGraphQtEditor(QWidget):
    graph_changed = Signal()
    message = Signal(str)

    def __init__(self, registry: NodeRegistry, parent=None) -> None:
        super().__init__(parent)
        from NodeGraphQt import NodeGraph, NodesPaletteWidget, PropertiesBinWidget
        from NodeGraphQt.constants import ViewerEnum

        self.registry = registry
        self._registered_type_ids: set[str] = set()
        self._graph_name = "Новий сценарій"
        self._loading = False
        self._live_values: dict[str, dict[str, Any]] = {}
        self._preview_generation = 0
        self._pending_live_node: str | None = None
        self._live_timer = QTimer(self)
        self._live_timer.setSingleShot(True)
        self._live_timer.setInterval(240)
        self._live_timer.timeout.connect(self._start_live_preview)
        self._thread_pool = QThreadPool.globalInstance()

        self.graph = NodeGraph()
        self.graph.set_acyclic(True)
        self.graph.set_background_color(11, 17, 32)
        self.graph.set_grid_color(30, 41, 59)
        self.graph.set_grid_mode(ViewerEnum.GRID_DISPLAY_DOTS.value)
        self._register_definitions()

        self.palette = NodesPaletteWidget(node_graph=self.graph)
        self.left_preview = FullDocumentPreviewWidget()

        self.left_tabs = QTabWidget()
        self.left_tabs.addTab(self.palette, "Палітра нод")
        self.left_tabs.addTab(self.left_preview, "📄 Прев'ю документа")

        self.properties = PropertiesBinWidget(node_graph=self.graph)
        self.properties.setVisible(False)
        splitter = QSplitter()
        splitter.addWidget(self.left_tabs)
        splitter.addWidget(self.graph.widget)
        splitter.addWidget(self.properties)
        splitter.setSizes([320, 900, 300])

        self.live_toggle = QCheckBox("Live-прев'ю")
        self.live_toggle.setChecked(True)
        self.live_toggle.setToolTip(
            "Автоматично перераховувати безпечні ноди після зміни параметрів"
        )
        self.live_toggle.toggled.connect(self._on_live_toggled)
        run_selected = QPushButton("Виконати вибрану")
        run_selected.setIcon(
            self.style().standardIcon(QStyle.StandardPixmap.SP_MediaPlay)
        )
        run_selected.clicked.connect(self.run_selected_node)

        group_backdrop_btn = QPushButton("🔲 Згрупувати рамкою")
        group_backdrop_btn.setIcon(
            self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogNewFolder)
        )
        group_backdrop_btn.setToolTip("Об'єднати виділені ноди у візуальну рамку з підписом (Ctrl+G)")
        group_backdrop_btn.clicked.connect(lambda: self.create_group_backdrop())

        delete_selected = QPushButton("Видалити")
        delete_selected.setIcon(
            self.style().standardIcon(QStyle.StandardPixmap.SP_TrashIcon)
        )
        delete_selected.clicked.connect(self.delete_selected_nodes)
        self.live_status = QLabel("LIVE · очікує змін")
        self.live_status.setStyleSheet("color: #93c5fd; padding: 0 8px;")
        self.properties_toggle = QPushButton("Права панель")
        self.properties_toggle.setCheckable(True)
        self.properties_toggle.setChecked(False)
        self.properties_toggle.toggled.connect(self.properties.setVisible)

        controls = QHBoxLayout()
        controls.setContentsMargins(8, 5, 8, 5)
        controls.addWidget(self.live_toggle)
        controls.addWidget(run_selected)
        controls.addWidget(group_backdrop_btn)
        controls.addWidget(delete_selected)
        controls.addWidget(self.properties_toggle)
        controls.addWidget(self.live_status)
        controls.addStretch(1)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addLayout(controls)
        layout.addWidget(splitter)

        delete_action = QAction("Видалити вибране", self)
        delete_action.setShortcut(QKeySequence(Qt.Key.Key_Delete))
        delete_action.setShortcutContext(Qt.ShortcutContext.WidgetWithChildrenShortcut)
        delete_action.triggered.connect(self._delete_from_shortcut)
        self.addAction(delete_action)
        run_action = QAction("Виконати вибрану ноду", self)
        run_action.setShortcut(QKeySequence("F6"))
        run_action.setShortcutContext(Qt.ShortcutContext.WidgetWithChildrenShortcut)
        run_action.triggered.connect(self.run_selected_node)
        self.addAction(run_action)

        group_action = QAction("Згрупувати рамкою", self)
        group_action.setShortcut(QKeySequence("Ctrl+G"))
        group_action.setShortcutContext(Qt.ShortcutContext.WidgetWithChildrenShortcut)
        group_action.triggered.connect(lambda: self.create_group_backdrop())
        self.addAction(group_action)

        space_action = QAction("Швидкий пошук та додавання ноди (Space)", self)
        space_action.setShortcut(QKeySequence(Qt.Key.Key_Space))
        space_action.setShortcutContext(Qt.ShortcutContext.WidgetWithChildrenShortcut)
        space_action.triggered.connect(self._open_quick_node_search)
        self.addAction(space_action)

        tab_action = QAction("Швидкий пошук та додавання ноди (Tab)", self)
        tab_action.setShortcut(QKeySequence(Qt.Key.Key_Tab))
        tab_action.setShortcutContext(Qt.ShortcutContext.WidgetWithChildrenShortcut)
        tab_action.triggered.connect(self._open_quick_node_search)
        self.addAction(tab_action)

        self.graph.node_created.connect(self._on_node_created)
        self.graph.nodes_deleted.connect(self._on_nodes_deleted)
        self.graph.port_connected.connect(self._on_connection_changed)
        self.graph.port_disconnected.connect(self._on_connection_changed)
        self.graph.property_changed.connect(self._on_property_changed)
        self.graph.node_selection_changed.connect(self._on_node_selection_changed)

        viewer = self.graph.viewer()
        viewer.setAcceptDrops(True)
        viewer.viewport().setAcceptDrops(True)
        self._drop_filter = _FileDropFilter(self)
        viewer.installEventFilter(self._drop_filter)
        viewer.viewport().installEventFilter(self._drop_filter)

    def _register_definitions(self) -> None:
        node_menu = self.graph.get_context_menu("nodes")
        for definition in self.registry.all():
            if definition.type_id in self._registered_type_ids:
                continue
            node_class = create_nodegraphqt_class(
                definition,
                self._handle_node_action,
            )
            self.graph.register_node(node_class, alias=definition.type_id)
            node_menu.add_command(
                "Виконати цю ноду",
                func=lambda _graph, node: self.run_node(node.id),
                node_class=node_class,
            )
            node_menu.add_command(
                "Видалити ноду",
                func=lambda graph, node: graph.delete_node(node),
                node_class=node_class,
            )
            self._registered_type_ids.add(definition.type_id)

    def reload_definitions(self) -> None:
        self._register_definitions()

    def create_group_backdrop(self, title: str | None = None) -> None:
        from PySide6.QtWidgets import QInputDialog, QLineEdit

        selected_nodes = [node for node in self.graph.selected_nodes() if type(node).__name__ != "BackdropNode"]

        if not title:
            default_name = "Група автоматизації" if selected_nodes else "Нова рамка"
            name, ok = QInputDialog.getText(
                self,
                "Створити рамку групи",
                "Введіть назву рамки/групи:",
                QLineEdit.EchoMode.Normal,
                default_name,
            )
            if not ok or not name.strip():
                return
            title = name.strip()

        backdrop = self.graph.create_node("nodeGraphQt.nodes.BackdropNode")
        backdrop.set_name(title)

        if selected_nodes:
            backdrop.wrap_nodes(selected_nodes)
        else:
            view_center = self.graph.viewer().mapToScene(self.graph.viewer().rect().center())
            backdrop.set_pos(view_center.x() - 100, view_center.y() - 100)

        self.graph_changed.emit()
        self.message.emit(f"Створено рамку групи: {title}")

    def _handle_node_action(self, action: str, node) -> None:
        if action == "pick_file":
            self._pick_file(node)
        elif action == "run":
            self.run_node(node.id)
        elif action == "help":
            self._show_node_help(node)

    def _show_node_help(self, node) -> None:
        definition: NodeDefinition = node.NAT_DEFINITION
        inputs = "\n".join(
            f"• {port.name} ({port.data_type}){' — обов’язковий' if port.required else ''}"
            for port in definition.inputs
        ) or "• немає"
        outputs = "\n".join(
            f"• {port.name} ({port.data_type})" for port in definition.outputs
        ) or "• немає"
        if definition.dynamic_outputs:
            outputs += "\n• додаткові виходи створюються після виконання"
        QMessageBox.information(
            self,
            definition.name,
            f"{definition.description or 'Опис відсутній'}\n\nВХОДИ\n{inputs}\n\nВИХОДИ\n{outputs}",
        )

    def _pick_file(self, node) -> None:
        definition: NodeDefinition = node.NAT_DEFINITION
        defaults = {port.name: port.default for port in definition.inputs}
        title = node.get_property("title") if node.has_property("title") else None
        file_filter = (
            node.get_property("file_filter") if node.has_property("file_filter") else None
        )
        initial_folder = (
            node.get_property("initial_folder")
            if node.has_property("initial_folder")
            else None
        )
        current_path = (
            node.get_property("selected_path")
            if node.has_property("selected_path")
            else ""
        )
        if current_path:
            initial_folder = current_path
        path, _selected_filter = QFileDialog.getOpenFileName(
            self,
            title or defaults.get("title") or "Виберіть файл",
            initial_folder or defaults.get("initial_folder") or "",
            file_filter or defaults.get("file_filter") or "Усі файли (*.*)",
        )
        if not path:
            return
        node.set_property("selected_path", path)
        self._invalidate_from(node.id)
        self._schedule_live_preview(node.id)
        self.graph_changed.emit()
    def _on_node_selection_changed(self, selected_nodes, unselected_nodes) -> None:
        if not selected_nodes:
            return
        node = selected_nodes[0]
        outputs = self._live_values.get(node.id)
        definition: NodeDefinition | None = getattr(node, "NAT_DEFINITION", None)
        node_name = node.name() if hasattr(node, "name") else "Нода"
        self.left_preview.display_node_output(node_name, outputs)

        if definition and (
            definition.category in ("Результат", "Word · Пакет", "Word", "Наказ")
            or "output" in definition.type_id
            or "show" in definition.type_id
            or "save" in definition.type_id
            or "visualize" in definition.type_id
        ):
            self.left_tabs.setCurrentIndex(1)

    def _handle_dropped_file_paths(self, paths: list[str], drop_pos) -> None:
        from pathlib import Path

        viewer = self.graph.viewer()
        scene_pos = viewer.mapToScene(drop_pos) if hasattr(viewer, "mapToScene") else QPointF(0, 0)
        offset = QPointF(0, 0)

        for file_path in paths:
            ext = Path(file_path).suffix.casefold()
            pos_list = [float(scene_pos.x() + offset.x()), float(scene_pos.y() + offset.y())]

            if ext == ".docx":
                node = self.graph.create_node("builtin.word.read_docx", pos=pos_list)
                if node and node.has_property("path"):
                    node.set_property("path", file_path)
                self.message.emit(f"Перетягнуто DOCX-файл: {Path(file_path).name}")
            elif ext in (".xlsx", ".csv"):
                node = self.graph.create_node("builtin.order.read_recipient_mapping", pos=pos_list)
                if node and node.has_property("path"):
                    node.set_property("path", file_path)
                self.message.emit(f"Перетягнуто таблицю відповідностей: {Path(file_path).name}")
            else:
                node = self.graph.create_node("builtin.windows.open_file", pos=pos_list)
                if node and node.has_property("selected_path"):
                    node.set_property("selected_path", file_path)
                self.message.emit(f"Перетягнуто файл: {Path(file_path).name}")

            offset += QPointF(40, 40)

        self._schedule_live_preview()
        self.graph_changed.emit()

    def _open_quick_node_search(self) -> None:
        from PySide6.QtGui import QCursor
        cursor_pos = QCursor.pos()
        dialog = QuickNodeSearchDialog(self.registry, parent=self)
        dialog.move(cursor_pos.x() - 230, cursor_pos.y() - 170)
        if dialog.exec() == QDialog.DialogCode.Accepted and dialog.selected_type_id:
            viewer = self.graph.viewer()
            local_pos = viewer.mapFromGlobal(cursor_pos)
            scene_pos = viewer.mapToScene(local_pos) if hasattr(viewer, "mapToScene") else QPointF(0, 0)
            pos_list = [float(scene_pos.x()), float(scene_pos.y())]
            node = self.graph.create_node(dialog.selected_type_id, pos=pos_list)
            if node:
                self.message.emit(f"Додано ноду з швидкого пошуку: {node.name()}")
            self.graph_changed.emit()

    def delete_selected_nodes(self) -> None:
        nodes = self.graph.selected_nodes()
        if not nodes:
            self.message.emit("Спочатку виберіть ноду")
            return
        count = len(nodes)
        self.graph.delete_nodes(nodes)
        self._live_values.clear()
        self.message.emit(f"Видалено нод: {count}")

    def _delete_from_shortcut(self) -> None:
        focus = QApplication.focusWidget()
        if isinstance(focus, (QLineEdit, QPlainTextEdit, QTextEdit)):
            return
        self.delete_selected_nodes()

    def run_selected_node(self) -> None:
        nodes = self.graph.selected_nodes()
        if not nodes:
            self.message.emit("Спочатку виберіть ноду")
            return
        self.run_node(nodes[0].id)

    def run_graph(self) -> None:
        graph = self.graph_model()
        self._live_values.clear()
        self._invalidate_from(None)
        try:
            result = GraphExecutor(self.registry).execute(
                graph,
                on_node_started=lambda current: self._set_node_state(
                    current, "running", "ВИКОНУЄТЬСЯ"
                ),
                on_node_finished=lambda current, outputs: self._show_node_outputs(
                    current, outputs
                ),
            )
            self._live_values.update(result.values)
            self.message.emit(f"Граф виконано. Нод: {len(result.order)}")
        except Exception as error:  # noqa: BLE001 - graph run boundary
            self.message.emit(f"ПОМИЛКА ГРАФА: {error}")
            raise

    def run_node(self, node_id: str) -> None:
        graph = self.graph_model()
        self._invalidate_from(node_id)
        executor = GraphExecutor(self.registry)
        try:
            result = executor.execute_target(
                graph,
                node_id,
                on_node_started=lambda current: self._set_node_state(
                    current, "running", "ВИКОНУЄТЬСЯ"
                ),
                on_node_finished=lambda current, outputs: self._show_node_outputs(
                    current, outputs
                ),
            )
            self._live_values.update(result.values)
            self.message.emit(f"Виконано нод: {len(result.order)}")
            if self.live_toggle.isChecked():
                self._schedule_live_preview(None)
        except Exception as error:  # noqa: BLE001 - explicit node run boundary
            self._set_node_state(node_id, "error", "ПОМИЛКА", str(error))
            self.message.emit(f"ПОМИЛКА НОДИ: {error}")

    def _on_node_created(self, node) -> None:
        if self._loading:
            return
        self._set_node_state(node.id, "idle", "ОЧІКУЄ")
        self.graph_changed.emit()

    def _on_nodes_deleted(self, node_ids) -> None:
        for node_id in node_ids:
            self._live_values.pop(node_id, None)
        if not self._loading:
            self.graph_changed.emit()

    def _on_connection_changed(self, _first, _second) -> None:
        if self._loading:
            return
        self._live_values.clear()
        self.graph_changed.emit()
        self._schedule_live_preview(None)

    def _on_property_changed(self, node, name: str, _value: Any) -> None:
        if self._loading or name.startswith("_nat_"):
            return
        self.graph_changed.emit()
        definition: NodeDefinition | None = getattr(node, "NAT_DEFINITION", None)
        if definition is None:
            return
        if definition.preview_policy == "auto" or name.startswith("selected_"):
            self._invalidate_from(node.id)
            self._schedule_live_preview(node.id)

    def _schedule_live_preview(self, node_id: str | None) -> None:
        if not self.live_toggle.isChecked():
            return
        self._preview_generation += 1
        self._pending_live_node = node_id
        self._live_timer.start()

    def _start_live_preview(self) -> None:
        graph = self.graph_model()
        generation = self._preview_generation
        self.live_status.setText("LIVE · обчислення…")
        self.live_status.setStyleSheet("color: #fbbf24; padding: 0 8px;")
        worker = _PreviewWorker(
            generation,
            self.registry,
            graph,
            self._pending_live_node,
            dict(self._live_values),
        )
        worker.signals.finished.connect(self._apply_live_result)
        self._thread_pool.start(worker)

    def _on_live_toggled(self, checked: bool) -> None:
        if checked:
            self._schedule_live_preview(None)
        else:
            self._preview_generation += 1
            self._live_timer.stop()
            self.live_status.setText("LIVE · вимкнено")
            self.live_status.setStyleSheet("color: #94a3b8; padding: 0 8px;")

    def _apply_live_result(self, generation: int, result: PreviewResult) -> None:
        if generation != self._preview_generation:
            return
        for node_id in result.order:
            if node_id in result.values:
                self._show_node_outputs(node_id, result.values[node_id])
        for node_id, error in result.errors.items():
            self._set_node_state(node_id, "error", "ПОМИЛКА", error)
        if result.errors:
            self.live_status.setText(f"LIVE · помилок: {len(result.errors)}")
            self.live_status.setStyleSheet("color: #f87171; padding: 0 8px;")
        else:
            self.live_status.setText("LIVE · актуально")
            self.live_status.setStyleSheet("color: #4ade80; padding: 0 8px;")

    def _show_node_outputs(self, node_id: str, outputs: dict[str, Any]) -> None:
        self._sync_dynamic_outputs(node_id, outputs)
        self._set_node_state(
            node_id,
            "success",
            "ГОТОВО",
            format_live_preview(outputs),
        )

    def _sync_dynamic_outputs(self, node_id: str, outputs: dict[str, Any]) -> None:
        node = self.graph.get_node_by_id(node_id)
        if node is None:
            return
        definition: NodeDefinition | None = getattr(node, "NAT_DEFINITION", None)
        if definition is None or not definition.dynamic_outputs:
            return
        existing = set(node.outputs())
        changed = False
        for name, value in outputs.items():
            if name in existing:
                continue
            data_type = "str" if isinstance(value, str) else "Any"
            color = TYPE_COLORS.get(data_type, (45, 212, 191))
            node.add_output(name, multi_output=True, color=color)
            changed = True
        if changed and not self._loading:
            self.graph_changed.emit()

    def _set_node_state(
        self,
        node_id: str,
        state: str,
        label: str,
        preview: str | None = None,
    ) -> None:
        node = self.graph.get_node_by_id(node_id)
        if node is None:
            return
        node.set_color(*NODE_STATE_COLORS.get(state, NODE_STATE_COLORS["idle"]))
        status_widget = node.get_widget("_nat_status")
        if status_widget is not None:
            status_widget.set_state(state, label)
        if preview is not None:
            preview_widget = node.get_widget("_nat_preview")
            if preview_widget is not None:
                preview_widget.set_value(preview)

    def _invalidate_from(self, node_id: str | None) -> set[str]:
        graph = self.graph_model()
        if node_id is None:
            invalidated = {node.id for node in graph.nodes}
        else:
            outgoing: dict[str, list[str]] = {}
            for connection in graph.connections:
                if connection.kind == "data":
                    outgoing.setdefault(connection.source_node, []).append(
                        connection.target_node
                    )
            invalidated = set()
            queue = deque([node_id])
            while queue:
                current = queue.popleft()
                if current in invalidated:
                    continue
                invalidated.add(current)
                queue.extend(outgoing.get(current, []))
        for current in invalidated:
            self._live_values.pop(current, None)
            self._set_node_state(current, "idle", "ОЧІКУЄ")
        return invalidated

    def set_graph_model(self, model: GraphModel) -> None:
        self._loading = True
        self._graph_name = model.name
        self._live_values.clear()
        try:
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
                for name, data_type in item.dynamic_outputs.items():
                    if name not in node.outputs():
                        node.add_output(
                            name,
                            multi_output=True,
                            color=TYPE_COLORS.get(data_type, (45, 212, 191)),
                        )
                for name, value in item.parameters.items():
                    if node.has_property(name):
                        node.set_property(name, value, push_undo=False)
                created[item.id] = node
            for connection in model.connections:
                source = created.get(connection.source_node)
                target = created.get(connection.target_node)
                if source is None or target is None:
                    continue
                out_port = source.get_output(connection.source_port)
                in_port = target.get_input(connection.target_port)
                if out_port is not None and in_port is not None:
                    out_port.connect_to(in_port, push_undo=False)
            self.graph.clear_undo_stack()
        finally:
            self._loading = False
        self.graph_changed.emit()
        self._schedule_live_preview(None)

    def graph_model(self) -> GraphModel:
        model = GraphModel(name=self._graph_name)
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
                dynamic_outputs={
                    name: "str"
                    for name in node.outputs()
                    if name not in {port.name for port in definition.outputs}
                },
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
