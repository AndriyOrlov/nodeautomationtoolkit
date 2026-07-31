from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QSettings, Qt
from PySide6.QtGui import QAction, QKeySequence
from PySide6.QtWidgets import (
    QFileDialog,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QSplitter,
    QTabWidget,
    QToolBar,
    QVBoxLayout,
    QWidget,
)

from nodeautomationtoolkit.core.executor import GraphExecutor
from nodeautomationtoolkit.core.models import GraphModel
from nodeautomationtoolkit.core.project import load_graph, save_graph
from nodeautomationtoolkit.core.registry import NodeRegistry

from .automation_assistant_dialog import AutomationAssistantDialog
from .graph_view import GraphScene, GraphView, NodePalette
from .node_assistant_dialog import NodeAssistantDialog
from .nodegraphqt_editor import NodeGraphQtEditor
from .properties import PropertiesPanel


class MainWindow(QMainWindow):
    def __init__(self, registry: NodeRegistry, plugin_dir: Path) -> None:
        super().__init__()
        self.registry = registry
        self.plugin_dir = plugin_dir
        self.current_path: Path | None = None
        self.dirty = False
        self.setWindowTitle("Node Automation Toolkit")
        self.resize(1500, 900)

        self.scene = GraphScene(registry)
        self.view = GraphView(self.scene)
        self.blueprint = NodeGraphQtEditor(registry)
        self.palette = NodePalette()
        self.search = QLineEdit()
        self.search.setPlaceholderText("Пошук нод…")
        self.search.textChanged.connect(lambda text: self.palette.populate(self.registry, text))
        self.palette.populate(self.registry)

        palette_widget = QWidget()
        palette_layout = QVBoxLayout(palette_widget)
        palette_layout.setContentsMargins(6, 6, 6, 6)
        palette_layout.addWidget(self.search)
        palette_layout.addWidget(self.palette)

        self.properties = PropertiesPanel()
        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.setMaximumBlockCount(1000)

        horizontal = QSplitter(Qt.Orientation.Horizontal)
        horizontal.addWidget(palette_widget)
        horizontal.addWidget(self.view)
        horizontal.addWidget(self.properties)
        horizontal.setSizes([260, 950, 290])

        self.editor_tabs = QTabWidget()
        self.editor_tabs.addTab(self.blueprint, "Blueprint 2.0")
        self.editor_tabs.addTab(horizontal, "Сумісність 1.0")

        vertical = QSplitter(Qt.Orientation.Vertical)
        vertical.addWidget(self.editor_tabs)
        vertical.addWidget(self.log)
        vertical.setSizes([750, 150])
        self.setCentralWidget(vertical)

        self.scene.node_selected.connect(self.properties.show_node)
        self.scene.graph_changed.connect(self._mark_dirty)
        self.blueprint.graph_changed.connect(self._mark_dirty)
        self.scene.message.connect(self._log)
        self.properties.changed.connect(self._mark_dirty)
        self._create_toolbar()
        self._restore_window_state()
        self._report_plugin_errors()
        self._update_title()

    def _create_toolbar(self) -> None:
        toolbar = QToolBar("Основні дії")
        toolbar.setMovable(False)
        self.addToolBar(toolbar)

        actions = [
            ("Новий", QKeySequence.StandardKey.New, self.new_graph),
            ("Відкрити", QKeySequence.StandardKey.Open, self.open_graph),
            ("Зберегти", QKeySequence.StandardKey.Save, self.save_graph),
            ("Зберегти як", QKeySequence.StandardKey.SaveAs, self.save_graph_as),
            ("Запустити", "F5", self.run_graph),
            ("AI-сценарій", "Ctrl+Shift+A", self.open_automation_assistant),
            ("AI-нода", "Ctrl+Shift+N", self.open_node_assistant),
            ("Оновити плагіни", "Ctrl+R", self.reload_plugins),
        ]
        for text, shortcut, callback in actions:
            action = QAction(text, self)
            action.setShortcut(shortcut)
            action.triggered.connect(callback)
            toolbar.addAction(action)

    def open_node_assistant(self) -> None:
        dialog = NodeAssistantDialog(self.plugin_dir, self)
        dialog.node_installed.connect(lambda path: self._log(f"Створено AI-ноду: {path}"))
        dialog.exec()

    def open_automation_assistant(self) -> None:
        dialog = AutomationAssistantDialog(self.registry, self._current_graph(), self)
        dialog.graph_created.connect(self._set_graph_everywhere)
        dialog.exec()

    def new_graph(self) -> None:
        if not self._confirm_discard():
            return
        self.scene.set_graph(GraphModel())
        self.blueprint.set_graph_model(GraphModel())
        self.current_path = None
        self.dirty = False
        self._update_title()
        self.log.clear()

    def open_graph(self) -> None:
        if not self._confirm_discard():
            return
        filename, _ = QFileDialog.getOpenFileName(
            self,
            "Відкрити сценарій",
            "",
            "Node Automation Toolkit (*.nat.json);;JSON (*.json)",
        )
        if not filename:
            return
        try:
            graph = load_graph(Path(filename))
            missing = [node.type_id for node in graph.nodes if not self._has_node(node.type_id)]
            if missing:
                raise ValueError("Не встановлені ноди:\n" + "\n".join(sorted(set(missing))))
            self.scene.set_graph(graph)
            self.blueprint.set_graph_model(graph)
            self.current_path = Path(filename)
            self.dirty = False
            self._update_title()
            self._log(f"Відкрито: {filename}")
        except Exception as error:  # noqa: BLE001 - UI boundary
            QMessageBox.critical(self, "Помилка відкриття", str(error))

    def save_graph(self) -> bool:
        if self.current_path is None:
            return self.save_graph_as()
        try:
            save_graph(self._current_graph(), self.current_path)
            self.dirty = False
            self._update_title()
            self._log(f"Збережено: {self.current_path}")
            return True
        except Exception as error:  # noqa: BLE001 - UI boundary
            QMessageBox.critical(self, "Помилка збереження", str(error))
            return False

    def save_graph_as(self) -> bool:
        filename, _ = QFileDialog.getSaveFileName(
            self,
            "Зберегти сценарій",
            str(self.current_path or Path.cwd() / "scenario.nat.json"),
            "Node Automation Toolkit (*.nat.json)",
        )
        if not filename:
            return False
        if not filename.lower().endswith(".nat.json"):
            filename += ".nat.json"
        self.current_path = Path(filename)
        return self.save_graph()

    def run_graph(self) -> None:
        self._log("Запуск сценарію…")
        try:
            graph = self._current_graph()
            names = {
                node.id: self.registry.get(node.type_id).name for node in graph.nodes
            }
            executor = GraphExecutor(self.registry)
            result = executor.execute(
                graph,
                on_node_started=lambda node_id: self._log(
                    f"{names[node_id]}: виконується"
                ),
                on_node_finished=lambda node_id, values: self._log(
                    f"{names[node_id]}: готово: {self._short_result(values)}"
                ),
            )
            self._log(f"Готово. Виконано нод: {len(result.order)}")
        except Exception as error:  # noqa: BLE001 - execution boundary
            self._log(f"ПОМИЛКА: {error}")
            QMessageBox.critical(self, "Помилка виконання", str(error))

    def reload_plugins(self) -> None:
        if self._current_graph().nodes:
            self._log("Плагіни не оновлено: спочатку відкрийте порожній сценарій")
            return
        self.registry.reload(self.plugin_dir)
        self.palette.populate(self.registry, self.search.text())
        self._report_plugin_errors()
        self._log(f"Завантажено нод: {len(self.registry.all())}")

    def _current_graph(self) -> GraphModel:
        if self.editor_tabs.currentWidget() is self.blueprint:
            return self.blueprint.graph_model()
        return self.scene.graph

    def _set_graph_everywhere(self, graph: GraphModel) -> None:
        self.scene.set_graph(graph)
        self.blueprint.set_graph_model(graph)

    def _has_node(self, type_id: str) -> bool:
        try:
            self.registry.get(type_id)
            return True
        except KeyError:
            return False

    def _set_node_status(self, node_id: str, status: str) -> None:
        item = self.scene.node_items[node_id]
        self._log(f"{item.definition.name}: {status}")

    @staticmethod
    def _short_result(values: dict) -> str:
        text = repr(values)
        return text if len(text) <= 120 else text[:117] + "…"

    def _report_plugin_errors(self) -> None:
        for error in self.registry.errors:
            self._log(f"ПОМИЛКА ПЛАГІНА: {error}")

    def _mark_dirty(self) -> None:
        self.dirty = True
        self._update_title()

    def _update_title(self) -> None:
        name = self.current_path.name if self.current_path else self.scene.graph.name
        marker = " *" if self.dirty else ""
        self.setWindowTitle(f"{name}{marker} — Node Automation Toolkit")

    def _log(self, message: str) -> None:
        self.log.appendPlainText(message)

    def _confirm_discard(self) -> bool:
        if not self.dirty:
            return True
        answer = QMessageBox.question(
            self,
            "Незбережені зміни",
            "Зберегти зміни перед продовженням?",
            QMessageBox.StandardButton.Save
            | QMessageBox.StandardButton.Discard
            | QMessageBox.StandardButton.Cancel,
        )
        if answer == QMessageBox.StandardButton.Save:
            return self.save_graph()
        return answer == QMessageBox.StandardButton.Discard

    def _restore_window_state(self) -> None:
        settings = QSettings()
        if geometry := settings.value("window/geometry"):
            self.restoreGeometry(geometry)

    def closeEvent(self, event) -> None:
        if not self._confirm_discard():
            event.ignore()
            return
        QSettings().setValue("window/geometry", self.saveGeometry())
        event.accept()
