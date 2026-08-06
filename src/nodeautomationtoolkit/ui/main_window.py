from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QSettings, QStandardPaths, Qt, QTimer
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
from nodeautomationtoolkit.core.patching import install_patch
from nodeautomationtoolkit.core.project import load_graph, save_graph
from nodeautomationtoolkit.core.registry import NodeRegistry
from nodeautomationtoolkit.core.templates import (
    build_order_senders_graph,
    build_word_smoke_graph,
)

from .automation_assistant_dialog import AutomationAssistantDialog
from .embedded_model_dialog import EmbeddedModelDialog
from .graph_view import GraphScene, GraphView, NodePalette
from .node_assistant_dialog import NodeAssistantDialog
from .nodegraphqt_editor import NodeGraphQtEditor
from .order_analysis_dialog import OrderAnalysisDialog
from .properties import PropertiesPanel


class MainWindow(QMainWindow):
    def __init__(self, registry: NodeRegistry, plugin_dir: Path) -> None:
        super().__init__()
        self.registry = registry
        self.plugin_dir = plugin_dir
        self.current_path: Path | None = None
        self.dirty = False
        self._suppress_graph_changes = False
        self.app_data_dir = Path(
            QStandardPaths.writableLocation(QStandardPaths.StandardLocation.AppDataLocation)
        )
        documents_dir = Path(
            QStandardPaths.writableLocation(QStandardPaths.StandardLocation.DocumentsLocation)
        )
        self.graphs_dir = documents_dir / "Node Automation Toolkit" / "Graphs"
        self.autosave_path = self.app_data_dir / "autosave.nat.json"
        self.graphs_dir.mkdir(parents=True, exist_ok=True)
        self.app_data_dir.mkdir(parents=True, exist_ok=True)
        self.autosave_timer = QTimer(self)
        self.autosave_timer.setSingleShot(True)
        self.autosave_timer.setInterval(700)
        self.autosave_timer.timeout.connect(self._autosave_graph)
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
        self.compatibility_widget = horizontal

        vertical = QSplitter(Qt.Orientation.Vertical)
        vertical.addWidget(self.editor_tabs)
        vertical.addWidget(self.log)
        vertical.setSizes([750, 150])
        self.setCentralWidget(vertical)

        self.scene.node_selected.connect(self.properties.show_node)
        self.scene.graph_changed.connect(self._mark_dirty)
        self.blueprint.graph_changed.connect(self._mark_dirty)
        self.blueprint.message.connect(self._log)
        self.scene.message.connect(self._log)
        self.properties.changed.connect(self._mark_dirty)
        self._create_toolbar()
        self._restore_window_state()
        self._restore_working_graph()
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
            ("🔍 Аналіз відправників", "Ctrl+Shift+S", self.open_order_senders_preset),
            ("🔍 Аналіз наказу", "Ctrl+Shift+O", self.open_order_analysis),
            ("AI-сценарій", "Ctrl+Shift+A", self.open_automation_assistant),
            ("AI-нода", "Ctrl+Shift+N", self.open_node_assistant),
            ("Локальна модель", "Ctrl+Shift+L", self.open_embedded_model),
            ("Оновити плагіни", "Ctrl+R", self.reload_plugins),
            ("Встановити патч", "Ctrl+Shift+P", self.install_offline_patch),
        ]
        for text, shortcut, callback in actions:
            action = QAction(text, self)
            action.setShortcut(shortcut)
            action.triggered.connect(callback)
            toolbar.addAction(action)
        toolbar.addSeparator()
        self.compatibility_action = QAction("Форма 1.0", self)
        self.compatibility_action.setCheckable(True)
        self.compatibility_action.setChecked(False)
        self.compatibility_action.setToolTip("Показати старий редактор лише за потреби")
        self.compatibility_action.toggled.connect(self._toggle_compatibility_tab)
        toolbar.addAction(self.compatibility_action)

    def _toggle_compatibility_tab(self, visible: bool) -> None:
        index = self.editor_tabs.indexOf(self.compatibility_widget)
        if visible and index < 0:
            self.editor_tabs.addTab(self.compatibility_widget, "Сумісність 1.0")
            self.editor_tabs.setCurrentWidget(self.compatibility_widget)
        elif not visible and index >= 0:
            self.editor_tabs.setCurrentWidget(self.blueprint)
            self.editor_tabs.removeTab(index)

    def install_offline_patch(self) -> None:
        filename, _ = QFileDialog.getOpenFileName(
            self,
            "Встановити офлайн-патч",
            "",
            "Node Automation Toolkit patch (*.natpatch.zip);;ZIP (*.zip)",
        )
        if not filename:
            return
        try:
            installed = install_patch(Path(filename), self.app_data_dir)
        except Exception as error:  # noqa: BLE001 - local patch boundary
            QMessageBox.critical(self, "Патч не встановлено", str(error))
            return
        self._log(f"Встановлено патч {installed.version}")
        QMessageBox.information(
            self,
            "Патч встановлено",
            f"Версію {installed.version} встановлено.\n\n"
            "Збережіть граф і один раз перезапустіть програму.",
        )

    def open_order_analysis(self) -> None:
        dialog = OrderAnalysisDialog(self.registry, self.plugin_dir, self)
        dialog.graph_created.connect(self._apply_assistant_graph)
        dialog.exec()

    def open_node_assistant(self) -> None:
        dialog = NodeAssistantDialog(self.plugin_dir, self)
        dialog.node_installed.connect(lambda path: self._log(f"Створено AI-ноду: {path}"))
        dialog.exec()

    def open_automation_assistant(self) -> None:
        dialog = AutomationAssistantDialog(
            self.registry,
            self._current_graph(),
            self.plugin_dir,
            self,
        )
        dialog.graph_created.connect(self._apply_assistant_graph)
        dialog.exec()

    def open_embedded_model(self) -> None:
        EmbeddedModelDialog(self).exec()

    def _apply_assistant_graph(self, graph: GraphModel) -> None:
        self.blueprint.reload_definitions()
        self._set_graph_everywhere(graph)

    def new_graph(self) -> None:
        if not self._confirm_discard():
            return
        self._set_graph_everywhere(GraphModel())
        self.current_path = None
        self.dirty = False
        self._autosave_graph()
        QSettings().setValue("workspace/dirty", False)
        self._update_title()
        self.log.clear()

    def open_word_smoke_graph(self) -> None:
        if not self._confirm_discard():
            return
        self._set_graph_everywhere(build_word_smoke_graph())
        self.current_path = None
        self.dirty = True
        QSettings().setValue("workspace/dirty", True)
        self._autosave_graph()
        self._update_title()
        self._log("Створено тестовий Word-граф. Натисніть F5 і виберіть DOCX.")

    def open_order_senders_preset(self) -> None:
        if not self._confirm_discard():
            return
        self._set_graph_everywhere(build_order_senders_graph())
        self.current_path = None
        self.dirty = True
        QSettings().setValue("workspace/dirty", True)
        self._autosave_graph()
        self._update_title()
        self._log("Завантажено сценарій: Аналіз відправників та відповідних пунктів наказу. Натисніть F5 для запуску.")

    def open_graph(self) -> None:
        if not self._confirm_discard():
            return
        filename, _ = QFileDialog.getOpenFileName(
            self,
            "Відкрити сценарій",
            str(self.graphs_dir),
            "Node Automation Toolkit (*.nat.json);;JSON (*.json)",
        )
        if not filename:
            return
        try:
            graph = load_graph(Path(filename))
            missing = [node.type_id for node in graph.nodes if not self._has_node(node.type_id)]
            if missing:
                raise ValueError("Не встановлені ноди:\n" + "\n".join(sorted(set(missing))))
            self._set_graph_everywhere(graph)
            self.current_path = Path(filename)
            self.dirty = False
            self._remember_current_path()
            QSettings().setValue("workspace/dirty", False)
            self._autosave_graph()
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
            self._remember_current_path()
            QSettings().setValue("workspace/dirty", False)
            self._autosave_graph()
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
            str(self.current_path or self.graphs_dir / "scenario.nat.json"),
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
            if self.editor_tabs.currentWidget() is self.blueprint:
                self.blueprint.run_graph()
                return
            graph = self._current_graph()
            names = {node.id: self.registry.get(node.type_id).name for node in graph.nodes}
            executor = GraphExecutor(self.registry)
            result = executor.execute(
                graph,
                on_node_started=lambda node_id: self._log(f"{names[node_id]}: виконується"),
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
        self.blueprint.reload_definitions()
        self.palette.populate(self.registry, self.search.text())
        self._report_plugin_errors()
        self._log(f"Завантажено нод: {len(self.registry.all())}")

    def _current_graph(self) -> GraphModel:
        if self.editor_tabs.currentWidget() is self.blueprint:
            return self.blueprint.graph_model()
        return self.scene.graph

    def _set_graph_everywhere(self, graph: GraphModel) -> None:
        previous = self._suppress_graph_changes
        self._suppress_graph_changes = True
        try:
            self.scene.set_graph(graph)
            self.blueprint.set_graph_model(graph)
        finally:
            self._suppress_graph_changes = previous

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
        if self._suppress_graph_changes:
            return
        self.dirty = True
        QSettings().setValue("workspace/dirty", True)
        self.autosave_timer.start()
        self._update_title()

    def _update_title(self) -> None:
        name = self.current_path.name if self.current_path else self._current_graph().name
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
        if answer == QMessageBox.StandardButton.Discard:
            self._discard_autosave_changes()
            return True
        return False

    def _autosave_graph(self) -> None:
        try:
            save_graph(self._current_graph(), self.autosave_path)
            self._remember_current_path()
        except Exception as error:  # noqa: BLE001 - recovery boundary
            self._log(f"ПОМИЛКА АВТОЗБЕРЕЖЕННЯ: {error}")

    def _restore_working_graph(self) -> None:
        settings = QSettings()
        if not self.autosave_path.is_file():
            return
        try:
            graph = load_graph(self.autosave_path)
            missing = [node.type_id for node in graph.nodes if not self._has_node(node.type_id)]
            if missing:
                raise ValueError("Не встановлені ноди: " + ", ".join(sorted(set(missing))))
            self._set_graph_everywhere(graph)
            stored_path = settings.value("workspace/current_path", "", type=str)
            self.current_path = Path(stored_path) if stored_path else None
            self.dirty = settings.value("workspace/dirty", False, type=bool)
            self._log(f"Відновлено робочий граф: {graph.name}")
        except Exception as error:  # noqa: BLE001 - recovery boundary
            self._log(f"Не вдалося відновити автозбереження: {error}")

    def _remember_current_path(self) -> None:
        QSettings().setValue(
            "workspace/current_path",
            str(self.current_path) if self.current_path is not None else "",
        )

    def _discard_autosave_changes(self) -> None:
        try:
            if self.current_path is not None and self.current_path.is_file():
                save_graph(load_graph(self.current_path), self.autosave_path)
            elif self.autosave_path.exists():
                self.autosave_path.unlink()
        finally:
            QSettings().setValue("workspace/dirty", False)

    def _restore_window_state(self) -> None:
        settings = QSettings()
        if geometry := settings.value("window/geometry"):
            self.restoreGeometry(geometry)

    def closeEvent(self, event) -> None:
        if not self._confirm_discard():
            event.ignore()
            return
        QSettings().setValue("window/geometry", self.saveGeometry())
        if not self.dirty:
            self._autosave_graph()
        event.accept()
