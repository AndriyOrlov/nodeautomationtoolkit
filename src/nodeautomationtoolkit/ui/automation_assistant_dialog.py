from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QObject, QThread, QUrl, Signal, Slot
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from nodeautomationtoolkit.core.automation_assistant import (
    AutomationAssistant,
    AutomationPlan,
)
from nodeautomationtoolkit.core.embedded_llm import MODEL_ALIAS, SERVER_API_KEY
from nodeautomationtoolkit.core.local_llm import (
    DEFAULT_BASE_URLS,
    PROVIDER_API_KEY_URLS,
    PROVIDER_PRESET_MODELS,
    LocalLlmClient,
    LocalLlmConfig,
    LocalLlmProvider,
    load_llm_settings,
    save_llm_settings,
)
from nodeautomationtoolkit.core.models import GraphModel
from nodeautomationtoolkit.core.registry import NodeRegistry

from .node_assistant_dialog import NodeAssistantDialog


class PlanWorker(QObject):
    finished = Signal(object)
    failed = Signal(str)

    def __init__(
        self,
        config: LocalLlmConfig,
        registry: NodeRegistry,
        prompt: str,
        graph: GraphModel,
    ) -> None:
        super().__init__()
        self.config = config
        self.registry = registry
        self.prompt = prompt
        self.graph = graph

    @Slot()
    def run(self) -> None:
        try:
            assistant = AutomationAssistant(LocalLlmClient(self.config), self.registry)
            self.finished.emit(assistant.create_plan(self.prompt, self.graph))
        except Exception as error:  # noqa: BLE001 - worker boundary
            self.failed.emit(str(error))


class AutomationAssistantDialog(QDialog):
    graph_created = Signal(object)

    def __init__(
        self,
        registry: NodeRegistry,
        graph: GraphModel,
        plugin_dir: Path,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.registry = registry
        self.graph = graph
        self.plugin_dir = plugin_dir
        self.plan: AutomationPlan | None = None
        self.thread: QThread | None = None
        self.worker: PlanWorker | None = None
        self.setWindowTitle("AI-помічник сценаріїв")
        self.resize(850, 700)

        saved = load_llm_settings()

        self.provider = QComboBox()
        self.provider.addItems([item.value for item in LocalLlmProvider])
        self.provider.setCurrentText(saved["provider"])

        self.base_url = QLineEdit(saved["base_url"])
        
        self.model = QComboBox()
        self.model.setEditable(True)

        self.refresh_models_btn = QPushButton("🔄 Оновити")
        self.refresh_models_btn.setToolTip("Запитати точний список доступних моделей у сервера провайдера")
        self.refresh_models_btn.clicked.connect(self._refresh_models_from_server)

        model_widget = QWidget()
        model_layout = QHBoxLayout(model_widget)
        model_layout.setContentsMargins(0, 0, 0, 0)
        model_layout.addWidget(self.model, 1)
        model_layout.addWidget(self.refresh_models_btn)

        self.api_key = QLineEdit(saved["api_key"])
        self.api_key.setEchoMode(QLineEdit.EchoMode.Password)
        self.api_key.setPlaceholderText("Зберігається локально")

        self.api_key_btn = QPushButton("🔑 Отримати API-ключ")
        self.api_key_btn.setToolTip("Відкрити офіційний сайт для отримання API-ключа у браузері")
        self.api_key_btn.clicked.connect(self._open_api_key_url)

        api_key_widget = QWidget()
        api_key_layout = QHBoxLayout(api_key_widget)
        api_key_layout.setContentsMargins(0, 0, 0, 0)
        api_key_layout.addWidget(self.api_key)
        api_key_layout.addWidget(self.api_key_btn)

        self.provider.currentTextChanged.connect(self._provider_changed)
        self.api_key.textChanged.connect(self._save_settings)
        self.base_url.textChanged.connect(self._save_settings)
        self.model.currentTextChanged.connect(self._save_settings)

        self._provider_changed(saved["provider"])

        form = QFormLayout()
        form.addRow("Провайдер", self.provider)
        form.addRow("Адреса API", self.base_url)
        form.addRow("Вибір моделі", model_widget)
        form.addRow("API-ключ", api_key_widget)

        self.prompt = QPlainTextEdit()
        self.prompt.setPlaceholderText(
            "Наприклад: Візьми всі DOCX із папки, отримай назви файлів, "
            "залиши ті, у назві яких є слово Наказ."
        )
        self.create_button = QPushButton("Підготувати план")
        self.create_button.clicked.connect(self.create_plan)
        self.preview = QPlainTextEdit()
        self.preview.setReadOnly(True)
        self.preview.setPlaceholderText("Тут з'являться заплановані зміни графа")
        self.create_missing_button = QPushButton("Створити відсутню Python-ноду…")
        self.create_missing_button.setEnabled(False)
        self.create_missing_button.clicked.connect(self.create_missing_node)

        self.buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        self.apply_button = self.buttons.addButton(
            "Застосувати до графа", QDialogButtonBox.ButtonRole.AcceptRole
        )
        self.apply_button.setEnabled(False)
        self.apply_button.clicked.connect(self.apply_plan)
        self.buttons.rejected.connect(self.reject)

        layout = QVBoxLayout(self)
        layout.addWidget(
            QLabel(
                "Помічник бачить тільки каталог нод і ваш опис. Вміст документів "
                "та значення параметрів до запиту не додаються. API не отримує інструментів."
            )
        )
        layout.addLayout(form)
        layout.addWidget(QLabel("Опишіть потрібну автоматизацію"))
        layout.addWidget(self.prompt, 1)
        layout.addWidget(self.create_button)
        layout.addWidget(self.create_missing_button)
        layout.addWidget(QLabel("План змін"))
        layout.addWidget(self.preview, 1)
        layout.addWidget(self.buttons)

    def _provider_changed(self, value: str) -> None:
        try:
            provider = LocalLlmProvider(value)
        except ValueError:
            provider = LocalLlmProvider.EMBEDDED

        saved = load_llm_settings(provider)
        self.base_url.setText(saved["base_url"])
        embedded = provider == LocalLlmProvider.EMBEDDED
        self.base_url.setEnabled(not embedded)
        self.model.setEnabled(True)

        self.model.clear()
        presets = PROVIDER_PRESET_MODELS.get(provider, [])
        if presets:
            self.model.addItems(presets)
        
        if saved["model"]:
            self.model.setCurrentText(saved["model"])
        elif presets:
            self.model.setCurrentIndex(0)

        url = PROVIDER_API_KEY_URLS.get(provider, "")
        self.api_key_btn.setVisible(bool(url))
        if provider == LocalLlmProvider.GEMINI:
            self.api_key_btn.setText("🔑 Отримати Google API Key")
        elif provider == LocalLlmProvider.OPENAI:
            self.api_key_btn.setText("🔑 Отримати OpenAI Key")
        elif provider == LocalLlmProvider.LM_STUDIO:
            self.api_key_btn.setText("🌐 Сайт LM Studio")
        elif provider == LocalLlmProvider.OLLAMA:
            self.api_key_btn.setText("🌐 Сайт Ollama")

        if embedded:
            self.api_key.setText(SERVER_API_KEY)
            self.api_key.setEnabled(False)
        else:
            self.api_key.setText(saved["api_key"])
            self.api_key.setEnabled(True)

        self._save_settings()

    def _refresh_models_from_server(self) -> None:
        self.refresh_models_btn.setEnabled(False)
        self.refresh_models_btn.setText("🔄 Завантаження…")
        try:
            client = LocalLlmClient(self._config())
            models = client.fetch_available_models()
            current = self.model.currentText().strip()
            self.model.clear()
            if models:
                self.model.addItems(models)
                if current in models:
                    self.model.setCurrentText(current)
                else:
                    self.model.setCurrentIndex(0)
        finally:
            self.refresh_models_btn.setEnabled(True)
            self.refresh_models_btn.setText("🔄 Оновити")

    def _save_settings(self) -> None:
        save_llm_settings(
            provider_value=self.provider.currentText(),
            base_url=self.base_url.text().strip(),
            model=self.model.currentText().strip(),
            api_key=self.api_key.text().strip(),
        )

    def _open_api_key_url(self) -> None:
        provider = LocalLlmProvider(self.provider.currentText())
        url_str = PROVIDER_API_KEY_URLS.get(provider, "")
        if url_str:
            QDesktopServices.openUrl(QUrl(url_str))

    def _config(self) -> LocalLlmConfig:
        provider = LocalLlmProvider(self.provider.currentText())
        model_name = self.model.currentText().strip()
        api_key = self.api_key.text().strip()
        if provider == LocalLlmProvider.EMBEDDED and not api_key:
            api_key = SERVER_API_KEY
        return LocalLlmConfig(
            provider=provider,
            base_url=self.base_url.text().strip(),
            model=model_name,
            api_key=api_key or "local",
        )

    def create_plan(self) -> None:
        model_name = self.model.currentText().strip()
        if not model_name or not self.prompt.toPlainText().strip():
            QMessageBox.information(
                self, "Недостатньо даних", "Вкажіть модель і опишіть автоматизацію"
            )
            return
        self.create_button.setEnabled(False)
        self.create_button.setText("Планування…")
        self.thread = QThread(self)
        self.worker = PlanWorker(
            self._config(),
            self.registry,
            self.prompt.toPlainText(),
            self.graph,
        )
        self.worker.moveToThread(self.thread)
        self.thread.started.connect(self.worker.run)
        self.worker.finished.connect(self._plan_ready)
        self.worker.failed.connect(self._plan_failed)
        self.worker.finished.connect(self.thread.quit)
        self.worker.failed.connect(self.thread.quit)
        self.thread.finished.connect(self._thread_finished)
        self.thread.start()

    @Slot(object)
    def _plan_ready(self, plan: AutomationPlan) -> None:
        self.plan = plan
        assistant = AutomationAssistant(LocalLlmClient(self._config()), self.registry)
        lines = [plan.title, plan.summary, ""] + assistant.preview(plan)
        self.preview.setPlainText("\n".join(line for line in lines if line is not None))
        self.create_missing_button.setEnabled(bool(plan.missing_nodes))
        if plan.missing_nodes:
            self.apply_button.setEnabled(False)
            return
        try:
            assistant.apply_plan(self.graph, plan)
            self.apply_button.setEnabled(True)
        except Exception as error:  # noqa: BLE001 - plan boundary
            self.preview.setPlainText(f"План не пройшов перевірку:\n{error}")
            self.apply_button.setEnabled(False)

    def create_missing_node(self) -> None:
        if self.plan is None or not self.plan.missing_nodes:
            return
        missing = self.plan.missing_nodes[0]
        prompt = missing.suggested_prompt or (
            f"Створи чисту Python-ноду '{missing.name}'.\n"
            f"Призначення: {missing.description}\n"
            f"Входи: {missing.inputs}\nВиходи: {missing.outputs}\n"
            "Нода перетворює тільки передані значення. Не читає і не записує файли."
        )
        dialog = NodeAssistantDialog(
            self.plugin_dir,
            self,
            initial_prompt=prompt,
            initial_config=self._config(),
            strict_no_filesystem=True,
        )
        dialog.node_installed.connect(self._missing_node_installed)
        dialog.exec()

    @Slot(str)
    def _missing_node_installed(self, path: str) -> None:
        self.registry.reload(self.plugin_dir)
        self.preview.appendPlainText(f"\nНоду встановлено: {Path(path).name}. Переплановую…")
        self.create_plan()

    @Slot(str)
    def _plan_failed(self, message: str) -> None:
        QMessageBox.critical(self, "Не вдалося створити план", message)

    @Slot()
    def _thread_finished(self) -> None:
        self.create_button.setEnabled(True)
        self.create_button.setText("Підготувати план")
        self.worker = None
        if self.thread is not None:
            self.thread.deleteLater()
        self.thread = None

    def apply_plan(self) -> None:
        if self.plan is None:
            return
        try:
            result = AutomationAssistant(LocalLlmClient(self._config()), self.registry).apply_plan(
                self.graph, self.plan
            )
        except Exception as error:  # noqa: BLE001 - UI boundary
            QMessageBox.critical(self, "План не застосовано", str(error))
            return
        self.graph_created.emit(result)
        self.accept()
