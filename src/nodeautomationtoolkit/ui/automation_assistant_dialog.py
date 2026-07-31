from __future__ import annotations

from PySide6.QtCore import QObject, QThread, Signal, Slot
from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QVBoxLayout,
)

from nodeautomationtoolkit.core.automation_assistant import (
    AutomationAssistant,
    AutomationPlan,
)
from nodeautomationtoolkit.core.local_llm import (
    DEFAULT_BASE_URLS,
    LocalLlmClient,
    LocalLlmConfig,
    LocalLlmProvider,
)
from nodeautomationtoolkit.core.models import GraphModel
from nodeautomationtoolkit.core.registry import NodeRegistry


class PlanWorker(QObject):
    finished = Signal(object)
    failed = Signal(str)

    def __init__(
        self,
        config: LocalLlmConfig,
        registry: NodeRegistry,
        prompt: str,
    ) -> None:
        super().__init__()
        self.config = config
        self.registry = registry
        self.prompt = prompt

    @Slot()
    def run(self) -> None:
        try:
            assistant = AutomationAssistant(LocalLlmClient(self.config), self.registry)
            self.finished.emit(assistant.create_plan(self.prompt))
        except Exception as error:  # noqa: BLE001 - worker boundary
            self.failed.emit(str(error))


class AutomationAssistantDialog(QDialog):
    graph_created = Signal(object)

    def __init__(self, registry: NodeRegistry, graph: GraphModel, parent=None) -> None:
        super().__init__(parent)
        self.registry = registry
        self.graph = graph
        self.plan: AutomationPlan | None = None
        self.thread: QThread | None = None
        self.worker: PlanWorker | None = None
        self.setWindowTitle("Створити автоматизацію локальною LLM")
        self.resize(800, 680)

        self.provider = QComboBox()
        self.provider.addItems([item.value for item in LocalLlmProvider])
        self.provider.currentTextChanged.connect(self._provider_changed)
        self.base_url = QLineEdit(DEFAULT_BASE_URLS[LocalLlmProvider.LM_STUDIO])
        self.model = QLineEdit()
        self.model.setPlaceholderText("Назва завантаженої локальної моделі")

        form = QFormLayout()
        form.addRow("Сервер", self.provider)
        form.addRow("Локальна адреса", self.base_url)
        form.addRow("Модель", self.model)

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
                "до запиту не додається."
            )
        )
        layout.addLayout(form)
        layout.addWidget(QLabel("Опишіть потрібну автоматизацію"))
        layout.addWidget(self.prompt, 1)
        layout.addWidget(self.create_button)
        layout.addWidget(QLabel("План змін"))
        layout.addWidget(self.preview, 1)
        layout.addWidget(self.buttons)

    def _provider_changed(self, value: str) -> None:
        self.base_url.setText(DEFAULT_BASE_URLS[LocalLlmProvider(value)])

    def _config(self) -> LocalLlmConfig:
        return LocalLlmConfig(
            provider=LocalLlmProvider(self.provider.currentText()),
            base_url=self.base_url.text().strip(),
            model=self.model.text().strip(),
        )

    def create_plan(self) -> None:
        if not self.model.text().strip() or not self.prompt.toPlainText().strip():
            QMessageBox.information(
                self, "Недостатньо даних", "Вкажіть модель і опишіть автоматизацію"
            )
            return
        self.create_button.setEnabled(False)
        self.create_button.setText("Планування…")
        self.thread = QThread(self)
        self.worker = PlanWorker(self._config(), self.registry, self.prompt.toPlainText())
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
        try:
            assistant.apply_plan(self.graph, plan)
            lines = [plan.title, plan.summary, ""] + assistant.preview(plan)
            self.preview.setPlainText("\n".join(line for line in lines if line is not None))
            self.apply_button.setEnabled(True)
        except Exception as error:  # noqa: BLE001 - plan boundary
            self.preview.setPlainText(f"План не пройшов перевірку:\n{error}")
            self.apply_button.setEnabled(False)

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
            result = AutomationAssistant(
                LocalLlmClient(self._config()), self.registry
            ).apply_plan(self.graph, self.plan)
        except Exception as error:  # noqa: BLE001 - UI boundary
            QMessageBox.critical(self, "План не застосовано", str(error))
            return
        self.graph_created.emit(result)
        self.accept()
