from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QObject, QThread, Signal, Slot
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QSplitter,
    QVBoxLayout,
    QWidget,
)

from nodeautomationtoolkit.core.local_llm import (
    DEFAULT_BASE_URLS,
    LocalLlmClient,
    LocalLlmConfig,
    LocalLlmProvider,
)
from nodeautomationtoolkit.core.node_draft import (
    NodeDraft,
    install_node_draft,
    review_node_code,
)


class GenerateWorker(QObject):
    finished = Signal(object)
    failed = Signal(str)

    def __init__(self, config: LocalLlmConfig, prompt: str) -> None:
        super().__init__()
        self.config = config
        self.prompt = prompt

    @Slot()
    def run(self) -> None:
        try:
            self.finished.emit(LocalLlmClient(self.config).generate_node(self.prompt))
        except Exception as error:  # noqa: BLE001 - background boundary
            self.failed.emit(str(error))


class NodeAssistantDialog(QDialog):
    node_installed = Signal(str)

    def __init__(self, plugin_dir: Path, parent=None) -> None:
        super().__init__(parent)
        self.plugin_dir = plugin_dir
        self.draft: NodeDraft | None = None
        self.thread: QThread | None = None
        self.worker: GenerateWorker | None = None
        self.setWindowTitle("Створити ноду локальною LLM")
        self.resize(1050, 720)

        self.provider = QComboBox()
        self.provider.addItems([item.value for item in LocalLlmProvider])
        self.base_url = QLineEdit(DEFAULT_BASE_URLS[LocalLlmProvider.LM_STUDIO])
        self.model = QLineEdit()
        self.model.setPlaceholderText("Назва завантаженої локальної моделі")
        self.provider.currentTextChanged.connect(self._provider_changed)

        settings = QFormLayout()
        settings.addRow("Сервер", self.provider)
        settings.addRow("Локальна адреса", self.base_url)
        settings.addRow("Модель", self.model)

        self.prompt = QPlainTextEdit()
        self.prompt.setPlaceholderText(
            "Наприклад: Створи ноду, яка приймає список рядків і залишає лише рядки "
            "з указаною фразою без урахування регістру."
        )
        self.generate_button = QPushButton("Створити чернетку")
        self.generate_button.clicked.connect(self.generate)

        request_panel = QWidget()
        request_layout = QVBoxLayout(request_panel)
        request_layout.addLayout(settings)
        request_layout.addWidget(QLabel("Опишіть потрібну ноду"))
        request_layout.addWidget(self.prompt, 1)
        request_layout.addWidget(self.generate_button)

        self.code = QPlainTextEdit()
        self.code.setLineWrapMode(QPlainTextEdit.LineWrapMode.NoWrap)
        self.code.textChanged.connect(self.review)
        self.review_label = QLabel("Спочатку створіть чернетку")
        self.review_label.setWordWrap(True)
        self.approval = QCheckBox("Я переглянув код і дозволяю встановити цю ноду")
        self.approval.toggled.connect(self._update_install_state)

        code_panel = QWidget()
        code_layout = QVBoxLayout(code_panel)
        code_layout.addWidget(QLabel("Код чернетки"))
        code_layout.addWidget(self.code, 1)
        code_layout.addWidget(self.review_label)
        code_layout.addWidget(self.approval)

        splitter = QSplitter()
        splitter.addWidget(request_panel)
        splitter.addWidget(code_panel)
        splitter.setSizes([420, 630])

        self.buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        self.install_button = self.buttons.addButton(
            "Встановити ноду", QDialogButtonBox.ButtonRole.AcceptRole
        )
        self.install_button.setEnabled(False)
        self.install_button.clicked.connect(self.install)
        self.buttons.rejected.connect(self.reject)

        layout = QVBoxLayout(self)
        layout.addWidget(
            QLabel(
                "Запит надсилається лише на вказаний localhost. Документи та результати "
                "сценарію не додаються до запиту."
            )
        )
        layout.addWidget(splitter, 1)
        layout.addWidget(self.buttons)

    def _provider_changed(self, value: str) -> None:
        provider = LocalLlmProvider(value)
        self.base_url.setText(DEFAULT_BASE_URLS[provider])

    def _config(self) -> LocalLlmConfig:
        return LocalLlmConfig(
            provider=LocalLlmProvider(self.provider.currentText()),
            base_url=self.base_url.text().strip(),
            model=self.model.text().strip(),
        )

    def generate(self) -> None:
        if not self.prompt.toPlainText().strip():
            QMessageBox.information(self, "Потрібен опис", "Опишіть потрібну ноду")
            return
        if not self.model.text().strip():
            QMessageBox.information(self, "Потрібна модель", "Вкажіть локальну модель")
            return
        self.generate_button.setEnabled(False)
        self.generate_button.setText("Генерування…")
        self.thread = QThread(self)
        self.worker = GenerateWorker(self._config(), self.prompt.toPlainText())
        self.worker.moveToThread(self.thread)
        self.thread.started.connect(self.worker.run)
        self.worker.finished.connect(self._generation_finished)
        self.worker.failed.connect(self._generation_failed)
        self.worker.finished.connect(self.thread.quit)
        self.worker.failed.connect(self.thread.quit)
        self.thread.finished.connect(self._thread_finished)
        self.thread.start()

    @Slot(object)
    def _generation_finished(self, draft: NodeDraft) -> None:
        self.draft = draft
        self.code.setPlainText(draft.code)
        self.approval.setChecked(False)
        self.review()

    @Slot(str)
    def _generation_failed(self, message: str) -> None:
        QMessageBox.critical(self, "Не вдалося створити ноду", message)

    @Slot()
    def _thread_finished(self) -> None:
        self.generate_button.setEnabled(True)
        self.generate_button.setText("Створити чернетку")
        self.worker = None
        if self.thread is not None:
            self.thread.deleteLater()
        self.thread = None

    def review(self) -> None:
        code = self.code.toPlainText()
        if not code.strip():
            self.review_label.setText("Код відсутній")
            self.install_button.setEnabled(False)
            return
        review = review_node_code(code)
        lines = []
        if review.errors:
            lines.extend(f"ПОМИЛКА: {item}" for item in review.errors)
        else:
            lines.append("Статична перевірка пройдена")
        lines.extend(f"УВАГА: {item}" for item in review.warnings)
        if review.permissions:
            lines.append("Дозволи: " + ", ".join(sorted(review.permissions)))
        else:
            lines.append("Додаткові дозволи не виявлені")
        self.review_label.setText("\n".join(lines))
        self._update_install_state()

    def _update_install_state(self) -> None:
        review = review_node_code(self.code.toPlainText())
        self.install_button.setEnabled(review.installable and self.approval.isChecked())

    def install(self) -> None:
        if self.draft is None:
            return
        updated = self.draft.model_copy(update={"code": self.code.toPlainText()})
        try:
            target = install_node_draft(updated, self.plugin_dir)
        except Exception as error:  # noqa: BLE001 - UI boundary
            QMessageBox.critical(self, "Ноду не встановлено", str(error))
            return
        self.node_installed.emit(str(target))
        QMessageBox.information(
            self,
            "Ноду встановлено",
            f"Створено {target.name}. Оновіть плагіни, щоб додати її до палітри.",
        )
        self.accept()
