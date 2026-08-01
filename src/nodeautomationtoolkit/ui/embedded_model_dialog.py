from __future__ import annotations

from PySide6.QtCore import QObject, QThread, Signal, Slot
from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QLabel,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QVBoxLayout,
)

from nodeautomationtoolkit.core.embedded_llm import (
    MODEL_NAME,
    MODEL_SIZE,
    EmbeddedLlmInstaller,
)


class ModelInstallWorker(QObject):
    progress = Signal(str, int)
    finished = Signal()
    failed = Signal(str)

    @Slot()
    def run(self) -> None:
        try:
            installer = EmbeddedLlmInstaller()

            def report(label: str, current: int, total: int) -> None:
                percent = int(current * 100 / total) if total else 0
                self.progress.emit(label, max(0, min(percent, 100)))

            installer.install(report)
            self.finished.emit()
        except Exception as error:  # noqa: BLE001 - installer boundary
            self.failed.emit(str(error))


class EmbeddedModelDialog(QDialog):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.thread: QThread | None = None
        self.worker: ModelInstallWorker | None = None
        self.setWindowTitle("Локальна модель")
        self.resize(520, 260)

        self.info = QLabel()
        self.info.setWordWrap(True)
        self.status = QLabel()
        self.status.setWordWrap(True)
        self.progress_label = QLabel("")
        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.install_button = QPushButton("Встановити локальну модель")
        self.install_button.clicked.connect(self.install)
        self.buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        self.buttons.rejected.connect(self.reject)

        layout = QVBoxLayout(self)
        layout.addWidget(self.info)
        layout.addWidget(self.status)
        layout.addWidget(self.progress_label)
        layout.addWidget(self.progress)
        layout.addWidget(self.install_button)
        layout.addStretch(1)
        layout.addWidget(self.buttons)
        self._refresh()

    def _refresh(self) -> None:
        size_gib = MODEL_SIZE / 1024**3
        self.info.setText(
            f"{MODEL_NAME} · приблизно {size_gib:.1f} ГБ · Apache 2.0.\n"
            "Після встановлення працює офлайн через вбудований llama.cpp/Vulkan. "
            "Документи не передаються в інтернет."
        )
        current = EmbeddedLlmInstaller().status()
        if current.ready:
            self.status.setText("✓ Модель і локальний рушій установлені")
            self.install_button.setText("Перевірити встановлення")
        else:
            missing = []
            if not current.runtime_installed:
                missing.append("рушій")
            if not current.model_installed:
                missing.append("модель")
            self.status.setText("Не встановлено: " + ", ".join(missing))

    def install(self) -> None:
        if EmbeddedLlmInstaller().status().ready:
            QMessageBox.information(self, "Локальна модель", "Модель готова до роботи")
            return
        self.install_button.setEnabled(False)
        self.thread = QThread(self)
        self.worker = ModelInstallWorker()
        self.worker.moveToThread(self.thread)
        self.thread.started.connect(self.worker.run)
        self.worker.progress.connect(self._on_progress)
        self.worker.finished.connect(self._installed)
        self.worker.failed.connect(self._failed)
        self.worker.finished.connect(self.thread.quit)
        self.worker.failed.connect(self.thread.quit)
        self.thread.finished.connect(self._thread_finished)
        self.thread.start()

    @Slot(str, int)
    def _on_progress(self, label: str, percent: int) -> None:
        self.progress_label.setText(label)
        self.progress.setValue(percent)

    @Slot()
    def _installed(self) -> None:
        self.progress.setValue(100)
        self.progress_label.setText("Установлення завершено")
        self._refresh()
        QMessageBox.information(self, "Локальна модель", "Модель установлена і готова")

    @Slot(str)
    def _failed(self, message: str) -> None:
        QMessageBox.critical(self, "Не вдалося встановити модель", message)

    @Slot()
    def _thread_finished(self) -> None:
        self.install_button.setEnabled(True)
        self.worker = None
        if self.thread is not None:
            self.thread.deleteLater()
        self.thread = None
