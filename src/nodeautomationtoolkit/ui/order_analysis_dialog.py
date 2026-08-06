"""
Діалог «Аналіз наказу AI» — режим «до / після».

Ліворуч  — оригінальний текст наказу (завантажений файл).
Праворуч — результат AI-обробки (структурований аналіз / витяги).
Знизу    — автоматично побудоване дерево нод (сценарій автоматизації).
"""
from __future__ import annotations

import json
import uuid
from pathlib import Path
from typing import Any

from PySide6.QtCore import (
    QObject,
    QRectF,
    QSizeF,
    Qt,
    QThread,
    Signal,
    Slot,
)
from PySide6.QtGui import (
    QBrush,
    QColor,
    QFont,
    QLinearGradient,
    QPainter,
    QPen,
    QSyntaxHighlighter,
    QTextCharFormat,
)
from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QFileDialog,
    QFrame,
    QGraphicsEllipseItem,
    QGraphicsItem,
    QGraphicsLineItem,
    QGraphicsRectItem,
    QGraphicsScene,
    QGraphicsTextItem,
    QGraphicsView,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QSizePolicy,
    QSplitter,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from nodeautomationtoolkit.core.local_llm import (
    DEFAULT_BASE_URLS,
    PROVIDER_PRESET_MODELS,
    LocalLlmClient,
    LocalLlmConfig,
    LocalLlmProvider,
    load_llm_settings,
    save_llm_settings,
)
from nodeautomationtoolkit.core.models import ConnectionModel, GraphModel, NodeModel
from nodeautomationtoolkit.core.registry import NodeRegistry


# ─────────────────────────────────────────────────────────────────────────────
# Системні промпти
# ─────────────────────────────────────────────────────────────────────────────
_SYS_ANALYZE = """Ти — спеціалізований AI-аналітик військових наказів та документів ЗСУ.
Проаналізуй наказ і поверни структурований Markdown-звіт ВИКЛЮЧНО за шаблоном:

### 📋 Метадані наказу
- Номер: …
- Дата: …
- Місце: …
- Вид: …

### 🏛 Виявлені військові частини
| Відкрита назва | Шифр (якщо відомий) | Кількість згадок |
|---|---|---|

### 📑 Структура пунктів
1. Пункт 1 — короткий зміст
2. …

### 👤 Особовий склад
| ПІБ | Звання | Дія | Підстава |
|---|---|---|---|

### ⚙️ Рекомендований сценарій
Коротко: які ноди потрібні для автоматизації цього наказу.

Відповідай ВИКЛЮЧНО українською мовою. Будь точним і лаконічним."""

_SYS_GRAPH = """Ти — AI-планувальник сценаріїв Node Automation Toolkit.
Отримуєш аналіз наказу і будуєш граф автоматизації.
Відповідай ВИКЛЮЧНО валідним JSON (без markdown, без коментарів):

{
  "title": "Назва",
  "summary": "Опис",
  "nodes": [
    {"id": "n1", "type_id": "builtin.flow.start", "label": "Старт", "x": 0, "y": 0, "params": {}},
    {"id": "n2", "type_id": "builtin.word.read_docx", "label": "Читати DOCX", "x": 220, "y": 0, "params": {"path": ""}},
    {"id": "n3", "type_id": "builtin.order.map_military_units", "label": "Знайти ВЧ", "x": 440, "y": 0, "params": {}},
    {"id": "n4", "type_id": "builtin.output.show_result", "label": "Результат", "x": 660, "y": 0, "params": {}}
  ],
  "edges": [
    {"from": "n1", "from_port": "then", "to": "n2", "to_port": "exec", "kind": "execution"},
    {"from": "n2", "from_port": "document", "to": "n3", "to_port": "text", "kind": "data"}
  ]
}

Доступні type_id (використовуй лише їх):
- builtin.flow.start · builtin.flow.sub_start · builtin.flow.branch · builtin.flow.sequence
- builtin.word.read_docx · builtin.word.replace_in_docx · builtin.word.merge_docx · builtin.word.word_count
- builtin.order.read_recipient_mapping · builtin.order.map_military_units · builtin.order.groups_to_ciphers
- builtin.order_batch.create_unit_extracts
- builtin.excel.read_sheet · builtin.excel.save_table · builtin.excel.write_cell
- builtin.files.list_files · builtin.files.move_file · builtin.files.create_folder
- builtin.text.replace · builtin.text.fill_template · builtin.text.today_date · builtin.text.clean
- builtin.output.show_result · builtin.output.show_table

ВАЖЛИВО: x координати — кратні 220, y — кратні 120 (для кожного паралельного потоку +120)."""


# ─────────────────────────────────────────────────────────────────────────────
# Кольорова схема типів нод
# ─────────────────────────────────────────────────────────────────────────────
NODE_COLORS: dict[str, str] = {
    "builtin.flow": "#6366f1",       # Потік — фіолетовий
    "builtin.word": "#0284c7",       # Word — синій
    "builtin.order": "#d97706",      # Наказ — помаранчевий
    "builtin.order_batch": "#b45309",
    "builtin.excel": "#15803d",      # Excel — зелений
    "builtin.files": "#64748b",      # Файли — сірий
    "builtin.text": "#7c3aed",       # Текст — пурпурний
    "builtin.output": "#dc2626",     # Вивід — червоний
}


def _node_color(type_id: str) -> str:
    for prefix, color in NODE_COLORS.items():
        if type_id.startswith(prefix):
            return color
    return "#334155"


# ─────────────────────────────────────────────────────────────────────────────
# Графічна нода для дерева нод
# ─────────────────────────────────────────────────────────────────────────────
NODE_W = 190
NODE_H = 56


class GraphNodeItem(QGraphicsRectItem):
    """Візуальна нода у дереві нод."""

    def __init__(self, node_data: dict) -> None:
        super().__init__(0, 0, NODE_W, NODE_H)
        self.node_data = node_data
        color = QColor(_node_color(node_data.get("type_id", "")))

        # Градієнт
        grad = QLinearGradient(0, 0, 0, NODE_H)
        grad.setColorAt(0, color.lighter(130))
        grad.setColorAt(1, color)
        self.setBrush(QBrush(grad))
        self.setPen(QPen(color.darker(140), 1.5))
        self.setPos(node_data.get("x", 0), node_data.get("y", 0))
        self.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsMovable, True)
        self.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsSelectable, True)

        # Підпис
        label = node_data.get("label", node_data.get("type_id", "").split(".")[-1])
        text = QGraphicsTextItem(label, self)
        text.setDefaultTextColor(QColor("white"))
        font = QFont("Segoe UI", 9, QFont.Weight.Bold)
        text.setFont(font)
        text.setTextWidth(NODE_W - 12)
        text.setPos(6, 6)

        # type_id рядок
        tid = QGraphicsTextItem(node_data.get("type_id", ""), self)
        tid.setDefaultTextColor(QColor(255, 255, 255, 140))
        small = QFont("Segoe UI", 7)
        tid.setFont(small)
        tid.setTextWidth(NODE_W - 12)
        tid.setPos(6, 30)

    def port_pos(self, side: str) -> tuple[float, float]:
        """Позиція входу або виходу ноди у координатах сцени."""
        sx, sy = self.x(), self.y()
        if side == "out":
            return sx + NODE_W, sy + NODE_H / 2
        return sx, sy + NODE_H / 2


class EdgeItem(QGraphicsLineItem):
    """З'єднання між двома нодами."""

    def __init__(self, src: GraphNodeItem, dst: GraphNodeItem, kind: str = "data") -> None:
        sx, sy = src.port_pos("out")
        dx, dy = dst.port_pos("in")
        super().__init__(sx, sy, dx, dy)
        color = QColor("#6366f1") if kind == "execution" else QColor("#94a3b8")
        pen = QPen(color, 2 if kind == "execution" else 1.5, Qt.PenStyle.SolidLine)
        self.setPen(pen)
        self.setZValue(-1)


# ─────────────────────────────────────────────────────────────────────────────
# Панель дерева нод
# ─────────────────────────────────────────────────────────────────────────────
class NodeTreePanel(QWidget):
    """Панель для відображення дерева нод з прокруткою та zoom."""

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        # Toolbar
        bar = QHBoxLayout()
        self._title = QLabel("⚙️ Дерево нод")
        self._title.setStyleSheet("font-weight: bold; font-size: 13px; padding: 2px 6px;")
        zoom_in = QPushButton("+")
        zoom_in.setFixedWidth(30)
        zoom_in.clicked.connect(lambda: self._view.scale(1.2, 1.2))
        zoom_out = QPushButton("−")
        zoom_out.setFixedWidth(30)
        zoom_out.clicked.connect(lambda: self._view.scale(1 / 1.2, 1 / 1.2))
        fit_btn = QPushButton("⊡ Вмістити")
        fit_btn.clicked.connect(self._fit)
        bar.addWidget(self._title)
        bar.addStretch()
        bar.addWidget(zoom_out)
        bar.addWidget(zoom_in)
        bar.addWidget(fit_btn)
        layout.addLayout(bar)

        self._scene = QGraphicsScene(self)
        self._scene.setBackgroundBrush(QBrush(QColor("#0f172a")))
        self._view = QGraphicsView(self._scene)
        self._view.setRenderHint(QPainter.RenderHint.Antialiasing)
        self._view.setDragMode(QGraphicsView.DragMode.ScrollHandDrag)
        layout.addWidget(self._view)

    def load_graph(self, graph_json: dict) -> None:
        self._scene.clear()
        nodes_data: list[dict] = graph_json.get("nodes", [])
        edges_data: list[dict] = graph_json.get("edges", [])

        id_to_item: dict[str, GraphNodeItem] = {}
        for nd in nodes_data:
            item = GraphNodeItem(nd)
            self._scene.addItem(item)
            id_to_item[nd["id"]] = item

        for ed in edges_data:
            src = id_to_item.get(ed.get("from", ""))
            dst = id_to_item.get(ed.get("to", ""))
            if src and dst:
                edge = EdgeItem(src, dst, ed.get("kind", "data"))
                self._scene.addItem(edge)

        title = graph_json.get("title", "")
        if title:
            self._title.setText(f"⚙️ {title}")

        self._draw_grid()
        self._fit()

    def clear(self) -> None:
        self._scene.clear()
        self._title.setText("⚙️ Дерево нод")

    def _fit(self) -> None:
        r = self._scene.itemsBoundingRect()
        if not r.isEmpty():
            self._view.fitInView(r.adjusted(-40, -40, 40, 40), Qt.AspectRatioMode.KeepAspectRatio)

    def _draw_grid(self) -> None:
        """Малює сітку на фоні."""
        r = self._scene.itemsBoundingRect().adjusted(-200, -200, 200, 200)
        pen = QPen(QColor("#1e293b"), 1)
        step = 40
        x = int(r.left() // step) * step
        while x < r.right():
            self._scene.addLine(x, r.top(), x, r.bottom(), pen)
            x += step
        y = int(r.top() // step) * step
        while y < r.bottom():
            self._scene.addLine(r.left(), y, r.right(), y, pen)
            y += step


# ─────────────────────────────────────────────────────────────────────────────
# Worker-потік
# ─────────────────────────────────────────────────────────────────────────────
class AiWorker(QObject):
    finished = Signal(str)
    failed = Signal(str)

    def __init__(self, config: LocalLlmConfig, system: str, user: str) -> None:
        super().__init__()
        self.config, self.system, self.user = config, system, user

    @Slot()
    def run(self) -> None:
        try:
            client = LocalLlmClient(self.config)
            result = client.chat(
                messages=[
                    {"role": "system", "content": self.system},
                    {"role": "user", "content": self.user},
                ]
            )
            self.finished.emit(result)
        except Exception as err:  # noqa: BLE001
            self.failed.emit(str(err))


# ─────────────────────────────────────────────────────────────────────────────
# Головний діалог
# ─────────────────────────────────────────────────────────────────────────────
class OrderAnalysisDialog(QDialog):
    """
    Вікно аналізу наказу:
    ┌─────────────────────────────────────────────────────────────────┐
    │ [Налаштування моделі — компактна панель]                        │
    ├──────────────────────┬──────────────────────────────────────────┤
    │ ДО (оригінал)  >>>>  │ ПІСЛЯ (результат AI)                    │
    │                      │                                          │
    ├──────────────────────┴──────────────────────────────────────────┤
    │ [Дерево нод — побудоване автоматично або за кнопкою]            │
    └─────────────────────────────────────────────────────────────────┘
    """

    graph_created = Signal(object)

    def __init__(self, registry: NodeRegistry, plugin_dir: Path, parent=None) -> None:
        super().__init__(parent)
        self.registry = registry
        self.plugin_dir = plugin_dir
        self._files: list[Path] = []
        self._last_analysis = ""
        self._last_graph_json: dict | None = None
        self._thread: QThread | None = None
        self._worker: AiWorker | None = None
        self._settings = load_llm_settings()

        self._build_ui()
        self._apply_settings()
        self.resize(1400, 850)
        self.setWindowTitle("🔍 Аналіз наказу — ДО  >>>>  ПІСЛЯ  +  Дерево нод")
        self.setStyleSheet("""
            QDialog { background: #0f172a; color: #e2e8f0; }
            QGroupBox { border: 1px solid #334155; border-radius: 6px;
                        margin-top: 8px; color: #94a3b8; font-size: 11px; }
            QGroupBox::title { subcontrol-origin: margin; left: 8px; padding: 0 4px; }
            QPlainTextEdit, QLineEdit { background: #1e293b; color: #e2e8f0;
                                        border: 1px solid #334155; border-radius: 4px; }
            QComboBox { background: #1e293b; color: #e2e8f0; border: 1px solid #334155; border-radius: 4px; }
            QPushButton { background: #334155; color: #e2e8f0; border: none;
                          border-radius: 4px; padding: 5px 12px; }
            QPushButton:hover { background: #475569; }
            QLabel { color: #e2e8f0; }
            QSplitter::handle { background: #334155; }
        """)

    # ── Побудова UI ────────────────────────────────────────────────────────────
    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(6)

        # ── Панель налаштувань (1 рядок) ──────────────────────────────────────
        cfg_box = QGroupBox("Модель AI")
        cfg_row = QHBoxLayout(cfg_box)
        cfg_row.setSpacing(8)

        self._provider_cb = QComboBox()
        for p in LocalLlmProvider:
            self._provider_cb.addItem(str(p), p)
        self._provider_cb.currentIndexChanged.connect(self._on_provider)
        cfg_row.addWidget(QLabel("Провайдер:"))
        cfg_row.addWidget(self._provider_cb)

        self._model_cb = QComboBox()
        self._model_cb.setEditable(True)
        self._model_cb.setMinimumWidth(220)
        cfg_row.addWidget(QLabel("Модель:"))
        cfg_row.addWidget(self._model_cb)

        self._apikey_ed = QLineEdit()
        self._apikey_ed.setEchoMode(QLineEdit.EchoMode.Password)
        self._apikey_ed.setPlaceholderText("API-ключ (не потрібен для Ollama)")
        self._apikey_ed.setMaximumWidth(220)
        cfg_row.addWidget(QLabel("Ключ:"))
        cfg_row.addWidget(self._apikey_ed)

        refresh_btn = QPushButton("🔄")
        refresh_btn.setFixedWidth(30)
        refresh_btn.setToolTip("Отримати список моделей")
        refresh_btn.clicked.connect(self._refresh_models)
        cfg_row.addWidget(refresh_btn)
        cfg_row.addStretch()
        root.addWidget(cfg_box)

        # ── Головний сплітер: (ДО + ПІСЛЯ) | Дерево нод ─────────────────────
        main_split = QSplitter(Qt.Orientation.Vertical)

        # ── Верхня частина: ДО >>>> ПІСЛЯ ─────────────────────────────────────
        top_split = QSplitter(Qt.Orientation.Horizontal)

        # ─── ДО: панель завантаження та оригінальний текст ────────────────────
        before_panel = QWidget()
        bv = QVBoxLayout(before_panel)
        bv.setContentsMargins(0, 0, 0, 0)
        bv.setSpacing(4)

        before_header = QHBoxLayout()
        before_lbl = QLabel("📄  ДО  — оригінальний текст")
        before_lbl.setStyleSheet("font-weight: bold; font-size: 13px; color: #60a5fa;")
        add_btn = QPushButton("📂 Відкрити файл(и)")
        add_btn.setStyleSheet("background: #1d4ed8; color: white; padding: 5px 12px; border-radius: 4px;")
        add_btn.clicked.connect(self._add_files)
        before_header.addWidget(before_lbl)
        before_header.addStretch()
        before_header.addWidget(add_btn)
        bv.addLayout(before_header)

        self._files_lbl = QLabel("Файли не завантажені")
        self._files_lbl.setStyleSheet("color: #64748b; font-size: 11px;")
        bv.addWidget(self._files_lbl)

        self._before_text = QPlainTextEdit()
        self._before_text.setReadOnly(True)
        self._before_text.setPlaceholderText("Тут з'явиться оригінальний текст наказу після відкриття файлу…")
        self._before_text.setStyleSheet(
            "QPlainTextEdit { background: #0f172a; color: #cbd5e1; "
            "border: 1px solid #1d4ed8; font-family: 'Consolas'; font-size: 12px; }"
        )
        bv.addWidget(self._before_text)
        top_split.addWidget(before_panel)

        # ─── Стрілка ──────────────────────────────────────────────────────────
        arrow_lbl = QLabel(">>>>")
        arrow_lbl.setAlignment(Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignHCenter)
        arrow_lbl.setStyleSheet(
            "color: #f59e0b; font-size: 22px; font-weight: bold; "
            "min-width: 50px; max-width: 50px; background: transparent;"
        )
        arrow_lbl.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Expanding)
        top_split.addWidget(arrow_lbl)

        # ─── ПІСЛЯ: результат аналізу ──────────────────────────────────────────
        after_panel = QWidget()
        av = QVBoxLayout(after_panel)
        av.setContentsMargins(0, 0, 0, 0)
        av.setSpacing(4)

        after_header = QHBoxLayout()
        after_lbl = QLabel("✅  ПІСЛЯ  — результат AI-аналізу")
        after_lbl.setStyleSheet("font-weight: bold; font-size: 13px; color: #34d399;")

        self._extra_prompt = QLineEdit()
        self._extra_prompt.setPlaceholderText("Додатковий запит (необов'язково)")
        self._extra_prompt.setMaximumWidth(300)

        self._analyze_btn = QPushButton("🔍 Аналізувати")
        self._analyze_btn.setStyleSheet(
            "QPushButton { background: #059669; color: white; padding: 5px 16px; border-radius: 4px; font-weight: bold; }"
            "QPushButton:disabled { background: #374151; color: #6b7280; }"
        )
        self._analyze_btn.clicked.connect(self._run_analyze)

        save_btn = QPushButton("💾")
        save_btn.setToolTip("Зберегти результат")
        save_btn.setFixedWidth(30)
        save_btn.clicked.connect(self._save_result)

        after_header.addWidget(after_lbl)
        after_header.addStretch()
        after_header.addWidget(self._extra_prompt)
        after_header.addWidget(self._analyze_btn)
        after_header.addWidget(save_btn)
        av.addLayout(after_header)

        self._after_text = QPlainTextEdit()
        self._after_text.setReadOnly(True)
        self._after_text.setPlaceholderText("Тут з'явиться структурований аналіз наказу…")
        self._after_text.setStyleSheet(
            "QPlainTextEdit { background: #0f172a; color: #d1fae5; "
            "border: 1px solid #059669; font-family: 'Segoe UI'; font-size: 12px; }"
        )
        av.addWidget(self._after_text)
        top_split.addWidget(after_panel)
        top_split.setSizes([520, 50, 520])

        main_split.addWidget(top_split)

        # ── Нижня частина: дерево нод ──────────────────────────────────────────
        node_panel = QWidget()
        nv = QVBoxLayout(node_panel)
        nv.setContentsMargins(0, 0, 0, 0)
        nv.setSpacing(4)

        node_header = QHBoxLayout()
        self._build_graph_btn = QPushButton("⚙️ Побудувати дерево нод з аналізу")
        self._build_graph_btn.setStyleSheet(
            "QPushButton { background: #7c3aed; color: white; padding: 6px 18px; border-radius: 4px; font-weight: bold; }"
            "QPushButton:disabled { background: #374151; color: #6b7280; }"
        )
        self._build_graph_btn.clicked.connect(self._run_graph)

        apply_btn = QPushButton("▶ Відкрити в редакторі")
        apply_btn.setStyleSheet("background: #d97706; color: white; padding: 6px 16px; border-radius: 4px; font-weight: bold;")
        apply_btn.clicked.connect(self._apply_graph)

        self._graph_prompt_ed = QLineEdit()
        self._graph_prompt_ed.setPlaceholderText("Опис сценарію (необов'язково)")
        self._graph_prompt_ed.setMaximumWidth(350)

        node_header.addWidget(self._build_graph_btn)
        node_header.addWidget(self._graph_prompt_ed)
        node_header.addStretch()
        node_header.addWidget(apply_btn)
        nv.addLayout(node_header)

        self._node_tree = NodeTreePanel()
        nv.addWidget(self._node_tree)
        main_split.addWidget(node_panel)

        main_split.setSizes([480, 300])
        root.addWidget(main_split)

        # ── Статус-рядок ───────────────────────────────────────────────────────
        self._status = QLabel("Готово. Відкрийте файл наказу та натисніть «Аналізувати».")
        self._status.setStyleSheet("color: #64748b; font-size: 11px; padding: 2px 4px;")
        root.addWidget(self._status)

    # ── Налаштування провайдера ────────────────────────────────────────────────
    def _apply_settings(self) -> None:
        provider = self._settings.get("provider", LocalLlmProvider.OLLAMA)
        idx = self._provider_cb.findData(provider)
        if idx >= 0:
            self._provider_cb.setCurrentIndex(idx)
        self._on_provider()
        saved = self._settings.get(f"model_{provider}", "")
        if saved:
            self._model_cb.setCurrentText(saved)
        self._apikey_ed.setText(self._settings.get(f"api_key_{provider}", ""))

    def _on_provider(self) -> None:
        provider: LocalLlmProvider = self._provider_cb.currentData()
        models = PROVIDER_PRESET_MODELS.get(provider, [])
        self._model_cb.clear()
        for m in models:
            clean = m.split("#")[0].strip() if "#" in m else m
            self._model_cb.addItem(clean)

    def _refresh_models(self) -> None:
        try:
            models = LocalLlmClient(self._build_config()).fetch_available_models()
            if models:
                self._model_cb.clear()
                for m in models:
                    self._model_cb.addItem(m)
                self._set_status(f"🔄 Знайдено {len(models)} моделей")
        except Exception as err:
            self._set_status(f"❌ {err}")

    def _save_settings(self) -> None:
        provider: LocalLlmProvider = self._provider_cb.currentData()
        s = dict(self._settings)
        s["provider"] = provider
        s[f"model_{provider}"] = self._model_cb.currentText().strip()
        s[f"api_key_{provider}"] = self._apikey_ed.text().strip()
        save_llm_settings(s)
        self._settings = s

    def _build_config(self) -> LocalLlmConfig:
        provider: LocalLlmProvider = self._provider_cb.currentData()
        return LocalLlmConfig(
            provider=provider,
            model=self._model_cb.currentText().strip(),
            api_key=self._apikey_ed.text().strip(),
            base_url=DEFAULT_BASE_URLS.get(provider, ""),
        )

    # ── Файли ────────────────────────────────────────────────────────────────
    def _add_files(self) -> None:
        names, _ = QFileDialog.getOpenFileNames(
            self, "Відкрити файли наказів", "",
            "Документи (*.docx *.txt);;Word (*.docx);;Текст (*.txt);;Всі (*)",
        )
        for f in names:
            p = Path(f)
            if p not in self._files:
                self._files.append(p)
        self._refresh_before_panel()

    def _refresh_before_panel(self) -> None:
        if not self._files:
            self._files_lbl.setText("Файли не завантажені")
            self._before_text.setPlainText("")
            return
        names = "  |  ".join(p.name for p in self._files)
        self._files_lbl.setText(f"📂 {names}")
        combined = []
        for p in self._files:
            try:
                if p.suffix.casefold() == ".docx":
                    from docx import Document
                    doc = Document(str(p))
                    text = "\n".join(par.text for par in doc.paragraphs if par.text.strip())
                else:
                    text = p.read_text(encoding="utf-8", errors="replace")
                combined.append(f"══════════ {p.name} ══════════\n{text}")
            except Exception as err:
                combined.append(f"[Помилка читання {p.name}: {err}]")
        self._before_text.setPlainText("\n\n".join(combined))

    def _file_text(self, max_chars: int = 12000) -> str:
        return self._before_text.toPlainText()[:max_chars]

    # ── Аналіз ───────────────────────────────────────────────────────────────
    def _run_analyze(self) -> None:
        if not self._files:
            QMessageBox.warning(self, "Файли відсутні", "Спочатку відкрийте файл наказу.")
            return
        self._save_settings()
        extra = self._extra_prompt.text().strip()
        user_msg = f"Проаналізуй наказ:\n\n{self._file_text()}"
        if extra:
            user_msg += f"\n\nДодаткове завдання: {extra}"
        self._set_busy(True)
        self._after_text.setPlainText("⏳ Виконується аналіз…")
        self._start_worker(_SYS_ANALYZE, user_msg, self._on_analyze_done)

    def _on_analyze_done(self, result: str) -> None:
        self._last_analysis = result
        self._after_text.setPlainText(result)
        self._set_busy(False)
        self._set_status("✅ Аналіз завершено. Натисніть «Побудувати дерево нод» або задайте сценарій.")
        # Автоматично запускаємо побудову графа
        self._run_graph()

    def _save_result(self) -> None:
        text = self._after_text.toPlainText()
        if not text:
            return
        fn, _ = QFileDialog.getSaveFileName(self, "Зберегти аналіз", "", "Markdown (*.md);;Текст (*.txt)")
        if fn:
            Path(fn).write_text(text, encoding="utf-8")
            self._set_status(f"💾 Збережено: {Path(fn).name}")

    # ── Побудова графа ────────────────────────────────────────────────────────
    def _run_graph(self) -> None:
        if not self._last_analysis and not self._file_text():
            QMessageBox.warning(self, "Немає даних", "Спочатку виконайте аналіз наказу.")
            return
        self._save_settings()

        node_list = "\n".join(
            f"- {d.type_id}: {d.name}" for d in self.registry.all()[:60]
        )
        user_task = self._graph_prompt_ed.text().strip() or "Автоматизувати обробку цього наказу"
        context = (
            f"Завдання: {user_task}\n\n"
            f"Аналіз наказу:\n{self._last_analysis[:3000]}\n\n"
            f"Доступні ноди:\n{node_list}"
        )
        self._set_busy(True)
        self._node_tree.clear()
        self._set_status("⏳ Генерація дерева нод…")
        self._start_worker(_SYS_GRAPH, context, self._on_graph_done)

    def _on_graph_done(self, result: str) -> None:
        self._set_busy(False)
        raw = result.strip()
        # Витягуємо JSON з markdown-огорток якщо є
        if "```json" in raw:
            raw = raw.split("```json")[1].split("```")[0].strip()
        elif "```" in raw:
            raw = raw.split("```")[1].split("```")[0].strip()
        try:
            graph_json = json.loads(raw)
            self._last_graph_json = graph_json
            self._node_tree.load_graph(graph_json)
            self._set_status(
                f"✅ Дерево нод побудовано: {len(graph_json.get('nodes', []))} нод, "
                f"{len(graph_json.get('edges', []))} з'єднань. Натисніть «Відкрити в редакторі»."
            )
        except json.JSONDecodeError as err:
            self._set_status(f"⚠️ AI повернув не-JSON: {err}. Спробуйте ще раз.")

    # ── Застосування графа ────────────────────────────────────────────────────
    def _apply_graph(self) -> None:
        if not self._last_graph_json:
            QMessageBox.information(self, "Граф відсутній", "Спочатку побудуйте дерево нод.")
            return
        try:
            gj = self._last_graph_json
            graph = GraphModel(name=gj.get("title", "Наказ — автоматизація"))
            id_map: dict[str, str] = {}
            for nd in gj.get("nodes", []):
                new_id = str(uuid.uuid4())[:8]
                id_map[nd["id"]] = new_id
                graph.nodes.append(NodeModel(
                    id=new_id,
                    type_id=nd.get("type_id", "builtin.flow.start"),
                    x=nd.get("x", 0),
                    y=nd.get("y", 0),
                    parameters=nd.get("params", {}),
                ))
            for ed in gj.get("edges", []):
                src = id_map.get(ed.get("from", ""))
                dst = id_map.get(ed.get("to", ""))
                if src and dst:
                    graph.connections.append(ConnectionModel(
                        source_node=src,
                        source_port=ed.get("from_port", "then"),
                        target_node=dst,
                        target_port=ed.get("to_port", "exec"),
                        kind=ed.get("kind", "data"),
                    ))
            self.graph_created.emit(graph)
            self.accept()
        except Exception as err:
            QMessageBox.critical(self, "Помилка", str(err))

    # ── Worker-інфраструктура ─────────────────────────────────────────────────
    def _start_worker(self, system: str, user: str, on_done) -> None:
        if self._thread and self._thread.isRunning():
            return
        self._thread = QThread(self)
        self._worker = AiWorker(self._build_config(), system, user)
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.finished.connect(on_done)
        self._worker.finished.connect(self._thread.quit)
        self._worker.failed.connect(self._on_fail)
        self._worker.failed.connect(self._thread.quit)
        self._thread.start()

    def _on_fail(self, err: str) -> None:
        self._set_busy(False)
        self._set_status(f"❌ Помилка: {err}")
        QMessageBox.critical(self, "Помилка AI", err)

    def _set_busy(self, busy: bool) -> None:
        self._analyze_btn.setEnabled(not busy)
        self._build_graph_btn.setEnabled(not busy)

    def _set_status(self, text: str) -> None:
        self._status.setText(text)

    def closeEvent(self, event) -> None:
        self._save_settings()
        super().closeEvent(event)
