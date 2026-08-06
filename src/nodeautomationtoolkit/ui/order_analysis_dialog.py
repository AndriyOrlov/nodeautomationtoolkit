"""
Діалог «Навчання за прикладом» — режим ДО / ПІСЛЯ.

Користувач завантажує:
  - Ліворуч (ДО)    — оригінальний наказ (вхідний файл)
  - Праворуч (ПІСЛЯ) — результат ручної обробки (очікуваний вихід)

AI аналізує РІЗНИЦЮ між двома документами та будує граф нод,
який автоматично відтворить ту саму трансформацію для будь-якого
нового наказу тієї ж структури.
"""
from __future__ import annotations

import json
import uuid
from pathlib import Path

from PySide6.QtCore import (
    QObject,
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
)
from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QFileDialog,
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
# Системний промпт — аналіз різниці та побудова графа
# ─────────────────────────────────────────────────────────────────────────────
_SYS_DIFF_TO_GRAPH = """Ти — AI-аналітик трансформацій документів для Node Automation Toolkit.

Тобі дають ДВА документи:
1. ВХІДНИЙ (ДО)  — оригінальний наказ
2. ОЧІКУВАНИЙ (ПІСЛЯ) — результат ручної обробки цього наказу

## Твоє завдання:
1. Знайти ВСІ відмінності між документами (що додано, що видалено, що замінено, що перегруповано)
2. Сформулювати кроки трансформації зрозумілою мовою (Markdown-звіт)
3. Побудувати JSON-граф нод, який автоматично відтворить ці самі трансформації

## Формат відповіді — СТРОГО:

===АНАЛІЗ===
### Виявлені трансформації
1. [крок 1] — наприклад: "Замінено відкрите найменування '167 окрема механізована бригада' → 'в/ч А0000'"
2. [крок 2] — наприклад: "Витягнуто тільки пункти де згадується конкретна ВЧ"
3. ...

### Що залишилось незмінним
- ...

### Шаблон трансформації
Короткий опис: цей граф можна застосувати до будь-якого наказу такої ж структури.

===ГРАФ===
{
  "title": "...",
  "summary": "...",
  "nodes": [
    {"id": "n1", "type_id": "builtin.flow.start", "label": "Старт", "x": 0, "y": 0, "params": {}},
    {"id": "n2", "type_id": "builtin.word.read_docx", "label": "Читати наказ", "x": 220, "y": 0, "params": {"path": ""}},
    ...
  ],
  "edges": [
    {"from": "n1", "from_port": "then", "to": "n2", "to_port": "exec", "kind": "execution"},
    ...
  ]
}

## Доступні type_id нод:
Потік: builtin.flow.start · builtin.flow.sub_start · builtin.flow.branch · builtin.flow.sequence
Word: builtin.word.read_docx · builtin.word.replace_in_docx · builtin.word.merge_docx · builtin.word.word_count · builtin.word.create_docx · builtin.word.save_selected_paragraphs
Наказ: builtin.order.read_recipient_mapping · builtin.order.map_military_units · builtin.order.analyze_senders · builtin.order.split_by_senders · builtin.order.groups_to_ciphers · builtin.order_batch.create_unit_extracts
Excel: builtin.excel.read_sheet · builtin.excel.save_table · builtin.excel.write_cell
Файли: builtin.files.list_files · builtin.files.move_file · builtin.files.create_folder · builtin.files.create
Текст: builtin.text.replace · builtin.text.fill_template · builtin.text.today_date · builtin.text.clean · builtin.text.regex_replace · builtin.text.split_lines
Вивід: builtin.output.show_result · builtin.output.show_table

## Правила JSON-графа:
- x координати кратні 220, y кратні 120 (паралельні потоки — різні y)
- Вузли виконання (exec/then): kind="execution"
- Дані: kind="data"
- Завжди починай з builtin.flow.start

## ВАЖЛИВО:
- Відповідай ВИКЛЮЧНО у форматі ===АНАЛІЗ=== ... ===ГРАФ=== { ... }
- Граф має відтворювати ТІЛЬКИ виявлені трансформації — не вигадуй зайвого
- Відповідай українською мовою"""


# ─────────────────────────────────────────────────────────────────────────────
# Кольори нод
# ─────────────────────────────────────────────────────────────────────────────
_NODE_COLORS = {
    "builtin.flow": "#6366f1",
    "builtin.word": "#0284c7",
    "builtin.order": "#d97706",
    "builtin.order_batch": "#b45309",
    "builtin.excel": "#15803d",
    "builtin.files": "#64748b",
    "builtin.text": "#7c3aed",
    "builtin.output": "#dc2626",
}

NODE_W, NODE_H = 200, 60


def _color(type_id: str) -> QColor:
    for prefix, hex_c in _NODE_COLORS.items():
        if type_id.startswith(prefix):
            return QColor(hex_c)
    return QColor("#334155")


class GraphNodeItem(QGraphicsRectItem):
    def __init__(self, nd: dict) -> None:
        super().__init__(0, 0, NODE_W, NODE_H)
        color = _color(nd.get("type_id", ""))
        grad = QLinearGradient(0, 0, 0, NODE_H)
        grad.setColorAt(0, color.lighter(140))
        grad.setColorAt(1, color)
        self.setBrush(QBrush(grad))
        self.setPen(QPen(color.darker(150), 1.5))
        self.setPos(nd.get("x", 0), nd.get("y", 0))
        self.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsMovable, True)
        self.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsSelectable, True)
        self._type_id = nd.get("type_id", "")

        label = nd.get("label", self._type_id.split(".")[-1])
        t = QGraphicsTextItem(label, self)
        t.setDefaultTextColor(QColor("white"))
        t.setFont(QFont("Segoe UI", 9, QFont.Weight.Bold))
        t.setTextWidth(NODE_W - 12)
        t.setPos(6, 6)

        sub = QGraphicsTextItem(self._type_id, self)
        sub.setDefaultTextColor(QColor(255, 255, 255, 130))
        sub.setFont(QFont("Segoe UI", 7))
        sub.setTextWidth(NODE_W - 12)
        sub.setPos(6, 32)

    def out_pos(self) -> tuple[float, float]:
        return self.x() + NODE_W, self.y() + NODE_H / 2

    def in_pos(self) -> tuple[float, float]:
        return self.x(), self.y() + NODE_H / 2


class EdgeItem(QGraphicsLineItem):
    def __init__(self, src: GraphNodeItem, dst: GraphNodeItem, kind: str = "data") -> None:
        sx, sy = src.out_pos()
        dx, dy = dst.in_pos()
        super().__init__(sx, sy, dx, dy)
        color = QColor("#818cf8") if kind == "execution" else QColor("#64748b")
        self.setPen(QPen(color, 2 if kind == "execution" else 1.5))
        self.setZValue(-1)


class NodeTreePanel(QWidget):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(2)

        bar = QHBoxLayout()
        self._lbl = QLabel("⚙️  Автоматично побудований граф")
        self._lbl.setStyleSheet("font-weight: bold; font-size: 13px; color: #a78bfa; padding: 2px 6px;")
        zi = QPushButton("+")
        zi.setFixedWidth(28)
        zi.clicked.connect(lambda: self._view.scale(1.2, 1.2))
        zo = QPushButton("−")
        zo.setFixedWidth(28)
        zo.clicked.connect(lambda: self._view.scale(1 / 1.2, 1 / 1.2))
        fit = QPushButton("⊡ Вмістити")
        fit.clicked.connect(self._fit)
        bar.addWidget(self._lbl)
        bar.addStretch()
        bar.addWidget(zo)
        bar.addWidget(zi)
        bar.addWidget(fit)
        layout.addLayout(bar)

        self._scene = QGraphicsScene(self)
        self._scene.setBackgroundBrush(QBrush(QColor("#0f172a")))
        self._view = QGraphicsView(self._scene)
        self._view.setRenderHint(QPainter.RenderHint.Antialiasing)
        self._view.setDragMode(QGraphicsView.DragMode.ScrollHandDrag)
        layout.addWidget(self._view)

    def load_graph(self, gj: dict) -> None:
        self._scene.clear()
        id_map: dict[str, GraphNodeItem] = {}
        for nd in gj.get("nodes", []):
            item = GraphNodeItem(nd)
            self._scene.addItem(item)
            id_map[nd["id"]] = item
        for ed in gj.get("edges", []):
            s = id_map.get(ed.get("from", ""))
            d = id_map.get(ed.get("to", ""))
            if s and d:
                self._scene.addItem(EdgeItem(s, d, ed.get("kind", "data")))
        title = gj.get("title", "")
        if title:
            self._lbl.setText(f"⚙️  {title}")
        self._draw_grid()
        self._fit()

    def clear(self) -> None:
        self._scene.clear()
        self._lbl.setText("⚙️  Автоматично побудований граф")

    def _fit(self) -> None:
        r = self._scene.itemsBoundingRect()
        if not r.isEmpty():
            self._view.fitInView(r.adjusted(-40, -40, 40, 40), Qt.AspectRatioMode.KeepAspectRatio)

    def _draw_grid(self) -> None:
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
# Worker
# ─────────────────────────────────────────────────────────────────────────────
class DiffWorker(QObject):
    finished = Signal(str)
    failed = Signal(str)

    def __init__(self, config: LocalLlmConfig, system: str, user: str) -> None:
        super().__init__()
        self.config, self.system, self.user = config, system, user

    @Slot()
    def run(self) -> None:
        try:
            result = LocalLlmClient(self.config).chat(
                messages=[
                    {"role": "system", "content": self.system},
                    {"role": "user", "content": self.user},
                ]
            )
            self.finished.emit(result)
        except Exception as err:  # noqa: BLE001
            self.failed.emit(str(err))


# ─────────────────────────────────────────────────────────────────────────────
# Панель завантаження одного файлу
# ─────────────────────────────────────────────────────────────────────────────
class FileDropPanel(QWidget):
    """Панель «завантажити файл + показати його текст»."""

    def __init__(self, title: str, color: str, placeholder: str, parent=None) -> None:
        super().__init__(parent)
        self._path: Path | None = None
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)

        header = QHBoxLayout()
        self._title_lbl = QLabel(title)
        self._title_lbl.setStyleSheet(
            f"font-weight: bold; font-size: 14px; color: {color}; padding: 2px 4px;"
        )
        self._open_btn = QPushButton("📂 Відкрити")
        self._open_btn.setStyleSheet(
            f"QPushButton {{ background: {color}; color: white; padding: 5px 14px; "
            f"border-radius: 4px; font-weight: bold; }}"
            f"QPushButton:hover {{ opacity: 0.8; }}"
        )
        self._open_btn.clicked.connect(self._open_file)
        self._clear_btn = QPushButton("✕")
        self._clear_btn.setFixedWidth(28)
        self._clear_btn.clicked.connect(self._clear)
        header.addWidget(self._title_lbl)
        header.addStretch()
        header.addWidget(self._open_btn)
        header.addWidget(self._clear_btn)
        layout.addLayout(header)

        self._file_lbl = QLabel("Файл не завантажений")
        self._file_lbl.setStyleSheet("color: #475569; font-size: 11px; padding: 2px 4px;")
        layout.addWidget(self._file_lbl)

        self._text = QPlainTextEdit()
        self._text.setReadOnly(True)
        self._text.setPlaceholderText(placeholder)
        self._text.setStyleSheet(
            f"QPlainTextEdit {{ background: #0f172a; color: #cbd5e1; "
            f"border: 1px solid {color}44; border-radius: 4px; "
            f"font-family: 'Consolas'; font-size: 11px; }}"
        )
        layout.addWidget(self._text)

    def _open_file(self) -> None:
        fn, _ = QFileDialog.getOpenFileName(
            self, "Відкрити файл", "",
            "Документи (*.docx *.txt);;Word (*.docx);;Текст (*.txt);;Всі (*)",
        )
        if fn:
            self._load(Path(fn))

    def _load(self, path: Path) -> None:
        self._path = path
        try:
            if path.suffix.casefold() == ".docx":
                from docx import Document
                doc = Document(str(path))
                text = "\n".join(p.text for p in doc.paragraphs if p.text.strip())
            else:
                text = path.read_text(encoding="utf-8", errors="replace")
            self._text.setPlainText(text)
            self._file_lbl.setText(f"✅ {path.name}  ({len(text):,} символів)")
            self._file_lbl.setStyleSheet("color: #34d399; font-size: 11px; padding: 2px 4px;")
        except Exception as err:
            self._text.setPlainText(f"Помилка читання: {err}")

    def _clear(self) -> None:
        self._path = None
        self._text.setPlainText("")
        self._file_lbl.setText("Файл не завантажений")
        self._file_lbl.setStyleSheet("color: #475569; font-size: 11px; padding: 2px 4px;")

    def get_text(self, max_chars: int = 8000) -> str:
        return self._text.toPlainText()[:max_chars]

    def has_file(self) -> bool:
        return bool(self._text.toPlainText().strip())


# ─────────────────────────────────────────────────────────────────────────────
# Головний діалог
# ─────────────────────────────────────────────────────────────────────────────
class OrderAnalysisDialog(QDialog):
    """
    Вікно «Навчання за прикладом»:

    ┌──────────────────────────────────────────────────────────────────┐
    │ [Модель AI — 1 рядок]                                            │
    ├──────────────────────┬──────────┬───────────────────────────────┤
    │  📄 ДО               │  >>>>   │  ✅ ПІСЛЯ                     │
    │  Оригінальний наказ  │         │  Результат ручної обробки     │
    │  (завантажити файл)  │         │  (завантажити файл)           │
    ├──────────────────────┴──────────┴───────────────────────────────┤
    │  🔬 Аналіз відмінностей (що змінилось — Markdown)              │
    ├──────────────────────────────────────────────────────────────────┤
    │  ⚙️ Граф нод (автоматично повторює виявлену трансформацію)      │
    └──────────────────────────────────────────────────────────────────┘
    """

    graph_created = Signal(object)

    def __init__(self, registry: NodeRegistry, plugin_dir: Path, parent=None) -> None:
        super().__init__(parent)
        self.registry = registry
        self.plugin_dir = plugin_dir
        self._last_analysis = ""
        self._last_graph_json: dict | None = None
        self._thread: QThread | None = None
        self._worker: DiffWorker | None = None
        self._settings = load_llm_settings()

        self._build_ui()
        self._apply_settings()
        self.resize(1400, 900)
        self.setWindowTitle("🔬 Навчання за прикладом — ДО  >>>>  ПІСЛЯ  →  Граф автоматизації")
        self.setStyleSheet("""
            QDialog  { background: #0f172a; color: #e2e8f0; }
            QGroupBox { border: 1px solid #334155; border-radius: 6px;
                        margin-top: 8px; color: #94a3b8; font-size: 11px; }
            QGroupBox::title { subcontrol-origin: margin; left: 8px; padding: 0 4px; }
            QPlainTextEdit, QLineEdit { background: #1e293b; color: #e2e8f0;
                                        border: 1px solid #334155; border-radius: 4px; }
            QComboBox { background: #1e293b; color: #e2e8f0; border: 1px solid #334155;
                        border-radius: 4px; min-height: 24px; }
            QPushButton { background: #1e293b; color: #e2e8f0; border: 1px solid #334155;
                          border-radius: 4px; padding: 4px 10px; }
            QPushButton:hover { background: #334155; }
            QLabel  { color: #e2e8f0; }
            QSplitter::handle { background: #1e293b; width: 4px; height: 4px; }
        """)

    # ── Побудова UI ────────────────────────────────────────────────────────────
    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(6)

        # ── Рядок налаштувань моделі ───────────────────────────────────────────
        cfg = QGroupBox("AI-модель")
        cfg_row = QHBoxLayout(cfg)
        cfg_row.setSpacing(6)

        self._prov = QComboBox()
        for p in LocalLlmProvider:
            self._prov.addItem(str(p), p)
        self._prov.currentIndexChanged.connect(self._on_provider)
        cfg_row.addWidget(QLabel("Провайдер:"))
        cfg_row.addWidget(self._prov)

        self._model = QComboBox()
        self._model.setEditable(True)
        self._model.setMinimumWidth(200)
        refresh_btn = QPushButton("🔄")
        refresh_btn.setFixedWidth(28)
        refresh_btn.setToolTip("Отримати список моделей із сервера")
        refresh_btn.clicked.connect(self._refresh_models)
        cfg_row.addWidget(QLabel("Модель:"))
        cfg_row.addWidget(self._model)
        cfg_row.addWidget(refresh_btn)

        self._apikey = QLineEdit()
        self._apikey.setEchoMode(QLineEdit.EchoMode.Password)
        self._apikey.setPlaceholderText("API-ключ (не потрібен для Ollama / Embedded)")
        self._apikey.setMaximumWidth(230)
        cfg_row.addWidget(QLabel("Ключ:"))
        cfg_row.addWidget(self._apikey)
        cfg_row.addStretch()
        root.addWidget(cfg)

        # ── Головний вертикальний сплітер ──────────────────────────────────────
        vsplit = QSplitter(Qt.Orientation.Vertical)

        # ── Верхня частина: ДО >>>> ПІСЛЯ (горизонтальний сплітер) ────────────
        top_split = QSplitter(Qt.Orientation.Horizontal)

        # Ліва панель — ДО
        self._before = FileDropPanel(
            title="📄  ДО — оригінальний наказ",
            color="#3b82f6",
            placeholder="Завантажте вхідний файл наказу (.docx або .txt)…\n\n"
                        "Це те, з чого починається обробка.",
        )
        top_split.addWidget(self._before)

        # Стрілка посередині
        arrow = QLabel(">>>>")
        arrow.setAlignment(Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignHCenter)
        arrow.setStyleSheet(
            "color: #f59e0b; font-size: 24px; font-weight: 900; "
            "min-width: 54px; max-width: 54px; background: transparent; letter-spacing: 2px;"
        )
        arrow.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Expanding)
        top_split.addWidget(arrow)

        # Права панель — ПІСЛЯ
        self._after = FileDropPanel(
            title="✅  ПІСЛЯ — очікуваний результат",
            color="#22c55e",
            placeholder="Завантажте файл результату ручної обробки (.docx або .txt)…\n\n"
                        "Це те, що має виходити на виході автоматизації.",
        )
        top_split.addWidget(self._after)
        top_split.setSizes([540, 54, 540])
        vsplit.addWidget(top_split)

        # ── Середня частина: аналіз відмінностей ──────────────────────────────
        diff_panel = QWidget()
        dv = QVBoxLayout(diff_panel)
        dv.setContentsMargins(0, 0, 0, 0)
        dv.setSpacing(4)

        diff_header = QHBoxLayout()
        diff_lbl = QLabel("🔬  Аналіз відмінностей — що змінилось і чому")
        diff_lbl.setStyleSheet("font-weight: bold; font-size: 13px; color: #f59e0b;")

        self._extra = QLineEdit()
        self._extra.setPlaceholderText("Додаткові пояснення (необов'язково)")
        self._extra.setMaximumWidth(340)

        self._run_btn = QPushButton("🔬 Порівняти та побудувати граф")
        self._run_btn.setStyleSheet(
            "QPushButton { background: #d97706; color: white; padding: 6px 20px; "
            "border-radius: 5px; font-weight: bold; font-size: 13px; }"
            "QPushButton:disabled { background: #374151; color: #6b7280; }"
        )
        self._run_btn.clicked.connect(self._run)

        save_btn = QPushButton("💾")
        save_btn.setToolTip("Зберегти аналіз")
        save_btn.setFixedWidth(30)
        save_btn.clicked.connect(self._save_analysis)

        diff_header.addWidget(diff_lbl)
        diff_header.addStretch()
        diff_header.addWidget(self._extra)
        diff_header.addWidget(self._run_btn)
        diff_header.addWidget(save_btn)
        dv.addLayout(diff_header)

        self._diff_text = QPlainTextEdit()
        self._diff_text.setReadOnly(True)
        self._diff_text.setPlaceholderText(
            "Тут з'явиться аналіз відмінностей між ДО та ПІСЛЯ:\n"
            "— що саме змінилось\n"
            "— які трансформації були виконані вручну\n"
            "— шаблон для автоматизації"
        )
        self._diff_text.setStyleSheet(
            "QPlainTextEdit { background: #0f172a; color: #fde68a; "
            "border: 1px solid #d9780644; border-radius: 4px; "
            "font-family: 'Segoe UI'; font-size: 12px; }"
        )
        self._diff_text.setMaximumHeight(180)
        dv.addWidget(self._diff_text)
        vsplit.addWidget(diff_panel)

        # ── Нижня частина: дерево нод ──────────────────────────────────────────
        tree_panel = QWidget()
        tv = QVBoxLayout(tree_panel)
        tv.setContentsMargins(0, 0, 0, 0)
        tv.setSpacing(4)

        tree_header = QHBoxLayout()

        apply_btn = QPushButton("▶  Відкрити граф в редакторі")
        apply_btn.setStyleSheet(
            "QPushButton { background: #7c3aed; color: white; padding: 6px 18px; "
            "border-radius: 5px; font-weight: bold; }"
            "QPushButton:hover { background: #6d28d9; }"
        )
        apply_btn.clicked.connect(self._apply_graph)

        tree_header.addStretch()
        tree_header.addWidget(apply_btn)
        tv.addLayout(tree_header)

        self._node_tree = NodeTreePanel()
        tv.addWidget(self._node_tree)
        vsplit.addWidget(tree_panel)

        vsplit.setSizes([400, 180, 280])
        root.addWidget(vsplit)

        # ── Статус ────────────────────────────────────────────────────────────
        self._status = QLabel(
            "1. Завантажте файл ДО (оригінал)  "
            "2. Завантажте файл ПІСЛЯ (ручний результат)  "
            "3. Натисніть «Порівняти та побудувати граф»"
        )
        self._status.setStyleSheet("color: #64748b; font-size: 11px; padding: 2px 6px;")
        root.addWidget(self._status)

    # ── Налаштування ────────────────────────────────────────────────────────────
    def _get_provider(self) -> LocalLlmProvider:
        data = self._prov.currentData()
        if isinstance(data, LocalLlmProvider):
            return data
        try:
            return LocalLlmProvider(str(data))
        except ValueError:
            return LocalLlmProvider.OLLAMA

    def _apply_settings(self) -> None:
        settings = load_llm_settings()
        try:
            provider = LocalLlmProvider(settings["provider"])
        except ValueError:
            provider = LocalLlmProvider.OLLAMA

        idx = self._prov.findData(provider)
        if idx >= 0:
            self._prov.setCurrentIndex(idx)
        self._on_provider()
        
        if settings["model"]:
            self._model.setCurrentText(settings["model"])
        self._apikey.setText(settings["api_key"])

    def _on_provider(self, *args) -> None:
        provider = self._get_provider()
        models = PROVIDER_PRESET_MODELS.get(provider, [])
        self._model.clear()
        for m in models:
            self._model.addItem(m.split("#")[0].strip() if "#" in m else m)
            
        settings = load_llm_settings(provider)
        if settings["model"]:
            self._model.setCurrentText(settings["model"])
        self._apikey.setText(settings["api_key"])

    def _refresh_models(self) -> None:
        try:
            models = LocalLlmClient(self._build_config()).fetch_available_models()
            if models:
                self._model.clear()
                for m in models:
                    self._model.addItem(m)
                self._set_status(f"🔄 Знайдено {len(models)} моделей")
        except Exception as err:
            self._set_status(f"❌ {err}")

    def _save_settings(self) -> None:
        provider = self._get_provider()
        base_url = load_llm_settings(provider)["base_url"]
        save_llm_settings(
            provider_value=provider.value,
            base_url=base_url,
            model=self._model.currentText().strip(),
            api_key=self._apikey.text().strip(),
        )

    def _build_config(self) -> LocalLlmConfig:
        provider = self._get_provider()
        base_url = load_llm_settings(provider)["base_url"]
        return LocalLlmConfig(
            provider=provider,
            model=self._model.currentText().strip(),
            api_key=self._apikey.text().strip(),
            base_url=base_url,
        )

    # ── Запуск аналізу ─────────────────────────────────────────────────────────
    def _run(self) -> None:
        if not self._before.has_file():
            QMessageBox.warning(self, "ДО відсутній", "Завантажте файл ДО (оригінальний наказ).")
            return
        if not self._after.has_file():
            QMessageBox.warning(self, "ПІСЛЯ відсутній", "Завантажте файл ПІСЛЯ (результат обробки).")
            return

        self._save_settings()

        def _compact_sample(text: str, max_chars: int = 2500) -> str:
            if len(text) <= max_chars:
                return text
            return text[:max_chars] + f"\n\n[... Текст фрагменту обрізано для аналізу: всього {len(text)} символів ...]"

        b_sample = _compact_sample(before_text)
        a_sample = _compact_sample(after_text)

        # Перелік доступних нод для контексту
        node_list = "\n".join(
            f"- {d.type_id}: {d.name} ({d.category})"
            for d in self.registry.all()[:60]
        )

        config = self._build_config()

        # ── Захист даних: для онлайнових AI не надсилаємо вміст документів/файлів! ──
        if not config.is_local:
            user_msg = (
                f"## ЗАХИСТ ДАНИХ (ОНЛАЙН AI ПРОВАЙДЕР):\n"
                f"[Вміст файлів та прикріплених документів вилучено з міркувань безпеки. Передається лише опис функціоналу доступних нод]\n\n"
                f"## ОПИС ЗАВДАННЯ АВТОМАТИЗАЦІЇ:\n"
                f"{extra or 'Побудуй граф автоматизації за наявним каталогом нод'}\n\n"
                f"## Доступні ноди для побудови графа:\n{node_list}"
            )
        else:
            user_msg = (
                f"## ВХІДНИЙ ДОКУМЕНТ (ДО):\n\n{b_sample}\n\n"
                f"{'=' * 60}\n\n"
                f"## ОЧІКУВАНИЙ РЕЗУЛЬТАТ (ПІСЛЯ):\n\n{a_sample}\n\n"
                f"{'=' * 60}\n\n"
                f"## Доступні ноди для побудови графа:\n{node_list}"
            )
            if extra:
                user_msg += f"\n\n## Додаткові пояснення від користувача:\n{extra}"

        self._set_busy(True)
        self._diff_text.setPlainText("⏳ AI аналізує відмінності між ДО та ПІСЛЯ…")
        self._node_tree.clear()
        self._set_status("⏳ Виконується аналіз та генерація графа…")

        self._thread = QThread(self)
        self._worker = DiffWorker(self._build_config(), _SYS_DIFF_TO_GRAPH, user_msg)
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.finished.connect(self._on_done)
        self._worker.finished.connect(self._thread.quit)
        self._worker.failed.connect(self._on_fail)
        self._worker.failed.connect(self._thread.quit)
        self._thread.start()

    def _on_done(self, raw: str) -> None:
        self._set_busy(False)

        # ── Розбираємо відповідь: ===АНАЛІЗ=== ... ===ГРАФ=== { ... } ─────────
        analysis_text = ""
        graph_json: dict | None = None

        if "===АНАЛІЗ===" in raw and "===ГРАФ===" in raw:
            parts = raw.split("===ГРАФ===", 1)
            analysis_text = parts[0].replace("===АНАЛІЗ===", "").strip()
            graph_raw = parts[1].strip()
        elif "===ГРАФ===" in raw:
            analysis_text = ""
            graph_raw = raw.split("===ГРАФ===", 1)[1].strip()
        else:
            # Спроба знайти JSON напряму
            analysis_text = raw
            graph_raw = raw

        # Витягуємо JSON
        if "```json" in graph_raw:
            graph_raw = graph_raw.split("```json")[1].split("```")[0].strip()
        elif "```" in graph_raw:
            graph_raw = graph_raw.split("```")[1].split("```")[0].strip()

        # Шукаємо перший { у рядку
        brace_idx = graph_raw.find("{")
        if brace_idx >= 0:
            graph_raw = graph_raw[brace_idx:]

        try:
            graph_json = json.loads(graph_raw)
        except json.JSONDecodeError:
            pass

        self._last_analysis = analysis_text
        self._diff_text.setPlainText(analysis_text if analysis_text else raw)

        if graph_json:
            self._last_graph_json = graph_json
            self._node_tree.load_graph(graph_json)
            n_nodes = len(graph_json.get("nodes", []))
            n_edges = len(graph_json.get("edges", []))
            self._set_status(
                f"✅ Готово! Граф побудовано: {n_nodes} нод, {n_edges} з'єднань. "
                f"Натисніть «Відкрити граф в редакторі»."
            )
        else:
            self._set_status("⚠️ AI не зміг сформувати JSON-граф. Перевірте аналіз і спробуйте ще раз.")

    def _on_fail(self, err: str) -> None:
        self._set_busy(False)
        self._set_status(f"❌ Помилка: {err}")
        QMessageBox.critical(self, "Помилка AI", err)

    def _save_analysis(self) -> None:
        text = self._diff_text.toPlainText()
        if not text:
            return
        fn, _ = QFileDialog.getSaveFileName(
            self, "Зберегти аналіз", "", "Markdown (*.md);;Текст (*.txt)"
        )
        if fn:
            Path(fn).write_text(text, encoding="utf-8")
            self._set_status(f"💾 Збережено: {Path(fn).name}")

    # ── Застосування графа ────────────────────────────────────────────────────
    def _apply_graph(self) -> None:
        if not self._last_graph_json:
            QMessageBox.information(self, "Граф відсутній", "Спочатку виконайте порівняння.")
            return
        try:
            gj = self._last_graph_json
            graph = GraphModel(name=gj.get("title", "Автоматизація за прикладом"))
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

    # ── Утиліти ────────────────────────────────────────────────────────────────
    def _set_busy(self, busy: bool) -> None:
        self._run_btn.setEnabled(not busy)

    def _set_status(self, text: str) -> None:
        self._status.setText(text)

    def closeEvent(self, event) -> None:
        self._save_settings()
        super().closeEvent(event)
