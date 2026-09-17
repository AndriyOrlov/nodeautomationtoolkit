"""Тема «Obsidian»: палітра, таблиця стилів і шрифти Qt-оболонки генератора.

Усе локальне. Шрифти беруться із системної теки Windows, жодних мережевих
запитів немає (PROJECT_RULES 1.1): мокап посилався на Google Fonts, тому
Inter та JetBrains Mono замінено на Segoe UI і Consolas.
"""

from __future__ import annotations

import ctypes
import os
import sys

from PySide6.QtGui import QColor, QFont, QFontDatabase, QPalette
from PySide6.QtWidgets import QApplication, QWidget

SANS = "Segoe UI"
MONO = "Consolas"

C = {
    "bg_deep": "#070b14",
    "bg_window": "#0b1120",
    "bg_raised": "#0f172a",
    "surface": "#111827",
    "surface_2": "#1e293b",
    "border": "#1f293d",
    "border_soft": "#1e293b",
    "border_strong": "#334155",
    "text_strong": "#f8fafc",
    "text": "#e2e8f0",
    "text_soft": "#cbd5e1",
    "text_muted": "#94a3b8",
    "text_faint": "#64748b",
    "text_dim": "#475569",
    "emerald": "#10b981",
    "emerald_soft": "#34d399",
    "emerald_deep": "#059669",
    "emerald_text": "#6ee7b7",
    "teal": "#14b8a6",
    "teal_soft": "#2dd4bf",
    "teal_deep": "#0d9488",
    "teal_text": "#5eead4",
    "indigo": "#4f46e5",
    "indigo_soft": "#6366f1",
    "indigo_text": "#a5b4fc",
    "violet": "#7c3aed",
    "violet_soft": "#8b5cf6",
    "violet_text": "#c4b5fd",
    "amber": "#f59e0b",
    "amber_soft": "#fbbf24",
    "amber_text": "#fcd34d",
    "sky": "#0284c7",
    "sky_soft": "#38bdf8",
    "sky_text": "#7dd3fc",
    "rose": "#e11d48",
    "rose_soft": "#fb7185",
    "rose_text": "#fda4af",
    "blue": "#60a5fa",
    "blue_text": "#93c5fd",
}

#: Пігулки в таблицях: (фон, рамка, текст).
PILL_TONES = {
    "emerald": ("#022c22", "#065f46", C["emerald_text"]),
    "rose": ("#2a0a12", "#881337", C["rose_text"]),
    "indigo": ("#1e1b4b", "#3730a3", C["indigo_text"]),
    "slate": ("#1e293b", "#334155", C["text_soft"]),
    "amber": ("#2a1d05", "#92400e", C["amber_text"]),
    "teal": ("#042f2e", "#115e59", C["teal_text"]),
}

#: Рівні журналу: (мітка, колір мітки, колір тексту). Мітки — лише візуальні:
#: у скопійований журнал вони не потрапляють.
LOG_TONES = {
    "error": ("ПОМИЛКА", C["rose_soft"], C["rose_text"]),
    "warning": ("УВАГА", C["amber_soft"], C["amber_text"]),
    "success": ("ГОТОВО", C["emerald_soft"], C["emerald_text"]),
    "info": ("ІНФО", C["teal_soft"], C["text_soft"]),
}

#: Рядки порівняння документів: (фон рядка або None, колір тексту, жирний).
DIFF_TONES = {
    "EQUAL": (None, C["text_soft"], False),
    "DELETED": ("#3b0d16", C["rose_text"], False),
    "INSERTED": ("#052e1b", "#86efac", False),
    "MODIFIED": ("#3a2806", C["amber_text"], False),
    "HEADER": (None, C["blue_text"], True),
}

_FONT_FILES = (
    "segoeui.ttf",
    "segoeuib.ttf",
    "seguisb.ttf",
    "consola.ttf",
    "consolab.ttf",
)


def load_local_fonts() -> None:
    """Реєструє системні шрифти з локальної теки Windows.

    На звичайному Windows вони й так доступні; реєстрація потрібна для
    offscreen-режиму (тести, знімки), де системних шрифтів Qt не бачить.
    """
    fonts_dir = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
    for name in _FONT_FILES:
        path = os.path.join(fonts_dir, name)
        if os.path.isfile(path):
            QFontDatabase.addApplicationFont(path)


def _palette() -> QPalette:
    palette = QPalette()
    roles = {
        QPalette.ColorRole.Window: C["bg_window"],
        QPalette.ColorRole.WindowText: C["text"],
        QPalette.ColorRole.Base: C["bg_deep"],
        QPalette.ColorRole.AlternateBase: C["bg_raised"],
        QPalette.ColorRole.Text: C["text"],
        QPalette.ColorRole.Button: C["surface_2"],
        QPalette.ColorRole.ButtonText: C["text"],
        QPalette.ColorRole.Highlight: C["indigo"],
        QPalette.ColorRole.HighlightedText: C["text_strong"],
        QPalette.ColorRole.PlaceholderText: C["text_faint"],
        QPalette.ColorRole.ToolTipBase: C["surface_2"],
        QPalette.ColorRole.ToolTipText: C["text"],
        QPalette.ColorRole.Link: C["violet_soft"],
        QPalette.ColorRole.BrightText: C["rose_soft"],
    }
    for role, color in roles.items():
        palette.setColor(role, QColor(color))
    for role in (
        QPalette.ColorRole.WindowText,
        QPalette.ColorRole.Text,
        QPalette.ColorRole.ButtonText,
    ):
        palette.setColor(QPalette.ColorGroup.Disabled, role, QColor(C["text_dim"]))
    return palette


def build_stylesheet() -> str:
    c = C
    return f"""
QWidget {{
    color: {c["text"]};
    font-family: "{SANS}";
    font-size: 9pt;
}}
QMainWindow, QWidget#AppRoot, QDialog, QMessageBox {{
    background: {c["bg_window"]};
}}
QToolTip {{
    background: {c["surface_2"]};
    color: {c["text"]};
    border: 1px solid {c["border_strong"]};
    padding: 4px 6px;
}}

/* ── Шапка ─────────────────────────────────────────────────────────── */
QFrame#TitleBar {{
    background: {c["bg_deep"]};
    border-bottom: 1px solid {c["border_soft"]};
}}
QLabel#AppLogo {{
    background: qlineargradient(x1:0, y1:1, x2:1, y2:0, stop:0 {c["emerald"]}, stop:1 {c["teal_soft"]});
    color: {c["bg_deep"]};
    border-radius: 4px;
    font-weight: 700;
}}
QLabel#AppName {{
    color: {c["text_strong"]};
    font-weight: 600;
    font-size: 10pt;
}}
QLabel#AppSubtitle {{
    color: {c["text_muted"]};
}}
QLabel#Divider {{
    color: {c["border_strong"]};
}}

/* ── Головні вкладки ───────────────────────────────────────────────── */
QTabWidget#MainTabs {{
    background: {c["bg_deep"]};
}}
QTabWidget#MainTabs::pane {{
    border: none;
    border-top: 1px solid {c["border_soft"]};
    background: {c["bg_window"]};
}}
QTabWidget#MainTabs QTabBar {{
    background: {c["bg_deep"]};
}}
QTabWidget#MainTabs QTabBar::tab {{
    background: transparent;
    color: {c["text_muted"]};
    padding: 7px 14px;
    margin: 5px 2px 0 6px;
    border: 1px solid transparent;
    border-bottom: none;
    border-top-left-radius: 4px;
    border-top-right-radius: 4px;
    font-weight: 500;
}}
QTabWidget#MainTabs QTabBar::tab:selected {{
    background: {c["bg_raised"]};
    color: {c["emerald_soft"]};
    border-color: {c["border_strong"]};
    border-bottom: 2px solid {c["emerald_soft"]};
    font-weight: 600;
}}
QTabWidget#MainTabs QTabBar::tab:hover:!selected {{
    color: {c["text"]};
    background: rgba(15, 23, 42, 0.6);
}}
QLabel#RibbonInfo {{
    color: {c["text_faint"]};
    font-family: "{MONO}";
    font-size: 8pt;
    padding: 0 12px;
}}
QScrollArea#TabScroll, QScrollArea#TabScroll > QWidget, QWidget#TabPage {{
    background: {c["bg_window"]};
    border: none;
}}

/* ── Внутрішні вкладки результатів ─────────────────────────────────── */
QTabWidget#ResultTabs::pane {{
    border: 1px solid {c["border_soft"]};
    border-radius: 4px;
    top: -1px;
    background: {c["bg_deep"]};
}}
QTabWidget#ResultTabs QTabBar::tab {{
    background: {c["bg_raised"]};
    color: {c["text_muted"]};
    padding: 5px 12px;
    border: 1px solid {c["border_soft"]};
    border-bottom: none;
    margin-right: 2px;
    border-top-left-radius: 4px;
    border-top-right-radius: 4px;
}}
QTabWidget#ResultTabs QTabBar::tab:selected {{
    background: {c["surface_2"]};
    color: {c["teal_text"]};
    font-weight: 600;
}}

/* ── Картки розділів ───────────────────────────────────────────────── */
QFrame#SectionCard {{
    background: {c["surface"]};
    border: 1px solid {c["border"]};
    border-radius: 6px;
}}
QFrame#SectionCard[variant="results"] {{
    border-color: #134e4a;
}}
QFrame#SectionCard[variant="log"] {{
    background: {c["bg_raised"]};
}}
QFrame#SectionRule {{
    background: {c["border_soft"]};
    border: none;
}}
QLabel#SectionTitle {{
    color: {c["text"]};
    font-size: 8pt;
    font-weight: 700;
    letter-spacing: 0.6px;
}}
QLabel#SectionTitle[accent="emerald"] {{ color: {c["emerald_soft"]}; }}
QLabel#SectionTitle[accent="amber"] {{ color: {c["amber_text"]}; }}
QLabel#SectionDot {{ border-radius: 3px; background: {c["text_muted"]}; }}
QLabel#SectionDot[accent="indigo"] {{ background: {c["indigo_soft"]}; }}
QLabel#SectionDot[accent="emerald"] {{ background: {c["emerald_soft"]}; }}
QLabel#SectionDot[accent="amber"] {{ background: {c["amber_soft"]}; }}
QLabel#SectionDot[accent="teal"] {{ background: {c["teal_soft"]}; }}

QLabel#MutedLabel {{
    color: {c["text_muted"]};
    font-size: 8pt;
}}
QLabel#FieldLabel {{
    color: {c["text_muted"]};
}}
QLabel#PersonTitle {{
    color: {c["text_soft"]};
    font-weight: 600;
    font-size: 8pt;
}}
QLabel#SummaryLabel {{
    color: {c["indigo_text"]};
    font-weight: 600;
}}

/* ── Бейджі ────────────────────────────────────────────────────────── */
QLabel[tone="emerald"], QLabel[tone="amber"], QLabel[tone="indigo"],
QLabel[tone="slate"], QLabel[tone="teal"], QLabel[tone="rose"] {{
    font-family: "{MONO}";
    font-size: 8pt;
    padding: 1px 6px;
    border-radius: 3px;
    border: 1px solid;
}}
QLabel[tone="emerald"] {{ background: #022c22; color: {c["emerald_text"]}; border-color: #065f46; }}
QLabel[tone="amber"] {{ background: #2a1d05; color: {c["amber_text"]}; border-color: #92400e; }}
QLabel[tone="indigo"] {{ background: #1e1b4b; color: {c["indigo_text"]}; border-color: #3730a3; }}
QLabel[tone="slate"] {{ background: {c["surface_2"]}; color: {c["text_soft"]}; border-color: {c["border_strong"]}; }}
QLabel[tone="teal"] {{ background: #042f2e; color: {c["teal_text"]}; border-color: #115e59; }}
QLabel[tone="rose"] {{ background: #2a0a12; color: {c["rose_text"]}; border-color: #881337; }}

/* ── Кнопки ────────────────────────────────────────────────────────── */
QPushButton {{
    background: {c["surface_2"]};
    color: {c["text"]};
    border: 1px solid {c["border_strong"]};
    border-radius: 4px;
    padding: 5px 10px;
    font-weight: 500;
}}
QPushButton:hover {{
    background: #273449;
    border-color: {c["text_dim"]};
}}
QPushButton:pressed {{
    background: {c["bg_raised"]};
}}
QPushButton:focus {{
    border-color: {c["violet_soft"]};
}}
QPushButton:disabled {{
    background: {c["bg_raised"]};
    color: {c["text_dim"]};
    border-color: {c["border_soft"]};
}}
QPushButton[variant="primary"] {{
    background: {c["indigo"]};
    color: white;
    border-color: rgba(99, 102, 241, 0.6);
    font-weight: 600;
}}
QPushButton[variant="primary"]:hover {{ background: {c["indigo_soft"]}; }}
QPushButton[variant="run"] {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 {c["emerald_deep"]}, stop:1 {c["teal_deep"]});
    color: white;
    border: 1px solid rgba(52, 211, 153, 0.35);
    font-weight: 700;
    padding: 6px 14px;
}}
QPushButton[variant="run"]:hover {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 {c["emerald"]}, stop:1 {c["teal"]});
}}
QPushButton[variant="violet"] {{
    background: {c["violet"]};
    color: white;
    border-color: {c["violet_soft"]};
    font-weight: 600;
}}
QPushButton[variant="violet"]:hover {{ background: {c["violet_soft"]}; }}
QPushButton[variant="teal"] {{
    background: {c["teal"]};
    color: {c["bg_deep"]};
    border-color: {c["teal_soft"]};
    font-weight: 700;
}}
QPushButton[variant="teal"]:hover {{ background: {c["teal_soft"]}; }}
QPushButton[variant="amber"] {{
    background: rgba(245, 158, 11, 0.10);
    color: {c["amber_text"]};
    border-color: rgba(245, 158, 11, 0.35);
}}
QPushButton[variant="amber"]:hover {{ background: rgba(245, 158, 11, 0.20); }}
QPushButton[variant="sky"] {{
    background: {c["bg_raised"]};
    color: {c["sky_text"]};
    border-color: rgba(2, 132, 199, 0.45);
}}
QPushButton[variant="sky"]:hover {{ background: {c["surface_2"]}; }}
QPushButton[variant="rose"] {{
    background: transparent;
    color: {c["rose_soft"]};
    border-color: rgba(136, 19, 55, 0.55);
}}
QPushButton[variant="rose"]:hover {{ background: #2a0a12; color: {c["rose_text"]}; }}
QPushButton[variant="blue"] {{
    background: #172554;
    color: {c["blue_text"]};
    border-color: #1e40af;
}}
QPushButton[variant="blue"]:hover {{ background: #1e3a8a; }}
QPushButton[variant="ghost"] {{
    background: transparent;
    color: {c["text_muted"]};
    border-color: transparent;
}}
QPushButton[variant="ghost"]:hover {{
    background: {c["surface_2"]};
    color: {c["text"]};
}}
QPushButton[variant="segment"] {{
    background: {c["bg_raised"]};
    color: {c["text_muted"]};
    border: 1px solid {c["border_strong"]};
    border-radius: 4px;
    padding: 3px 9px;
    font-size: 8pt;
}}
QPushButton[variant="segment"]:checked {{
    background: rgba(79, 70, 229, 0.35);
    color: {c["indigo_text"]};
    border-color: {c["indigo_soft"]};
    font-weight: 600;
}}
QPushButton[variant="primary"]:disabled, QPushButton[variant="run"]:disabled,
QPushButton[variant="violet"]:disabled, QPushButton[variant="teal"]:disabled,
QPushButton[variant="amber"]:disabled, QPushButton[variant="sky"]:disabled,
QPushButton[variant="rose"]:disabled, QPushButton[variant="blue"]:disabled {{
    background: {c["bg_raised"]};
    color: {c["text_dim"]};
    border-color: {c["border_soft"]};
}}

/* ── Поля введення ─────────────────────────────────────────────────── */
QLineEdit, QComboBox {{
    background: {c["bg_deep"]};
    border: 1px solid {c["border_strong"]};
    border-radius: 4px;
    padding: 4px 6px;
    color: {c["text"]};
    selection-background-color: {c["indigo"]};
}}
QLineEdit:focus, QComboBox:focus {{
    border: 1px solid {c["violet_soft"]};
}}
QLineEdit:disabled {{
    color: {c["text_dim"]};
}}
QComboBox QAbstractItemView {{
    background: {c["surface"]};
    border: 1px solid {c["border_strong"]};
    selection-background-color: {c["indigo"]};
}}
QCheckBox {{
    spacing: 6px;
    color: {c["text_soft"]};
}}
/* Рамку прапорця малюємо явно: у темній палітрі Fusion знята галочка
   в списку наказів майже зливалася з фоном. */
QCheckBox::indicator, QTreeView::indicator {{
    width: 13px;
    height: 13px;
    border: 1px solid {c["text_dim"]};
    border-radius: 3px;
    background: {c["bg_deep"]};
}}
QCheckBox::indicator:hover, QTreeView::indicator:hover {{
    border-color: {c["violet_soft"]};
}}
QCheckBox::indicator:checked, QTreeView::indicator:checked {{
    background: {c["teal_deep"]};
    border-color: {c["teal_soft"]};
    image: url("{os.path.join(os.path.dirname(os.path.abspath(__file__)), "icons", "check.svg").replace(os.sep, "/")}");
}}
QCheckBox::indicator:disabled, QTreeView::indicator:disabled {{
    background: {c["bg_raised"]};
    border-color: {c["border_soft"]};
}}

/* ── Таблиці ───────────────────────────────────────────────────────── */
QTreeWidget#DataTable {{
    background: {c["bg_deep"]};
    alternate-background-color: rgba(15, 23, 42, 0.65);
    border: 1px solid {c["border_soft"]};
    border-radius: 4px;
    color: {c["text_soft"]};
    font-size: 8.5pt;
}}
QTreeWidget#DataTable::item {{
    padding: 3px 4px;
    border-bottom: 1px solid rgba(30, 41, 59, 0.7);
}}
QTreeWidget#DataTable::item:hover {{
    background: rgba(30, 41, 59, 0.45);
}}
QTreeWidget#DataTable::item:selected {{
    background: rgba(79, 70, 229, 0.35);
    color: {c["text_strong"]};
}}
QTreeWidget#DataTable QHeaderView::section {{
    background: {c["surface_2"]};
    color: {c["text_soft"]};
    border: none;
    border-bottom: 1px solid {c["border_strong"]};
    border-right: 1px solid {c["border_soft"]};
    padding: 4px 6px;
    font-size: 7.5pt;
    font-weight: 700;
}}
QTreeWidget#DataTable[headerTone="teal"] QHeaderView::section {{
    background: #042f2e;
    color: {c["teal_text"]};
    border-bottom: 1px solid #134e4a;
}}
QTreeWidget#DataTable[headerTone="emerald"] QHeaderView::section {{
    background: #022c22;
    color: {c["emerald_text"]};
    border-bottom: 1px solid #065f46;
}}

/* ── Журнал ────────────────────────────────────────────────────────── */
QPlainTextEdit#LogConsole, QTextEdit#DiffPane {{
    background: {c["bg_deep"]};
    border: 1px solid {c["border_soft"]};
    border-radius: 4px;
    padding: 4px;
    color: {c["text_soft"]};
    selection-background-color: {c["indigo"]};
}}

/* ── Смуга дій і підказки ──────────────────────────────────────────── */
QFrame#ActionStrip {{
    background: {c["surface"]};
    border: 1px solid {c["border"]};
    border-radius: 6px;
}}
QLabel#StripStatus {{
    font-family: "{MONO}";
    font-size: 8pt;
    color: {c["text_muted"]};
}}
QFrame#HintBar {{
    background: rgba(15, 23, 42, 0.9);
    border: 1px solid {c["border_soft"]};
    border-radius: 4px;
}}
QFrame#OrderDropZone {{
    background: rgba(15, 23, 42, 0.9);
    border: 2px dashed {c["border"]};
    border-radius: 6px;
}}
QFrame#OrderDropZone[active="true"] {{
    background: rgba(20, 184, 166, 0.12);
    border: 2px dashed {c["teal_soft"]};
}}
QLabel#DropZoneText {{
    color: {c["text_muted"]};
    font-size: 10pt;
}}
QLabel#HintText {{
    color: {c["text_muted"]};
    font-size: 8pt;
}}

/* ── Рядок стану ───────────────────────────────────────────────────── */
QFrame#StatusStrip {{
    background: {c["bg_deep"]};
    border-top: 1px solid {c["border_soft"]};
}}
QFrame#StatusStrip QLabel {{
    font-size: 8pt;
    color: {c["text_muted"]};
}}
QLabel#StatusDot {{ border-radius: 3px; background: {c["emerald"]}; }}
QLabel#StatusDot[state="busy"] {{ background: {c["amber_soft"]}; }}
QFrame#StatusStrip QLabel#StatusState {{ color: {c["emerald_soft"]}; font-weight: 600; }}
QFrame#StatusStrip QLabel#StatusState[state="busy"] {{ color: {c["amber_text"]}; }}
QFrame#StatusStrip QLabel#StatusTech {{
    font-family: "{MONO}";
    color: {c["text_faint"]};
}}
QProgressBar#SlimProgress {{
    background: {c["surface_2"]};
    border: none;
    border-radius: 3px;
    max-height: 6px;
}}
QProgressBar#SlimProgress::chunk {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 {c["emerald"]}, stop:1 {c["teal_soft"]});
    border-radius: 3px;
}}

/* ── Прокрутка й роздільники ───────────────────────────────────────── */
QScrollBar:vertical {{
    background: {c["bg_raised"]};
    width: 8px;
    margin: 0;
}}
QScrollBar:horizontal {{
    background: {c["bg_raised"]};
    height: 8px;
    margin: 0;
}}
QScrollBar::handle:vertical, QScrollBar::handle:horizontal {{
    background: {c["border_strong"]};
    border-radius: 4px;
    min-height: 24px;
    min-width: 24px;
}}
QScrollBar::handle:hover {{
    background: {c["text_dim"]};
}}
QScrollBar::add-line, QScrollBar::sub-line {{
    width: 0;
    height: 0;
}}
QScrollBar::add-page, QScrollBar::sub-page {{
    background: none;
}}
QSplitter::handle {{
    background: {c["border_soft"]};
}}
"""


def apply_theme(app: QApplication) -> None:
    """Вмикає «Obsidian» для всього застосунку."""
    load_local_fonts()
    app.setStyle("Fusion")
    app.setPalette(_palette())
    app.setFont(QFont(SANS, 9))
    app.setStyleSheet(build_stylesheet())


def repolish(widget: QWidget) -> None:
    """Перечитує стиль після зміни динамічної властивості (тон, стан)."""
    style = widget.style()
    style.unpolish(widget)
    style.polish(widget)
    widget.update()


def use_dark_title_bar(widget: QWidget) -> None:
    """Темна системна рамка вікна на Windows 10/11 (локальний виклик DWM)."""
    if sys.platform != "win32":
        return
    try:
        hwnd = int(widget.winId())
        enabled = ctypes.c_int(1)
        # 20 — DWMWA_USE_IMMERSIVE_DARK_MODE; 19 — те саме у старих збірках.
        for attribute in (20, 19):
            result = ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd, attribute, ctypes.byref(enabled), ctypes.sizeof(enabled)
            )
            if result == 0:
                break
    except Exception:
        pass
