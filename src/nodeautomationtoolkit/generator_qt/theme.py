"""Тема Fluent: світла й темна палітри, перемикання на льоту, шрифти.

Кнопки, поля, навігацію й діалоги малює PySide6-Fluent-Widgets (`fluent.py`).
Тут — лише те, що наше власне: картки-таблиці, бейджі, журнал, рядок стану,
рамка перетягування, порівняння документів. Кольори лежать у словниках `C`,
`PILL_TONES`, `LOG_TONES`, `DIFF_TONES`, `WORD_TONES`: під час зміни теми вони
оновлюються НА МІСЦІ, тож код, що читає `theme.C[...]`, завжди бачить поточну.

Усе локальне: шрифти із системної теки Windows, жодних мережевих запитів
(PROJECT_RULES 1.1).
"""

from __future__ import annotations

import ctypes
import os
import sys
from typing import Callable

from PySide6.QtGui import QColor, QFont, QFontDatabase, QPalette
from PySide6.QtWidgets import QApplication, QWidget

from .fluent import Theme, setTheme, setThemeColor

SANS = "Segoe UI"
MONO = "Consolas"

LIGHT, DARK = "light", "dark"
#: Акцент Fluent — бірюзовий, як і в попередній темі «Obsidian».
ACCENT = "#009FAA"

_COMMON_KEYS = (
    "bg_deep", "bg_window", "bg_raised", "surface", "surface_2", "border", "border_soft",
    "border_strong", "text_strong", "text", "text_soft", "text_muted", "text_faint", "text_dim",
)

PALETTES: dict[str, dict[str, str]] = {
    DARK: {
        "bg_deep": "#1b1b1b",
        "bg_window": "#202020",
        "bg_raised": "#272727",
        "surface": "#2b2b2b",
        "surface_2": "#333333",
        "border": "#3a3a3a",
        "border_soft": "#303030",
        "border_strong": "#4a4a4a",
        "text_strong": "#ffffff",
        "text": "#e8e8e8",
        "text_soft": "#d0d0d0",
        "text_muted": "#a8a8a8",
        "text_faint": "#8b8b8b",
        "text_dim": "#6f6f6f",
        "emerald": "#6ccb5f",
        "emerald_soft": "#7fd473",
        "emerald_deep": "#3f9a33",
        "emerald_text": "#a5e39d",
        "emerald_bg": "#1e3a22",
        "teal": "#2fb5bd",
        "teal_soft": "#4cc9d0",
        "teal_deep": "#1a8f96",
        "teal_text": "#8fe3e8",
        "indigo": "#7b83eb",
        "indigo_soft": "#8f96f0",
        "indigo_text": "#c0c4f7",
        "violet": "#a78bfa",
        "violet_soft": "#b9a2fb",
        "violet_text": "#d6c9fd",
        "amber": "#f2c661",
        "amber_soft": "#f5d17f",
        "amber_text": "#f8de9f",
        "sky": "#60cdff",
        "sky_soft": "#7dd6ff",
        "sky_text": "#a8e3ff",
        "rose": "#ff99a4",
        "rose_soft": "#ffadb6",
        "rose_text": "#ffc4cb",
        "blue": "#60cdff",
        "blue_text": "#99dcff",
    },
    LIGHT: {
        "bg_deep": "#ffffff",
        "bg_window": "#f3f3f3",
        "bg_raised": "#f9f9f9",
        "surface": "#ffffff",
        "surface_2": "#f0f0f0",
        "border": "#e0e0e0",
        "border_soft": "#ebebeb",
        "border_strong": "#c8c8c8",
        "text_strong": "#1a1a1a",
        "text": "#242424",
        "text_soft": "#3b3b3b",
        "text_muted": "#5c5c5c",
        "text_faint": "#7a7a7a",
        "text_dim": "#9a9a9a",
        "emerald": "#107c10",
        "emerald_soft": "#13841a",
        "emerald_deep": "#0b5a0b",
        "emerald_text": "#0e6b0e",
        "emerald_bg": "#dff6dd",
        "teal": "#00807a",
        "teal_soft": "#007f86",
        "teal_deep": "#006b66",
        "teal_text": "#005f5a",
        "indigo": "#4f46e5",
        "indigo_soft": "#4f46e5",
        "indigo_text": "#3730a3",
        "violet": "#6d28d9",
        "violet_soft": "#7c3aed",
        "violet_text": "#5b21b6",
        "amber": "#9a5b00",
        "amber_soft": "#a86400",
        "amber_text": "#7a4800",
        "sky": "#0369a1",
        "sky_soft": "#0277b5",
        "sky_text": "#075985",
        "rose": "#c42b1c",
        "rose_soft": "#c42b1c",
        "rose_text": "#a4262c",
        "blue": "#005fb8",
        "blue_text": "#004c99",
    },
}

#: Пігулки в таблицях: (фон, рамка, текст).
_PILLS = {
    DARK: {
        "emerald": ("#1e3a22", "#2f5a2a", "#a5e39d"),
        "rose": ("#3d1d22", "#6b2c35", "#ffc4cb"),
        "indigo": ("#272a4a", "#3d4380", "#c0c4f7"),
        "slate": ("#333333", "#4a4a4a", "#d0d0d0"),
        "amber": ("#3d3115", "#6b5520", "#f8de9f"),
        "teal": ("#163a3c", "#22595c", "#8fe3e8"),
    },
    LIGHT: {
        "emerald": ("#dff6dd", "#a7dba3", "#0e6b0e"),
        "rose": ("#fde7e9", "#f1b1b7", "#a4262c"),
        "indigo": ("#e8e8fc", "#bfc2f5", "#3730a3"),
        "slate": ("#f0f0f0", "#d1d1d1", "#3b3b3b"),
        "amber": ("#fff4ce", "#f0d283", "#7a4800"),
        "teal": ("#dcf3f3", "#9ed8d9", "#005f5a"),
    },
}

#: Рядки порівняння документів: (фон рядка або None, колір тексту, жирний).
_DIFFS = {
    DARK: {
        "EQUAL": (None, "#d0d0d0", False),
        "DELETED": ("#3d1d22", "#ffc4cb", False),
        "INSERTED": ("#1e3a22", "#a5e39d", False),
        "MODIFIED": ("#3d3115", "#f8de9f", False),
        "HEADER": (None, "#99dcff", True),
    },
    LIGHT: {
        "EQUAL": (None, "#3b3b3b", False),
        "DELETED": ("#fde7e9", "#a4262c", False),
        "INSERTED": ("#dff6dd", "#0e6b0e", False),
        "MODIFIED": ("#fff4ce", "#7a4800", False),
        "HEADER": (None, "#004c99", True),
    },
}
#: Слова в зміненому рядку порівняння: розбіжність — жовте, інша частина (шифр) — червоне.
_WORDS = {
    DARK: {"changed": ("#6b5520", "#fff1c7", False), "unit": ("#7f1d2d", "#ffe4e6", True)},
    LIGHT: {"changed": ("#ffe08a", "#5c3c00", False), "unit": ("#f1707b", "#ffffff", True)},
}

C: dict[str, str] = dict(PALETTES[DARK])
PILL_TONES: dict[str, tuple] = {}
LOG_TONES: dict[str, tuple] = {}
DIFF_TONES: dict[str, tuple] = {}
WORD_TONES: dict[str, tuple] = {}

_mode = DARK
_listeners: list[Callable[[], None]] = []


def _fill_tones(mode: str) -> None:
    C.clear()
    C.update(PALETTES[mode])
    PILL_TONES.clear()
    PILL_TONES.update(_PILLS[mode])
    #: Рівні журналу: (мітка, колір мітки, колір тексту). Мітки — лише візуальні:
    #: у скопійований журнал вони не потрапляють.
    LOG_TONES.clear()
    LOG_TONES.update(
        {
            "error": ("ПОМИЛКА", C["rose_soft"], C["rose_text"]),
            "warning": ("УВАГА", C["amber_soft"], C["amber_text"]),
            "success": ("ГОТОВО", C["emerald_soft"], C["emerald_text"]),
            "info": ("ІНФО", C["teal_soft"], C["text_soft"]),
        }
    )
    DIFF_TONES.clear()
    DIFF_TONES.update(_DIFFS[mode])
    WORD_TONES.clear()
    WORD_TONES.update(_WORDS[mode])


_fill_tones(DARK)


def current_mode() -> str:
    return _mode


def is_dark() -> bool:
    return _mode == DARK


def normalize_mode(mode: str | None) -> str:
    return LIGHT if str(mode or "").strip().casefold() == LIGHT else DARK


def on_change(callback: Callable[[], None]) -> Callable[[], None]:
    """Викликати `callback()` після кожної зміни теми (перемалювати HTML, рядки)."""
    _listeners.append(callback)
    return callback


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
    """Палітра для звичайних (не Fluent) віджетів: діалоги, таблиці, текст."""
    palette = QPalette()
    roles = {
        QPalette.ColorRole.Window: C["bg_window"],
        QPalette.ColorRole.WindowText: C["text"],
        QPalette.ColorRole.Base: C["surface"],
        QPalette.ColorRole.AlternateBase: C["bg_raised"],
        QPalette.ColorRole.Text: C["text"],
        QPalette.ColorRole.Button: C["surface_2"],
        QPalette.ColorRole.ButtonText: C["text"],
        QPalette.ColorRole.Highlight: ACCENT,
        QPalette.ColorRole.HighlightedText: "#ffffff",
        QPalette.ColorRole.PlaceholderText: C["text_faint"],
        QPalette.ColorRole.ToolTipBase: C["surface_2"],
        QPalette.ColorRole.ToolTipText: C["text"],
        QPalette.ColorRole.Link: C["blue"],
        QPalette.ColorRole.BrightText: C["rose_soft"],
    }
    for role, color in roles.items():
        palette.setColor(role, QColor(color))
    for role in (QPalette.ColorRole.WindowText, QPalette.ColorRole.Text, QPalette.ColorRole.ButtonText):
        palette.setColor(QPalette.ColorGroup.Disabled, role, QColor(C["text_dim"]))
    return palette


def build_stylesheet() -> str:
    """Таблиця стилів лише для наших елементів; Fluent-віджети стилізуються самі."""
    c = C
    tones = "\n".join(
        f'QLabel[tone="{tone}"] {{ background: {bg}; border: 1px solid {border}; color: {fg}; }}'
        for tone, (bg, border, fg) in PILL_TONES.items()
    )
    return f"""
QWidget#AppRoot {{
    background: {c["bg_window"]};
    color: {c["text"]};
}}
QScrollArea#TabScroll, QScrollArea#TabScroll > QWidget > QWidget#TabPage, QWidget#TabPage {{
    background: transparent;
    border: none;
}}
QLabel[tone] {{
    border-radius: 4px;
    padding: 1px 8px;
    font-size: 8pt;
}}
{tones}
QFrame#ActionStrip {{
    background: transparent;
    border: none;
}}
QFrame#HintBar {{
    background: transparent;
    border: 1px dashed {c["border_strong"]};
    border-radius: 6px;
}}
QFrame#OrderDropZone {{
    background: {c["bg_raised"]};
    border: 1px dashed {c["border_strong"]};
    border-radius: 8px;
}}
QFrame#OrderDropZone[active="true"] {{
    border: 1px solid {ACCENT};
    background: {c["surface_2"]};
}}
QFrame#StatusStrip {{
    background: transparent;
    border-top: 1px solid {c["border_soft"]};
}}
QLabel#StatusDot {{
    border-radius: 3px;
    background: {c["emerald"]};
}}
QLabel#StatusDot[state="busy"] {{
    background: {c["amber"]};
}}
QLabel#StatusState {{
    color: {c["text"]};
}}
QLabel#StatusState[state="busy"] {{
    color: {c["amber_text"]};
}}
QTreeWidget#DataTable {{
    background: {c["surface"]};
    alternate-background-color: {c["bg_raised"]};
    border: 1px solid {c["border"]};
    border-radius: 6px;
    color: {c["text"]};
    selection-background-color: {c["surface_2"]};
    selection-color: {c["text_strong"]};
    outline: 0;
}}
QTreeWidget#DataTable::item {{
    padding: 3px 4px;
    border: none;
}}
QTreeWidget#DataTable::item:selected {{
    background: {c["surface_2"]};
    color: {c["text_strong"]};
}}
QHeaderView::section {{
    background: {c["bg_raised"]};
    color: {c["text_muted"]};
    border: none;
    border-bottom: 1px solid {c["border"]};
    padding: 5px 6px;
    font-size: 8pt;
    font-weight: 600;
}}
QPlainTextEdit#LogConsole, QTextBrowser#InstructionText, QTextEdit#DiffPane {{
    background: {c["bg_deep"]};
    color: {c["text"]};
    border: 1px solid {c["border"]};
    border-radius: 6px;
    selection-background-color: {ACCENT};
}}
QTabWidget#ResultTabs::pane {{
    border: none;
    border-top: 1px solid {c["border"]};
}}
QTabWidget#ResultTabs > QTabBar::tab {{
    background: transparent;
    color: {c["text_muted"]};
    padding: 6px 12px;
    border: none;
    border-bottom: 2px solid transparent;
}}
QTabWidget#ResultTabs > QTabBar::tab:selected {{
    color: {c["text_strong"]};
    border-bottom: 2px solid {ACCENT};
}}
QTabWidget#ResultTabs > QTabBar::tab:hover {{
    color: {c["text_strong"]};
}}
QToolTip {{
    background: {c["surface_2"]};
    color: {c["text"]};
    border: 1px solid {c["border_strong"]};
}}
QScrollBar:vertical {{
    background: transparent;
    width: 10px;
    margin: 2px;
}}
QScrollBar:horizontal {{
    background: transparent;
    height: 10px;
    margin: 2px;
}}
QScrollBar::handle {{
    background: {c["border_strong"]};
    border-radius: 3px;
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


def apply_theme(app: QApplication, mode: str | None = None) -> str:
    """Вмикає світлу чи темну тему для всього застосунку; повертає назву теми."""
    global _mode
    _mode = normalize_mode(mode if mode is not None else _mode)
    _fill_tones(_mode)
    load_local_fonts()
    # Стиль ставимо ОДИН раз: `setStyle` видаляє попередній об'єкт стилю, а кнопки
    # Fluent тримають на нього вказівник (`setStyle(QApplication.style())`) —
    # повторна заміна під час перемикання теми валила б малювання.
    # Тому — лише поки жодного віджета ще немає (запуск програми).
    if app.style().name().casefold() != "fusion" and not app.allWidgets():
        app.setStyle("Fusion")
    app.setPalette(_palette())
    app.setFont(QFont(SANS, 9))
    setThemeColor(ACCENT)
    setTheme(Theme.DARK if _mode == DARK else Theme.LIGHT)
    app.setStyleSheet(build_stylesheet())
    for widget in app.topLevelWidgets():
        if widget.isWindow() and widget.isVisible():
            use_dark_title_bar(widget)
    for callback in tuple(_listeners):
        try:
            callback()
        except RuntimeError:
            # Віджет слухача вже знищено (закрите вікно) — прибираємо його.
            _listeners.remove(callback)
    return _mode


def repolish(widget: QWidget) -> None:
    """Перечитує стиль після зміни динамічної властивості (тон, стан)."""
    style = widget.style()
    style.unpolish(widget)
    style.polish(widget)
    widget.update()


def use_dark_title_bar(widget: QWidget) -> None:
    """Системна рамка вікна в тон теми (Windows 10/11, локальний виклик DWM).

    Fluent-вікна мають власну рамку; це для звичайних діалогів і вікон.
    """
    if sys.platform != "win32":
        return
    try:
        hwnd = int(widget.winId())
        enabled = ctypes.c_int(1 if is_dark() else 0)
        # 20 — DWMWA_USE_IMMERSIVE_DARK_MODE; 19 — те саме у старих збірках.
        for attribute in (20, 19):
            result = ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd, attribute, ctypes.byref(enabled), ctypes.sizeof(enabled)
            )
            if result == 0:
                break
    except Exception:
        pass
