"""Тихий імпорт PySide6-Fluent-Widgets (GPLv3, https://github.com/zhiyiYo/PyQt-Fluent-Widgets).

Бібліотека під час імпорту друкує в консоль рекламу своєї Pro-версії. Програма
працює без мережі, а консольний режим (макрос Word) читає вивід — тож рядок
глушимо. Решта модулів оболонки імпортує Fluent лише звідси.
"""

from __future__ import annotations

import contextlib
import io

with contextlib.redirect_stdout(io.StringIO()):
    import qfluentwidgets  # noqa: F401

from qfluentwidgets import (  # noqa: E402,F401
    BodyLabel,
    CaptionLabel,
    CheckBox,
    FlowLayout,
    FluentIcon,
    FluentWindow,
    LineEdit,
    MessageBox,
    NavigationItemPosition,
    PlainTextEdit,
    PrimaryPushButton,
    ProgressBar,
    PushButton,
    SimpleCardWidget,
    StrongBodyLabel,
    SubtitleLabel,
    Theme,
    TitleLabel,
    ToolButton,
    TogglePushButton,
    TransparentPushButton,
    isDarkTheme,
    setTheme,
    setThemeColor,
)
