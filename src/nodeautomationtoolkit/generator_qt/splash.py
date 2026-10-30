"""Заставка під час запуску: з'являється одразу, поки вантажаться модулі й будується вікно.

Картинка — `assets/splash.png` (720×400; якщо поруч є `splash@2x.png` — для
екранів високої щільності). Її немає — заставку малює код: фон темної теми,
бірюзова смуга акценту, назва програми, версія. Назву, версію й «Завантаження…»
програма пише сама в нижній частині (нижні ~80 px картинки мають бути рівними).

Тут свідомо лише PySide6: заставка має з'явитися ДО імпорту Fluent і логіки.
"""

from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QRectF, Qt
from PySide6.QtGui import QColor, QFont, QPainter, QPixmap
from PySide6.QtWidgets import QApplication, QSplashScreen

ASSETS = Path(__file__).resolve().parent / "assets"
WIDTH, HEIGHT = 720, 400
_BACKGROUND = QColor("#202020")
_ACCENT = QColor("#009FAA")
_TEXT = QColor("#ffffff")
_MUTED = QColor("#a8a8a8")


def _version() -> str:
    try:
        import nodeautomationtoolkit

        return str(getattr(nodeautomationtoolkit, "__version__", "") or "")
    except Exception:
        return ""


def _base_pixmap(ratio: float) -> QPixmap:
    for name in ("splash@2x.png", "splash.png") if ratio > 1.25 else ("splash.png", "splash@2x.png"):
        path = ASSETS / name
        if path.is_file():
            picture = QPixmap(str(path))
            if not picture.isNull():
                picture = picture.scaled(
                    int(WIDTH * ratio), int(HEIGHT * ratio),
                    Qt.AspectRatioMode.KeepAspectRatioByExpanding,
                    Qt.TransformationMode.SmoothTransformation,
                )
                picture.setDevicePixelRatio(ratio)
                return picture
    picture = QPixmap(int(WIDTH * ratio), int(HEIGHT * ratio))
    picture.setDevicePixelRatio(ratio)
    picture.fill(_BACKGROUND)
    painter = QPainter(picture)
    painter.fillRect(QRectF(0, 0, 6, HEIGHT), _ACCENT)
    painter.end()
    return picture


#: Назва на заставці (користувач, 30.10.2026): «NAT» і «Automation Toolkit»;
#: під ними — яка саме програма й версія.
BRAND = "NAT"
BRAND_LINE = "Automation Toolkit"


def build_pixmap(title: str, subtitle: str = "", ratio: float = 1.0) -> QPixmap:
    """`title` — назва програми; на заставці не показується (лише бренд і версія)."""
    picture = _base_pixmap(ratio)
    painter = QPainter(picture)
    painter.setRenderHint(QPainter.RenderHint.TextAntialiasing)
    left = Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
    painter.setPen(_TEXT)
    painter.setFont(QFont("Segoe UI", 40, QFont.Weight.Bold))
    painter.drawText(QRectF(36, HEIGHT - 236, WIDTH - 72, 64), left, BRAND)
    painter.setPen(_ACCENT)
    painter.setFont(QFont("Segoe UI", 16, QFont.Weight.DemiBold))
    painter.drawText(QRectF(36, HEIGHT - 172, WIDTH - 72, 32), left, BRAND_LINE)
    # Назву програми (`title`) на заставці не пишемо — користувач прибрав її
    # 30.10.2026; лишаються бренд, версія й «Завантаження…».
    painter.setPen(_MUTED)
    painter.setFont(QFont("Segoe UI", 10))
    version = _version()
    line = "  ·  ".join(part for part in (subtitle, f"версія {version}" if version else "") if part)
    painter.drawText(QRectF(36, HEIGHT - 136, WIDTH - 72, 24), left, line)
    painter.setFont(QFont("Segoe UI", 9))
    painter.drawText(QRectF(36, HEIGHT - 52, WIDTH - 72, 24), Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter, "Завантаження…")
    painter.end()
    return picture


def _load_fonts_offscreen(app: QApplication) -> None:
    """Без екрана (тести, знімки) Qt не бачить системних шрифтів — реєструємо Segoe UI."""
    if app.platformName() != "offscreen":
        return
    import os

    from PySide6.QtGui import QFontDatabase

    for name in ("segoeui.ttf", "seguisb.ttf"):
        path = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts", name)
        if os.path.isfile(path):
            QFontDatabase.addApplicationFont(path)


def show_splash(app: QApplication, title: str, subtitle: str = "") -> QSplashScreen:
    """Показати заставку й одразу намалювати її (до важких імпортів)."""
    _load_fonts_offscreen(app)
    # Стиль застосунку ставиться лише поки немає жодного віджета (theme.apply_theme),
    # а заставка — перший віджет. Тож Fusion — тут, до неї.
    if app.style().name().casefold() != "fusion":
        app.setStyle("Fusion")
    screen = app.primaryScreen()
    ratio = screen.devicePixelRatio() if screen is not None else 1.0
    splash = QSplashScreen(build_pixmap(title, subtitle, ratio), Qt.WindowType.WindowStaysOnTopHint)
    splash.show()
    app.processEvents()
    return splash
