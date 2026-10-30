"""Запуск генератора (PySide6, тема «Obsidian»).

Логіка генерації — у `generate_extracts.py`, інтерфейс — у
`nodeautomationtoolkit.generator_qt`. Tk-версії більше немає.
"""

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))


def main() -> int:
    # Макрос Word «Перевірити програмою»: перевірка без вікна (word_macro/NATProgramCheck.bas).
    if "--review-text" in sys.argv:
        import generate_extracts as legacy  # сам додає src у sys.path

        return legacy.run_review_cli(sys.argv[1:])

    splash = _splash("Генератор витягів, примірників і повідомлень")
    import generate_extracts as legacy  # сам додає src у sys.path

    from nodeautomationtoolkit.generator_qt.main_window import launch

    return launch(legacy, sys.argv, splash=splash)


def _splash(title: str):
    """Заставка одразу, ще до імпорту логіки й Fluent (найдовша частина запуску)."""
    if str(PROJECT_ROOT / "src") not in sys.path:
        sys.path.insert(0, str(PROJECT_ROOT / "src"))
    try:
        from PySide6.QtWidgets import QApplication

        from nodeautomationtoolkit.generator_qt.splash import show_splash

        app = QApplication.instance() or QApplication(list(sys.argv))
        return show_splash(app, title)
    except Exception:
        return None  # без заставки програма працює так само


if __name__ == "__main__":
    raise SystemExit(main())
