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
    import generate_extracts as legacy  # сам додає src у sys.path

    # Макрос Word «Перевірити програмою»: перевірка без вікна (word_macro/NATProgramCheck.bas).
    if "--review-text" in sys.argv:
        return legacy.run_review_cli(sys.argv[1:])

    from nodeautomationtoolkit.generator_qt.main_window import launch

    return launch(legacy, sys.argv)


if __name__ == "__main__":
    raise SystemExit(main())
