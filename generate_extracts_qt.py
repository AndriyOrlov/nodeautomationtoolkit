"""Запуск генератора в новій Qt-оболонці (PySide6, тема «Obsidian»).

Логіка генерації — та сама, що в `generate_extracts.py`: оболонка лише
замінює інтерфейс. Стара Tk-версія лишається доступною через
`start_generator.bat`.
"""

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))


def main() -> int:
    import generate_extracts as legacy  # сам додає src у sys.path

    from nodeautomationtoolkit.generator_qt.main_window import launch

    return launch(legacy, sys.argv)


if __name__ == "__main__":
    raise SystemExit(main())
