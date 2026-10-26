"""Окрема програма: лише повідомлення про прийняття кадрових рішень.

Та сама сторінка «Повідомлення», що й у генераторі, без примірників і витягів.
Логіка — `generate_extracts.py`, вікно — `nodeautomationtoolkit.generator_qt.messages_window`.
Запуск із коду — `start_messages_qt.bat`, один EXE — `build_messages_qt.bat`.
"""

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))


def main() -> int:
    import generate_extracts as legacy  # сам додає src у sys.path
    from nodeautomationtoolkit.generator_qt.messages_window import launch_messages

    return launch_messages(legacy, sys.argv)


if __name__ == "__main__":
    raise SystemExit(main())
