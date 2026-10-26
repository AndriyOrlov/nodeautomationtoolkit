"""Запуск окремої програми «Повідомлення за наказом» (PySide6).

Логіка — у `generate_extracts.py`, вікно — у
`nodeautomationtoolkit.generator_qt.messages_window`. Збирається в EXE
файлом `build_messages_qt.bat`.
"""

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))


def main() -> int:
    import generate_extracts as legacy  # сам додає src у sys.path

    from nodeautomationtoolkit.generator_qt.messages_window import launch

    return launch(legacy, sys.argv)


if __name__ == "__main__":
    raise SystemExit(main())
