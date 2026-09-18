"""Тестовий індексатор посад у наказах: вікно PySide6 у стилі генератора.

    python index_positions_qt.py                          — вікно
    python index_positions_qt.py --cli НАКАЗИ ПАПКА_ІНДЕКСУ  — без вікна

Усе локально: індекс — SQLite-файл у папці, яку обирає користувач.
"""

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent
SRC = PROJECT_ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))


def run_cli(orders_folder: str, index_folder: str) -> int:
    from nodeautomationtoolkit.order_index import store

    summary = store.build_index(orders_folder, index_folder, log=print)
    print(summary)
    rows = store.load_positions(index_folder)
    print(f"Посад у індексі: {len(rows)} → {store.index_path(index_folder)}")
    return 0


def main() -> int:
    if len(sys.argv) == 4 and sys.argv[1] == "--cli":
        return run_cli(sys.argv[2], sys.argv[3])

    from nodeautomationtoolkit.order_index.window import launch

    return launch(sys.argv)


if __name__ == "__main__":
    raise SystemExit(main())
