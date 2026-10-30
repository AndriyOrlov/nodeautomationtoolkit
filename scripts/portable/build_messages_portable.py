"""Сумісність: те саме, що `build_portable.py messages` → `dist\\Povidomlennya_bat\\`."""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from build_portable import build  # noqa: E402

if __name__ == "__main__":
    raise SystemExit(build("messages"))
