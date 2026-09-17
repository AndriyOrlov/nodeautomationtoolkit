"""Підвищує версію програми на 0.1 перед збіркою (викликається з build_generator_qt.bat).

Версія живе у двох місцях і має збігатися:
- `src/nodeautomationtoolkit/__init__.py` — `__version__`, її показує вікно генератора;
- `pyproject.toml` — `version` пакета.

«+0.1» — друге число версії: 0.4.0 → 0.5.0, 0.9.0 → 0.10.0 (третє скидається в 0).

Використання:
    python scripts/bump_version.py            підвищити, вивести нову версію
    python scripts/bump_version.py --set X    записати версію X (відкат після збірки, що не вдалася)
    python scripts/bump_version.py --show     вивести поточну версію
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
INIT_FILE = ROOT / "src" / "nodeautomationtoolkit" / "__init__.py"
PYPROJECT_FILE = ROOT / "pyproject.toml"

_INIT_RE = re.compile(r'^(__version__\s*=\s*")([^"]+)(")', re.MULTILINE)
_PYPROJECT_RE = re.compile(r'^(version\s*=\s*")([^"]+)(")', re.MULTILINE)
_VERSION_RE = re.compile(r"^\d+\.\d+\.\d+$")


def read_version() -> str:
    match = _INIT_RE.search(INIT_FILE.read_text(encoding="utf-8"))
    if not match:
        raise SystemExit(f"Не знайдено __version__ у {INIT_FILE}")
    return match.group(2)


def next_version(version: str) -> str:
    if not _VERSION_RE.match(version):
        raise SystemExit(f"Версія «{version}» не у форматі X.Y.Z")
    major, minor, _patch = (int(part) for part in version.split("."))
    return f"{major}.{minor + 1}.0"


def _replace(path: Path, pattern: re.Pattern, version: str) -> None:
    text = path.read_text(encoding="utf-8")
    new_text, count = pattern.subn(rf"\g<1>{version}\g<3>", text, count=1)
    if count != 1:
        raise SystemExit(f"Не знайдено рядок версії у {path}")
    path.write_text(new_text, encoding="utf-8", newline="")


def write_version(version: str) -> None:
    if not _VERSION_RE.match(version):
        raise SystemExit(f"Версія «{version}» не у форматі X.Y.Z")
    _replace(INIT_FILE, _INIT_RE, version)
    _replace(PYPROJECT_FILE, _PYPROJECT_RE, version)


def main(argv: list[str]) -> int:
    if argv[:1] == ["--show"]:
        print(read_version())
        return 0
    if argv[:1] == ["--set"] and len(argv) == 2:
        write_version(argv[1])
        print(argv[1])
        return 0
    if argv:
        print(__doc__)
        return 2
    version = next_version(read_version())
    write_version(version)
    print(version)
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
