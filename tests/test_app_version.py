"""Номер версії — те, за чим користувач бачить, яка збірка запущена.

Домовленість 18.09.2026: кожна зміна в програмі збільшує останнє число на
одиницю. Тест стежить, щоб номер не розʼїхався між місцями, де він записаний,
і щоб він реально доходив до заголовка вікна.
"""

import re
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

import nodeautomationtoolkit


def _pyproject_version() -> str:
    text = (PROJECT_ROOT / "pyproject.toml").read_text(encoding="utf-8")
    match = re.search(r'^version = "([^"]+)"', text, re.MULTILINE)
    assert match, "у pyproject.toml немає рядка version"
    return match.group(1)


def test_version_is_three_numbers():
    assert re.fullmatch(r"\d+\.\d+\.\d+", nodeautomationtoolkit.__version__)


def test_pyproject_and_package_agree():
    """Два місця з номером мають збігатися, інакше у вікні буде не те число."""
    assert nodeautomationtoolkit.__version__ == _pyproject_version()


def test_window_title_shows_the_version():
    source = (PROJECT_ROOT / "generate_extracts.py").read_text(encoding="utf-8")
    assert "from nodeautomationtoolkit import __version__ as APP_VERSION" in source
    assert "{APP_VERSION}" in source
