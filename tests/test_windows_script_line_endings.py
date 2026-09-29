"""Windows-скрипти й модулі VBA — лише з CRLF (AGENT.md, розд. 8.5).

З LF-закінченнями cmd розсипає рядки `.bat` на уламки («'usebackq' is not
recognized…», «'ined' is not recognized…») і ламає мітки `goto` — збірка
`build_generator_qt.bat` падала саме так. Word з LF-файлу `.bas` імпортує
модуль зламаним. `.gitattributes` видає їх із CRLF, а тест ловить файл, який
перезаписав редактор чи скрипт із LF.
"""

from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SKIP_DIRS = {".git", ".venv", "python", "build", "dist", "node_modules"}


def _windows_scripts() -> list[Path]:
    found = []
    for pattern in ("*.bat", "*.cmd", "*.bas"):
        for path in PROJECT_ROOT.rglob(pattern):
            relative = path.relative_to(PROJECT_ROOT)
            if relative.parts and relative.parts[0] in SKIP_DIRS:
                continue
            if ".claude" in relative.parts:
                continue
            found.append(path)
    return sorted(found)


def test_there_are_windows_scripts_to_check():
    names = {path.name for path in _windows_scripts()}
    assert {"build_generator_qt.bat", "start_generator_qt.bat"} <= names


@pytest.mark.parametrize(
    "path", _windows_scripts(), ids=lambda path: str(path.relative_to(PROJECT_ROOT))
)
def test_windows_script_uses_crlf(path):
    data = path.read_bytes()
    bare_lf = data.count(b"\n") - data.count(b"\r\n")
    assert bare_lf == 0, f"{path.name}: {bare_lf} рядків із LF замість CRLF"
