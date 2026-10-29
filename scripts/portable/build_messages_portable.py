"""Програма «Повідомлення» БЕЗ власного EXE — для ПК з «Інтелектуальним керуванням програмами».

Windows 11 (Smart App Control) блокує непідписаний `Povidomlennya.exe`: файл,
зібраний PyInstaller, унікальний, видавця немає. Безкоштовного сертифіката,
якому довіряє Windows, для закритої програми не існує. Тому тут — тека, де
програму запускає ОФІЦІЙНИЙ `pythonw.exe` (підписаний Python Software
Foundation), а Qt-бібліотеки підписані The Qt Company. Наш код — звичайні
`.py`, які Windows як програми не перевіряє.

Збирання (на ПК розробника, потрібен інтернет для pip — лише тут):

    .venv\\Scripts\\python.exe scripts\\portable\\build_messages_portable.py

Результат: `dist\\Povidomlennya_portable\\` — переносити ВСЮ теку, запуск
`Povidomlennya.bat`. Конфіг — `messages_config.json` у цій же теці.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
SOURCE_PYTHON = ROOT / "python"  # embeddable Python 3.14 x64 (див. README.md поруч)
TARGET = ROOT / "dist" / "Povidomlennya_portable"

#: Лише те, що потрібно програмі повідомлень. PySide6-Essentials замість повного
#: PySide6 (без Addons) — у кілька разів менше. Fluent і рамка вікна — без
#: залежностей, бо вони тягнуть повний PySide6.
PACKAGES = [
    "PySide6-Essentials==6.11.1",
    "shiboken6==6.11.1",
    "pywin32==312",
    "openpyxl==3.1.5",
    "python-docx==1.2.0",
    "lxml",
    "pillow",
    "darkdetect==0.8.0",
]
NO_DEPS_PACKAGES = [
    "PySide6-Fluent-Widgets==1.11.3",
    "PySideSix-Frameless-Window==0.8.2",
]

LAUNCHER = (
    "@echo off\r\n"
    "rem Messages app without an own EXE: the signed official pythonw.exe runs the\r\n"
    "rem program sources (Smart App Control blocks an unsigned EXE).\r\n"
    "set \"PYTHONNOUSERSITE=1\"\r\n"
    "start \"\" \"%~dp0python\\pythonw.exe\" \"%~dp0generate_messages_qt.py\"\r\n"
)

README = """Повідомлення за наказом — версія без власного EXE

Запуск: Povidomlennya.bat (двічі клацнути).
Переносити треба ВСЮ теку. Потрібен лише Microsoft Word.

Чому так: «Інтелектуальне керування програмами» Windows 11 блокує
непідписаний Povidomlennya.exe. Тут програму запускає офіційний Python
(python\\pythonw.exe, підписаний Python Software Foundation).

Якщо Windows заблокує сам .bat — створіть ярлик:
  Об'єкт:  <ця тека>\\python\\pythonw.exe "<ця тека>\\generate_messages_qt.py"
  Робоча тека: <ця тека>
Якщо заблокує якийсь інший файл — надішліть знімок вікна: у ньому видно, який.
"""


def _ignore(directory: str, names: list[str]) -> set[str]:
    return {name for name in names if name in ("__pycache__", ".pytest_cache") or name.endswith(".pyc")}


def main() -> int:
    if not (SOURCE_PYTHON / "python.exe").is_file():
        print(f"Немає портативного Python: {SOURCE_PYTHON}")
        return 1
    if TARGET.exists():
        shutil.rmtree(TARGET)
    TARGET.mkdir(parents=True)

    # 1. Інтерпретатор — без чужих бібліотек: site-packages збираємо наново.
    shutil.copytree(
        SOURCE_PYTHON,
        TARGET / "python",
        ignore=lambda d, names: _ignore(d, names)
        | ({"site-packages"} if Path(d).name == "Lib" else set()),
    )
    site_packages = TARGET / "python" / "Lib" / "site-packages"
    site_packages.mkdir(parents=True, exist_ok=True)

    # 2. Бібліотеки — у теку пакета, НЕ з профілю користувача (PYTHONNOUSERSITE).
    python = str(TARGET / "python" / "python.exe")
    env = dict(os.environ, PYTHONNOUSERSITE="1")
    pip = [sys.executable, "-m", "pip", "install", "--quiet", "--disable-pip-version-check",
           "--target", str(site_packages), "--python-version", "3.14", "--only-binary=:all:",
           "--platform", "win_amd64", "--implementation", "cp"]
    subprocess.run(pip + PACKAGES, check=True, env=env)
    subprocess.run(pip + ["--no-deps"] + NO_DEPS_PACKAGES, check=True, env=env)

    # 3. Код програми.
    for name in ("generate_messages_qt.py", "generate_extracts.py"):
        shutil.copy2(ROOT / name, TARGET / name)
    shutil.copytree(ROOT / "src" / "nodeautomationtoolkit", TARGET / "src" / "nodeautomationtoolkit", ignore=_ignore)

    (TARGET / "Povidomlennya.bat").write_bytes(LAUNCHER.encode("ascii"))
    (TARGET / "ПРОЧИТАЙ.txt").write_text(README, encoding="utf-8-sig")

    # 4. Самоперевірка: усе імпортується саме з теки пакета.
    check = (
        "import contextlib, io, sys; sys.path.insert(0, 'src');\n"
        "import PySide6, win32com.client, openpyxl, docx, PIL\n"
        "with contextlib.redirect_stdout(io.StringIO()): import qfluentwidgets\n"
        "assert 'Povidomlennya_portable' in PySide6.__file__, PySide6.__file__\n"
        "print('OK', PySide6.__version__)"
    )
    subprocess.run([python, "-c", check], check=True, env=env, cwd=TARGET)
    size = sum(f.stat().st_size for f in TARGET.rglob("*") if f.is_file()) / 1e6
    print(f"Готово: {TARGET}  ({size:.0f} МБ)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
