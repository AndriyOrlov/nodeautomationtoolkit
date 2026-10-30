"""Програма БЕЗ власного EXE — запуск через .bat (для ПК з «Інтелектуальним керуванням програмами»).

Windows 11 (Smart App Control) блокує непідписаний EXE, зібраний PyInstaller:
файл унікальний, видавця немає. Безкоштовного сертифіката, якому довіряє
Windows, для закритої програми не існує. Тому тут — тека, де програму запускає
ОФІЦІЙНИЙ `pythonw.exe` (підписаний Python Software Foundation), а Qt-бібліотеки
підписані The Qt Company. Наш код — звичайні `.py` (+ заздалегідь скомпільований
байткод), які Windows як програми не перевіряє.

    .venv\\Scripts\\python.exe scripts\\portable\\build_portable.py main
    .venv\\Scripts\\python.exe scripts\\portable\\build_portable.py messages

Викликається й наприкінці `build_generator_qt.bat` / `build_messages_qt.bat`
(30.10.2026: «у білд обох програм — на всяк випадок запуск через .bat без EXE»).
pip потребує інтернету лише тут, під час збирання; програма лишається офлайн.

Результат: `dist\\GeneratorVytyagivQt_bat\\` або `dist\\Povidomlennya_bat\\` —
переносити ВСЮ теку, запуск — `.bat` у ній. Налаштування — поруч (`config.json` /
`messages_config.json`).
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
SOURCE_PYTHON = ROOT / "python"  # embeddable Python 3.14 x64 (див. README.md поруч)

APPS = {
    "main": {
        "folder": "GeneratorVytyagivQt_bat",
        "bat": "GeneratorVytyagivQt.bat",
        "script": "generate_extracts_qt.py",
        "title": "Генератор витягів, примірників і повідомлень",
        "extra_dirs": ["word_macro"],
    },
    "messages": {
        "folder": "Povidomlennya_bat",
        "bat": "Povidomlennya.bat",
        "script": "generate_messages_qt.py",
        "title": "Повідомлення за наказом",
        "extra_dirs": [],
    },
}

#: PySide6-Essentials замість повного PySide6 (без Addons) — у кілька разів менше.
#: Fluent і рамка вікна — без залежностей, бо вони тягнуть повний PySide6.
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


def launcher(script: str) -> str:
    return (
        "@echo off\r\n"
        "rem Program without an own EXE: the signed official pythonw.exe runs the\r\n"
        "rem program sources (Smart App Control blocks an unsigned EXE).\r\n"
        "set \"PYTHONNOUSERSITE=1\"\r\n"
        f"start \"\" \"%~dp0python\\pythonw.exe\" \"%~dp0{script}\"\r\n"
    )


def readme(app: dict) -> str:
    return f"""{app["title"]} — версія без власного EXE

Запуск: {app["bat"]} (двічі клацнути).
Переносити треба ВСЮ теку. Потрібен лише Microsoft Word.

Чому так: «Інтелектуальне керування програмами» Windows 11 блокує
непідписаний EXE. Тут програму запускає офіційний Python
(python\\pythonw.exe, підписаний Python Software Foundation).

Якщо Windows заблокує сам .bat — створіть ярлик:
  Об'єкт:  <ця тека>\\python\\pythonw.exe "<ця тека>\\{app["script"]}"
  Робоча тека: <ця тека>
Якщо заблокує якийсь інший файл — надішліть знімок вікна: у ньому видно, який.
"""


def _ignore(directory: str, names: list[str]) -> set[str]:
    return {name for name in names if name in ("__pycache__", ".pytest_cache") or name.endswith(".pyc")}


def build(app_key: str) -> int:
    app = APPS[app_key]
    target = ROOT / "dist" / app["folder"]
    if not (SOURCE_PYTHON / "python.exe").is_file():
        print(f"Немає портативного Python: {SOURCE_PYTHON}")
        return 1
    if target.exists():
        shutil.rmtree(target)
    target.mkdir(parents=True)

    # 1. Інтерпретатор — без чужих бібліотек: site-packages збираємо наново.
    shutil.copytree(
        SOURCE_PYTHON,
        target / "python",
        ignore=lambda d, names: _ignore(d, names)
        | ({"site-packages"} if Path(d).name == "Lib" else set()),
    )
    site_packages = target / "python" / "Lib" / "site-packages"
    site_packages.mkdir(parents=True, exist_ok=True)

    # 2. Бібліотеки — у теку пакета, НЕ з профілю користувача (PYTHONNOUSERSITE).
    python = str(target / "python" / "python.exe")
    env = dict(os.environ, PYTHONNOUSERSITE="1")
    pip = [sys.executable, "-m", "pip", "install", "--quiet", "--disable-pip-version-check",
           "--target", str(site_packages), "--python-version", "3.14", "--only-binary=:all:",
           "--platform", "win_amd64", "--implementation", "cp"]
    subprocess.run(pip + PACKAGES, check=True, env=env)
    subprocess.run(pip + ["--no-deps"] + NO_DEPS_PACKAGES, check=True, env=env)

    # 3. Код програми (+ потрібні теки поруч).
    for name in (app["script"], "generate_extracts.py"):
        shutil.copy2(ROOT / name, target / name)
    shutil.copytree(ROOT / "src" / "nodeautomationtoolkit", target / "src" / "nodeautomationtoolkit", ignore=_ignore)
    for folder in app["extra_dirs"]:
        if (ROOT / folder).is_dir():
            shutil.copytree(ROOT / folder, target / folder, ignore=_ignore)

    # Байткод заздалегідь — Python-ом САМОГО пакета (той самий тег версії). Без цього
    # перший запуск на ПК компілює всі модулі, а в теці без права запису — щоразу.
    subprocess.run(
        [python, "-m", "compileall", "-q", "-j", "0",
         "src", app["script"], "generate_extracts.py", str(site_packages)],
        cwd=target, env=env, check=False,
    )

    (target / app["bat"]).write_bytes(launcher(app["script"]).encode("ascii"))
    (target / "ПРОЧИТАЙ.txt").write_text(readme(app), encoding="utf-8-sig")

    # 4. Самоперевірка: усе імпортується саме з теки пакета.
    check = (
        "import contextlib, io, sys; sys.path.insert(0, 'src');\n"
        "import PySide6, win32com.client, openpyxl, docx, PIL\n"
        "with contextlib.redirect_stdout(io.StringIO()): import qfluentwidgets\n"
        f"assert {app['folder']!r} in PySide6.__file__, PySide6.__file__\n"
        "print('OK', PySide6.__version__)"
    )
    subprocess.run([python, "-c", check], check=True, env=env, cwd=target)
    size = sum(f.stat().st_size for f in target.rglob("*") if f.is_file()) / 1e6
    print(f"Done: {target}  ({size:.0f} MB)")
    return 0


def main(argv: list[str]) -> int:
    if len(argv) != 1 or argv[0] not in APPS:
        print("usage: build_portable.py main|messages")
        return 2
    return build(argv[0])


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
