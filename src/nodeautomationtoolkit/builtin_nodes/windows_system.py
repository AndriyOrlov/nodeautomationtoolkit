from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path
from typing import Any

from nodeautomationtoolkit.core.definition import node


@node(
    name="Запустити команду Windows",
    category="Система Windows",
    description="Виконує системну команду PowerShell або CMD у Windows та повертає результат.",
    type_id="builtin.windows.run_command",
    outputs={
        "stdout": "str",
        "stderr": "str",
        "exit_code": "int",
        "summary": "str",
    },
    execution_inputs=("exec",),
    execution_outputs=("then",),
    preview_policy="never",
)
def run_command(
    command: str = "",
    use_powershell: bool = True,
    working_dir: str = "",
) -> dict:
    if not command.strip():
        raise ValueError("Вкажіть команду для виконання")

    cwd = Path(working_dir).expanduser() if working_dir.strip() else None
    if cwd and not cwd.is_dir():
        cwd = None

    if use_powershell:
        cmd_args = [
            "powershell.exe",
            "-NoProfile",
            "-NonInteractive",
            "-ExecutionPolicy",
            "Bypass",
            "-Command",
            command,
        ]
    else:
        cmd_args = ["cmd.exe", "/c", command]

    process = subprocess.run(
        cmd_args,
        capture_output=True,
        text=True,
        cwd=cwd,
        encoding="utf-8",
        errors="replace",
    )

    stdout = process.stdout.strip()
    stderr = process.stderr.strip()
    exit_code = process.returncode

    status = "Успішно" if exit_code == 0 else f"Помилка ({exit_code})"
    summary = f"Виконано: {status} · {command[:60]}"

    return {
        "stdout": stdout,
        "stderr": stderr,
        "exit_code": exit_code,
        "summary": summary,
    }


@node(
    name="Відкрити файл чи програму",
    category="Система Windows",
    description="Відкриває файл, папку, програму або URL у Windows за допомогою стандартного обробника.",
    type_id="builtin.windows.open_path",
    outputs={
        "summary": "str",
    },
    execution_inputs=("exec",),
    execution_outputs=("then",),
    preview_policy="never",
)
def open_path(path_or_url: str = "") -> dict:
    target = path_or_url.strip()
    if not target:
        raise ValueError("Вкажіть шлях або посилання для відкриття")

    if target.startswith(("http://", "https://")):
        import webbrowser

        webbrowser.open(target)
        summary = f"Відкрито посилання: {target}"
    else:
        path = Path(target).expanduser()
        if hasattr(os, "startfile"):
            os.startfile(str(path))  # noqa: S606
            summary = f"Відкрито об'єкт Windows: {path.name or target}"
        else:
            subprocess.run(["cmd.exe", "/c", "start", "", str(path)])
            summary = f"Запущено: {target}"

    return {"summary": summary}


@node(
    name="Буфер обміну Windows",
    category="Система Windows",
    description="Записує або зчитує текст із системного буфера обміну Windows.",
    type_id="builtin.windows.clipboard",
    outputs={
        "text": "str",
        "summary": "str",
    },
    execution_inputs=("exec",),
    execution_outputs=("then",),
    preview_policy="never",
)
def clipboard(text: str = "", mode: str = "write") -> dict:
    from PySide6.QtGui import QGuiApplication
    from PySide6.QtWidgets import QApplication

    app = QApplication.instance()
    created_app = False
    if app is None:
        app = QApplication(sys.argv)
        created_app = True

    cb = QGuiApplication.clipboard()
    if mode.lower() == "write" or text.strip():
        cb.setText(text)
        result_text = text
        summary = f"Текст скопійовано у буфер обміну ({len(text)} симв.)"
    else:
        result_text = cb.text()
        summary = f"Зчитано з буфера обміну ({len(result_text)} симв.)"

    return {"text": result_text, "summary": summary}


@node(
    name="Сповіщення Windows",
    category="Система Windows",
    description="Надсилає системне сповіщення Windows у центр повідомлень.",
    type_id="builtin.windows.send_notification",
    outputs={
        "summary": "str",
    },
    execution_inputs=("exec",),
    execution_outputs=("then",),
    preview_policy="never",
)
def send_notification(title: str = "Node Automation Toolkit", message: str = "") -> dict:
    if not message.strip():
        message = "Сценарій успішно виконано!"

    ps_script = (
        f"[reflection.assembly]::loadwithpartialname('System.Windows.Forms');"
        f"$notify = new-object system.windows.forms.notifyicon;"
        f"$notify.icon = [system.drawing.systemicons]::information;"
        f"$notify.visible = $true;"
        f"$notify.showballoontip(5000, '{title.replace("'", "''")}', '{message.replace("'", "''")}', [system.windows.forms.tooltipicon]::info);"
    )

    try:
        subprocess.run(
            ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", ps_script],
            capture_output=True,
            timeout=3,
        )
    except Exception:
        pass

    return {"summary": f"Сповіщення надіслано: {title} — {message}"}


@node(
    name="Упакувати у standalone-пакет",
    category="Проєкт",
    description="Упаковує граф дій у самостійний виконуваний пакет (.bat + .json) для передачі на інші комп'ютери.",
    type_id="builtin.project.export_standalone",
    outputs={
        "package_dir": "str",
        "launcher_bat": "str",
        "summary": "str",
    },
    execution_inputs=("exec",),
    execution_outputs=("then",),
    preview_policy="never",
)
def export_standalone(
    graph_json_path: str = "",
    output_folder: str = "",
    package_name: str = "Моя_Автоматизація",
) -> dict:
    if not graph_json_path.strip() or not Path(graph_json_path).is_file():
        raise FileNotFoundError(f"Файл сценарію не знайдено: {graph_json_path}")

    out_base = Path(output_folder).expanduser() if output_folder.strip() else Path(graph_json_path).parent / "Standalone_Packages"
    pkg_dir = out_base / package_name.strip()
    pkg_dir.mkdir(parents=True, exist_ok=True)

    target_json = pkg_dir / "scenario.nat.json"
    target_json.write_bytes(Path(graph_json_path).read_bytes())

    bat_content = f"""@echo off
chcp 65001 > nul
title Виконання сценарію {package_name}
echo ===================================================
echo  Запуск автоматизації Node Automation Toolkit
echo ===================================================
echo.

if exist "%~dp0NodeAutomationToolkit.exe" (
    "%~dp0NodeAutomationToolkit.exe" --run "%~dp0scenario.nat.json"
    goto END
)

python -m nodeautomationtoolkit --run "%~dp0scenario.nat.json"

:END
echo.
echo ===================================================
echo  Виконання завершено.
pause
"""
    bat_file = pkg_dir / f"Запустити_{package_name}.bat"
    bat_file.write_text(bat_content, encoding="utf-8")

    summary = f"Автономний пакет створено у {pkg_dir.name} (лаунчер: {bat_file.name})"
    return {
        "package_dir": str(pkg_dir),
        "launcher_bat": str(bat_file),
        "summary": summary,
    }
