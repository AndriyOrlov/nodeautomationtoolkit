@echo off
chcp 65001 >nul
echo Запуск Генератора Витягів та Примірників...
cd /d "%~dp0"
if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" generate_extracts.py
) else (
    python generate_extracts.py
)
if errorlevel 1 pause
