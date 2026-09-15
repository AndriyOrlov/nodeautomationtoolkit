@echo off
chcp 65001 >nul
title Generator Vytyagiv (Qt)
cd /d "%~dp0"

rem New PySide6 shell over the same generator logic.
rem The classic Tk window is still available via start_generator.bat.
rem Interpreter lookup order:
rem   1) portable  - python\python.exe next to the program (nothing to install)
rem   2) .venv     - developer virtual environment
rem   3) system    - python from PATH
rem Comments are kept ASCII on purpose: Cyrillic inside .bat code breaks
rem depending on the file encoding and the console codepage.

set "PY="
if exist "python\python.exe" set "PY=python\python.exe"
if not defined PY if exist ".venv\Scripts\python.exe" set "PY=.venv\Scripts\python.exe"

if defined PY goto :run

where python >nul 2>nul
if errorlevel 1 goto :nopython
set "PY=python"

:run
echo Iнтерпретатор: %PY%
"%PY%" generate_extracts_qt.py
if errorlevel 1 pause
exit /b 0

:nopython
echo.
echo Python не знайдено.
echo.
echo Покладiть portable-Python у теку "python" поруч iз цим файлом.
echo Iнструкцiя: scripts\portable\README.md
echo.
pause
exit /b 1
