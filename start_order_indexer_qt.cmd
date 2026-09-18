@echo off
chcp 65001 >nul
title Order position indexer (Qt, test)
cd /d "%~dp0"

rem Test indexer of positions in orders (PySide6, same theme as the generator).
rem Interpreter lookup order:
rem   1) portable  - python\python.exe next to the program
rem   2) .venv     - developer virtual environment
rem   3) .venv of the main checkout (when started from .claude\worktrees\...)
rem   4) system    - python from PATH
rem Comments are kept ASCII on purpose: Cyrillic inside .cmd code breaks
rem depending on the file encoding and the console codepage.

set "PY="
if exist "python\python.exe" set "PY=python\python.exe"
if not defined PY if exist ".venv\Scripts\python.exe" set "PY=.venv\Scripts\python.exe"
if not defined PY if exist "..\..\..\.venv\Scripts\python.exe" set "PY=..\..\..\.venv\Scripts\python.exe"

if defined PY goto :check

where python >nul 2>nul
if errorlevel 1 goto :nopython
set "PY=python"

:check
"%PY%" -c "import PySide6" >nul 2>nul
if errorlevel 1 goto :nopyside

echo Iнтерпретатор: %PY%
"%PY%" index_positions_qt.py %*
if errorlevel 1 pause
exit /b 0

:nopyside
echo.
echo У %PY% не встановлено PySide6.
echo Встановiть: "%PY%" -m pip install "PySide6>=6.7,<7"
echo.
pause
exit /b 1

:nopython
echo.
echo Python не знайдено.
echo.
echo Покладiть portable-Python у теку "python" поруч iз цим файлом.
echo Iнструкцiя: scripts\portable\README.md
echo.
pause
exit /b 1
