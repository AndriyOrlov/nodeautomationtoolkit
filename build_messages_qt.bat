@echo off
chcp 65001 >nul
title Build Povidomlennya (Qt)
cd /d "%~dp0"

rem Builds the stand-alone messages program (generate_messages_qt.py) into ONE
rem file with PyInstaller:
rem   dist\Povidomlennya.exe
rem Settings are kept in messages_config.json next to the exe.
rem Spec and temp files go to build\pyinstaller_messages (ignored by git).
rem Interpreter lookup order:
rem   1) .venv     - developer virtual environment
rem   2) portable  - python\python.exe next to the program
rem   3) system    - python from PATH
rem Comments are kept ASCII on purpose: Cyrillic inside .bat code breaks
rem depending on the file encoding and the console codepage.

set "APP_NAME=Povidomlennya"
set "ROOT=%~dp0"

if not exist "generate_messages_qt.py" goto :noscript

set "PY="
if exist ".venv\Scripts\python.exe" set "PY=.venv\Scripts\python.exe"
if not defined PY if exist "python\python.exe" set "PY=python\python.exe"

if defined PY goto :check

where python >nul 2>nul
if errorlevel 1 goto :nopython
set "PY=python"

:check
echo Iнтерпретатор: %PY%
"%PY%" -m PyInstaller --version >nul 2>nul
if errorlevel 1 goto :nopyinstaller

echo.
echo Збираю %APP_NAME% ...
echo.

rem --add-data / --paths use absolute paths: with --specpath PyInstaller
rem resolves relative ones against the spec folder, not the project root.
"%PY%" -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --onefile ^
  --windowed ^
  --name "%APP_NAME%" ^
  --distpath "%ROOT%dist" ^
  --workpath "%ROOT%build\pyinstaller_messages" ^
  --specpath "%ROOT%build\pyinstaller_messages" ^
  --paths "%ROOT%src" ^
  --collect-submodules nodeautomationtoolkit.builtin_nodes ^
  --collect-submodules nodeautomationtoolkit.generator_qt ^
  --collect-submodules nodeautomationtoolkit.order_review ^
  --collect-all docx ^
  --hidden-import win32timezone ^
  --add-data "%ROOT%src\nodeautomationtoolkit\generator_qt\icons;nodeautomationtoolkit\generator_qt\icons" ^
  --add-data "%ROOT%src\nodeautomationtoolkit\personnel\dictionaries;nodeautomationtoolkit\personnel\dictionaries" ^
  --add-data "%ROOT%src\nodeautomationtoolkit\order_index\positions_all.csv;nodeautomationtoolkit\order_index" ^
  --add-data "%ROOT%src\nodeautomationtoolkit\order_index\unit_markers.txt;nodeautomationtoolkit\order_index" ^
  "%ROOT%generate_messages_qt.py"
if errorlevel 1 goto :failed

if not exist "dist\%APP_NAME%.exe" goto :failed

echo.
echo Готово: dist\%APP_NAME%.exe
echo Це один файл - його можна копiювати куди завгодно.
echo.
pause
exit /b 0

:failed
echo.
echo Збiрка не вдалася - дивiться повiдомлення вище.
echo.
pause
exit /b 1

:noscript
echo.
echo Поруч немає generate_messages_qt.py.
echo.
pause
exit /b 1

:nopyinstaller
echo.
echo PyInstaller не встановлено для %PY%.
echo Встановiть командою:
echo   "%PY%" -m pip install pyinstaller
echo.
pause
exit /b 1

:nopython
echo.
echo Python не знайдено.
echo.
pause
exit /b 1
