@echo off
chcp 65001 >nul
title Build Povidomlennia (one exe)
cd /d "%~dp0"

rem Builds the standalone messages program (generate_messages_qt.py: only the
rem messages page, no copies / extracts / order generator) into ONE file:
rem   dist\PovidomlenniaKadrovi.exe
rem Everything the program needs (Python, PySide6, pywin32, icons) is inside
rem the exe. Excel, templates and config_messages.json stay separate files:
rem the config is written next to the exe.
rem After the build the exe is signed (Authenticode, SHA-256) by
rem scripts\signing\sign_exe.ps1 if a code-signing certificate is available.
rem Spec and temp files go to build\pyinstaller-messages (ignored by git).
rem Interpreter lookup order:
rem   1) .venv     - developer virtual environment
rem   2) portable  - python\python.exe next to the program
rem   3) system    - python from PATH
rem Comments are kept ASCII on purpose: Cyrillic inside .bat code breaks
rem depending on the file encoding and the console codepage.

rem In GitHub Actions (CI=true) nobody presses a key: skip pause there.
set "PAUSE=pause"
if defined CI set "PAUSE=rem"

set "APP_NAME=PovidomlenniaKadrovi"
set "ROOT=%~dp0"
set "WORK=%ROOT%build\pyinstaller-messages"
set "ICONS=%ROOT%src\nodeautomationtoolkit\generator_qt\icons"

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

if not exist "%WORK%" mkdir "%WORK%"
"%PY%" scripts\messages_exe\make_version_info.py "%WORK%\version_info.txt" >nul
if errorlevel 1 goto :noversion

echo.
echo Збираю %APP_NAME%.exe (один файл) ...
echo.

rem --onefile: one exe. --noupx: packed exe files are flagged by antivirus
rem far more often. --add-data / --paths use absolute paths: with --specpath
rem PyInstaller resolves relative ones against the spec folder.
"%PY%" -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --onefile ^
  --windowed ^
  --noupx ^
  --name "%APP_NAME%" ^
  --icon "%ICONS%\messages_logo.ico" ^
  --version-file "%WORK%\version_info.txt" ^
  --distpath "%ROOT%dist" ^
  --workpath "%WORK%" ^
  --specpath "%WORK%" ^
  --paths "%ROOT%src" ^
  --collect-submodules nodeautomationtoolkit.builtin_nodes ^
  --collect-submodules nodeautomationtoolkit.generator_qt ^
  --collect-submodules nodeautomationtoolkit.order_generator ^
  --collect-submodules nodeautomationtoolkit.order_index ^
  --collect-submodules nodeautomationtoolkit.order_review ^
  --collect-all docx ^
  --hidden-import win32timezone ^
  --add-data "%ICONS%;nodeautomationtoolkit\generator_qt\icons" ^
  --add-data "%ROOT%src\nodeautomationtoolkit\personnel\dictionaries;nodeautomationtoolkit\personnel\dictionaries" ^
  --add-data "%ROOT%src\nodeautomationtoolkit\generator_qt\instruction.md;nodeautomationtoolkit\generator_qt" ^
  --add-data "%ROOT%src\nodeautomationtoolkit\order_index\positions_all.csv;nodeautomationtoolkit\order_index" ^
  --add-data "%ROOT%src\nodeautomationtoolkit\order_index\unit_markers.txt;nodeautomationtoolkit\order_index" ^
  "%ROOT%generate_messages_qt.py"
if errorlevel 1 goto :failed
if not exist "dist\%APP_NAME%.exe" goto :failed

echo.
echo Пiдписую %APP_NAME%.exe ...
powershell -NoProfile -ExecutionPolicy Bypass -File "%ROOT%scripts\signing\sign_exe.ps1" -Path "%ROOT%dist\%APP_NAME%.exe"
set "SIGN_RESULT=%errorlevel%"

echo.
echo Готово: dist\%APP_NAME%.exe  (один файл, без теки _internal)
if "%SIGN_RESULT%"=="0" echo Пiдпис: поставлено.
if "%SIGN_RESULT%"=="2" echo Пiдпис: НЕМАЄ сертифiката - див. scripts\signing\README.md
if not "%SIGN_RESULT%"=="0" if not "%SIGN_RESULT%"=="2" echo Пiдпис: НЕ вдався - дивiться повiдомлення вище.
echo Налаштування програма пише в config_messages.json поруч iз exe.
echo.
%PAUSE%
exit /b 0

:failed
echo.
echo Збiрка не вдалася - дивiться повiдомлення вище.
echo.
%PAUSE%
exit /b 1

:noversion
echo.
echo Не вдалося прочитати версiю - scripts\messages_exe\make_version_info.py.
echo.
%PAUSE%
exit /b 1

:noscript
echo.
echo Поруч немає generate_messages_qt.py.
echo Покладiть цей файл у корiнь проєкту, поруч iз start_messages_qt.bat.
echo.
%PAUSE%
exit /b 1

:nopyinstaller
echo.
echo PyInstaller не встановлено для %PY%.
echo Встановiть командою:
echo   "%PY%" -m pip install pyinstaller
echo.
%PAUSE%
exit /b 1

:nopython
echo.
echo Python не знайдено.
echo.
echo Покладiть portable-Python у теку "python" поруч iз цим файлом.
echo Iнструкцiя: scripts\portable\README.md
echo.
%PAUSE%
exit /b 1
