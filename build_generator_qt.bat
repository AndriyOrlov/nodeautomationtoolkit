@echo off
chcp 65001 >nul
title Build Generator Vytyagiv (Qt)
cd /d "%~dp0"

rem Builds the PySide6 generator (generate_extracts_qt.py, started by
rem start_generator_qt.bat) into a portable folder with PyInstaller:
rem   dist\GeneratorVytyagivQt\GeneratorVytyagivQt.exe
rem Spec and temp files go to build\pyinstaller (ignored by git).
rem Interpreter lookup order:
rem   1) .venv     - developer virtual environment
rem   2) portable  - python\python.exe next to the program
rem   3) system    - python from PATH
rem Comments are kept ASCII on purpose: Cyrillic inside .bat code breaks
rem depending on the file encoding and the console codepage.

set "APP_NAME=GeneratorVytyagivQt"
set "ROOT=%~dp0"

if not exist "generate_extracts_qt.py" goto :noscript

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

rem Version shown in the program goes up by 0.1 (0.4.0 -> 0.5.0) before every
rem build. If the build fails, the previous version is written back.
set "OLD_VERSION="
set "NEW_VERSION="
for /f "usebackq delims=" %%V in (`"%PY%" scripts\bump_version.py --show`) do set "OLD_VERSION=%%V"
if not defined OLD_VERSION goto :noversion
for /f "usebackq delims=" %%V in (`"%PY%" scripts\bump_version.py`) do set "NEW_VERSION=%%V"
if not defined NEW_VERSION goto :noversion
echo Версiя: %OLD_VERSION% -^> %NEW_VERSION%

echo.
echo Збираю %APP_NAME% %NEW_VERSION% ...
echo.

rem --add-data / --paths use absolute paths: with --specpath PyInstaller
rem resolves relative ones against the spec folder, not the project root.
"%PY%" -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --windowed ^
  --name "%APP_NAME%" ^
  --distpath "%ROOT%dist" ^
  --workpath "%ROOT%build\pyinstaller" ^
  --specpath "%ROOT%build\pyinstaller" ^
  --paths "%ROOT%src" ^
  --collect-submodules nodeautomationtoolkit.builtin_nodes ^
  --collect-submodules nodeautomationtoolkit.generator_qt ^
  --collect-all docx ^
  --collect-all ttkbootstrap ^
  --hidden-import win32timezone ^
  --add-data "%ROOT%src\nodeautomationtoolkit\generator_qt\icons;nodeautomationtoolkit\generator_qt\icons" ^
  --add-data "%ROOT%src\nodeautomationtoolkit\personnel\dictionaries;nodeautomationtoolkit\personnel\dictionaries" ^
  "%ROOT%generate_extracts_qt.py"
if errorlevel 1 goto :failed

if not exist "dist\%APP_NAME%\%APP_NAME%.exe" goto :failed

echo.
echo Готово: dist\%APP_NAME%\%APP_NAME%.exe  (версiя %NEW_VERSION%)
echo Запускати треба разом з усiєю текою dist\%APP_NAME%.
echo.
pause
exit /b 0

:failed
if defined OLD_VERSION "%PY%" scripts\bump_version.py --set %OLD_VERSION% >nul
echo.
echo Збiрка не вдалася - дивiться повiдомлення вище.
if defined OLD_VERSION echo Версiю повернуто на %OLD_VERSION%.
echo.
pause
exit /b 1

:noversion
echo.
echo Не вдалося прочитати або змiнити версiю - scripts\bump_version.py.
echo.
pause
exit /b 1

:noscript
echo.
echo Поруч немає generate_extracts_qt.py.
echo Покладiть цей файл у корiнь проєкту, поруч iз start_generator_qt.bat.
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
echo Покладiть portable-Python у теку "python" поруч iз цим файлом.
echo Iнструкцiя: scripts\portable\README.md
echo.
pause
exit /b 1
