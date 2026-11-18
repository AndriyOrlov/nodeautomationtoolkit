@echo off
title Build .bat versions (no own exe)
cd /d "%~dp0"

rem Builds the versions WITHOUT an own exe (start via .bat): the signed official
rem pythonw.exe runs the sources - for PCs where Smart App Control blocks the exe.
rem Not part of build_generator_qt.bat / build_messages_qt.bat any more (06.11.2026).
rem   build_bat_versions.bat           - both programs
rem   build_bat_versions.bat main      - only the main program
rem   build_bat_versions.bat messages  - only Povidomlennya
rem Result: dist\GeneratorVytyagivQt_bat\ and dist\Povidomlennya_bat\
rem ASCII only on purpose: Cyrillic inside .bat breaks with the console codepage.

set "ROOT=%~dp0"
set "PY="
if exist ".venv\Scripts\python.exe" set "PY=.venv\Scripts\python.exe"
if not defined PY if exist "python\python.exe" set "PY=python\python.exe"
if defined PY goto :run

where python >nul 2>nul
if errorlevel 1 goto :nopython
set "PY=python"

:run
set "WHAT=%~1"
if "%WHAT%"=="" set "WHAT=main messages"
set "FAILED="
for %%K in (%WHAT%) do (
    echo.
    echo Building the .bat version: %%K ...
    "%PY%" "%ROOT%scripts\portable\build_portable.py" %%K
    if errorlevel 1 set "FAILED=1"
)

echo.
if defined FAILED (
    echo WARNING: some .bat version was not built - see the messages above.
    pause
    exit /b 1
)
echo Done: dist\GeneratorVytyagivQt_bat\ and/or dist\Povidomlennya_bat\
pause
exit /b 0

:nopython
echo.
echo Python not found (.venv, python\python.exe or PATH).
echo.
pause
exit /b 1
