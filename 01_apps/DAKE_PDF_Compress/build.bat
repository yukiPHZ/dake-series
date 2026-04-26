@echo off
setlocal

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist *.spec del /q *.spec

set "PYTHON_CMD=python"
where python >nul 2>&1
if errorlevel 1 (
    where py >nul 2>&1
    if errorlevel 1 (
        echo Python was not found. Please install Python and run this file again.
        pause
        exit /b 1
    )
    set "PYTHON_CMD=py -3"
)

set "DND_OPTION="
%PYTHON_CMD% -c "import tkinterdnd2" >nul 2>&1
if %errorlevel%==0 set "DND_OPTION=--collect-data=tkinterdnd2"

%PYTHON_CMD% -m PyInstaller ^
--onefile ^
--noconsole ^
--clean ^
--icon=..\..\02_assets\dake_icon.ico ^
%DND_OPTION% ^
--name DakePDF_Compress ^
main.py

pause
