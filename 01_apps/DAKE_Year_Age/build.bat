@echo off
setlocal
del /q version_info.txt 2>nul

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

%PYTHON_CMD% ..\..\tools\generate_version_info.py --app-dir . --out version_info.txt
if errorlevel 1 (
    echo VersionInfo generation failed.
    pause
    exit /b 1
)

%PYTHON_CMD% -m PyInstaller ^
--onefile ^
--noconsole ^
--clean ^
--name DakeYear_Age ^
--paths=..\..\00_core ^
--icon=..\..\02_assets\dake_icon.ico ^
--version-file version_info.txt ^
--add-data "..\..\02_assets\dake_icon.ico;." ^
main.py

pause
