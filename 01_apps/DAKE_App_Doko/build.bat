@echo off
chcp 65001 > nul
cd /d "%~dp0"

rmdir /s /q build 2>nul
rmdir /s /q dist 2>nul
del /q *.spec 2>nul
del /q version_info.txt 2>nul

set "PYTHON_CMD=python"
where python >nul 2>&1
if errorlevel 1 (
  set "PYTHON_CMD=%LocalAppData%\Programs\Python\Python312\python.exe"
  if not exist "%PYTHON_CMD%" (
    echo Python was not found. Please install Python and run this file again.
    pause
    exit /b 1
  )
)

"%PYTHON_CMD%" ..\..\tools\generate_version_info.py --app-dir . --out version_info.txt
if errorlevel 1 (
  echo VersionInfo generation failed.
  pause
  exit /b 1
)

"%PYTHON_CMD%" -m PyInstaller ^
--onefile ^
--noconsole ^
--clean ^
--paths=..\..\00_core ^
--icon=..\..\02_assets\dake_icon.ico ^
--version-file version_info.txt ^
--exclude-module pandas ^
--exclude-module numpy ^
--exclude-module PIL ^
--exclude-module openpyxl ^
--exclude-module lxml ^
--exclude-module matplotlib ^
--exclude-module scipy ^
--name DakeApp_Doko ^
main.py

pause
