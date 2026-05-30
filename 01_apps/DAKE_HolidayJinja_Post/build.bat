@echo off
cd /d "%~dp0"
del /q version_info.txt 2>nul

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist Dake_HolidayJinja_Post.spec del Dake_HolidayJinja_Post.spec

set "PYTHON_EXE=python"
if exist "%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" (
  set "PYTHON_EXE=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
)

python ..\..\tools\generate_version_info.py --app-dir . --out version_info.txt
if errorlevel 1 (
    echo VersionInfo generation failed.
    pause
    exit /b 1
)

"%PYTHON_EXE%" -m PyInstaller ^
--onefile ^
--noconsole ^
--clean ^
--collect-all customtkinter ^
--paths=..\..\00_core ^
--icon=..\..\02_assets\dake_icon.ico ^
--version-file version_info.txt ^
--name Dake_HolidayJinja_Post ^
main.py

pause
