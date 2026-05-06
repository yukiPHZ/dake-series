@echo off
cd /d "%~dp0"

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist Dake_HolidayJinja_Post.spec del Dake_HolidayJinja_Post.spec

set "PYTHON_EXE=python"
if exist "%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" (
  set "PYTHON_EXE=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
)

"%PYTHON_EXE%" -m PyInstaller ^
--onefile ^
--noconsole ^
--clean ^
--collect-all customtkinter ^
--icon=..\..\02_assets\dake_icon.ico ^
--version-file version_info.txt ^
--name Dake_HolidayJinja_Post ^
main.py

pause
