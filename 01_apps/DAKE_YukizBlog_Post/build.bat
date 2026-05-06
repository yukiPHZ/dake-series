@echo off
setlocal
cd /d "%~dp0"

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist Dake_YukizBlog_Post.spec del Dake_YukizBlog_Post.spec

set "PYTHON_EXE=python"
if exist "%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" (
  set "PYTHON_EXE=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
)

set "ICON_OPT="
if exist "..\..\02_assets\dake_icon.ico" (
  set "ICON_OPT=--icon=..\..\02_assets\dake_icon.ico"
)

"%PYTHON_EXE%" -m PyInstaller ^
  --onefile ^
  --noconsole ^
  --clean ^
  --collect-all customtkinter ^
  %ICON_OPT% ^
  --version-file version_info.txt ^
  --name Dake_YukizBlog_Post ^
  main.py

endlocal
