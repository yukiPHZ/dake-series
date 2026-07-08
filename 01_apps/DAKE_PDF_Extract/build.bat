@echo off
chcp 65001 > nul
cd /d "%~dp0"

rmdir /s /q build 2>nul
rmdir /s /q dist 2>nul
del /q *.spec 2>nul
del /q version_info.txt 2>nul

python -m pip install -r requirements.txt
if errorlevel 1 (
  echo Dependency install failed.
  pause
  exit /b 1
)

python ..\..\tools\generate_version_info.py --app-dir . --out version_info.txt
if errorlevel 1 (
  echo VersionInfo generation failed.
  pause
  exit /b 1
)

pyinstaller ^
--onefile ^
--noconsole ^
--clean ^
--paths=..\..\00_core ^
--name DakePDF_Extract ^
--icon=..\..\02_assets\dake_icon.ico ^
--add-data "..\..\02_assets\dake_icon.ico;." ^
--version-file version_info.txt ^
--hidden-import=tkinterdnd2 ^
--collect-all=tkinterdnd2 ^
--hidden-import=fitz ^
--collect-all=fitz ^
main.py

pause
