@echo off
chcp 65001 > nul
cd /d "%~dp0"

rmdir /s /q build 2>nul
rmdir /s /q dist 2>nul
del /q *.spec 2>nul
del /q version_info.txt 2>nul

python ..\..\tools\generate_version_info.py --app-dir . --out version_info.txt
if errorlevel 1 (
    echo VersionInfo generation failed.
    pause
    exit /b 1
)

pyinstaller ^
--noconfirm ^
--onefile ^
--noconsole ^
--clean ^
--name DakePDF_Insert ^
--paths=..\..\00_core ^
--icon=..\..\02_assets\dake_icon.ico ^
--add-data "..\..\02_assets\dake_icon.ico;." ^
--version-file version_info.txt ^
--hidden-import=tkinterdnd2 ^
--hidden-import=fitz ^
--exclude-module=numpy ^
--exclude-module=matplotlib ^
--exclude-module=pandas ^
--exclude-module=pygame ^
--exclude-module=pytest ^
--exclude-module=IPython ^
main.py

if errorlevel 1 exit /b 1
if not exist "dist\DakePDF_Insert.exe" exit /b 1
pause
