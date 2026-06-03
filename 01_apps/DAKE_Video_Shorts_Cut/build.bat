@echo off

cd /d %~dp0

rmdir /s /q build 2>nul
rmdir /s /q dist 2>nul
del /q *.spec 2>nul
del /q version_info.txt 2>nul

python ..\..\tools\generate_version_info.py --app-dir . --out version_info.txt

pyinstaller ^
--name DakeVideo_Shorts_Cut ^
--onefile ^
--noconsole ^
--clean ^
--paths=..\..\00_core ^
--icon=..\..\02_assets\dake_icon.ico ^
--version-file version_info.txt ^
main.py

pause
