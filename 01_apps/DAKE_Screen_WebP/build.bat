@echo off
chcp 65001 > nul
cd /d "%~dp0"

rmdir /s /q build 2>nul
rmdir /s /q dist 2>nul
del *.spec 2>nul

pyinstaller ^
--onefile ^
--noconsole ^
--clean ^
--name DakeScreen_WebP ^
--icon=..\..\02_assets\dake_icon.ico ^
--version-file=version_info.txt ^
main.py

pause
