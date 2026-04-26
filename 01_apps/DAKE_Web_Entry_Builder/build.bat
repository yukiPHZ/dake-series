@echo off
chcp 65001 > nul

rmdir /s /q build 2>nul
rmdir /s /q dist 2>nul
del *.spec 2>nul

pyinstaller ^
--onefile ^
--noconsole ^
--clean ^
--noconfirm ^
--name DakeWebOne_Builder ^
--icon=..\..\02_assets\dake_icon.ico ^
main.py

pause
