@echo off
chcp 65001 > nul
cd /d "%~dp0"

rmdir /s /q build 2>nul
rmdir /s /q dist 2>nul
del /q *.spec 2>nul

pyinstaller ^
--onefile ^
--noconsole ^
--clean ^
--icon=..\..\02_assets\dake_icon.ico ^
--collect-all=tkinterdnd2 ^
--hidden-import=extract_msg ^
--collect-all=extract_msg ^
--name DakeMail_List ^
main.py

pause
