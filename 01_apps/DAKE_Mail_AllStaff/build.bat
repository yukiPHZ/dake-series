@echo off

rmdir /s /q build
rmdir /s /q dist
del *.spec

pyinstaller ^
--onefile ^
--noconsole ^
--clean ^
--icon=..\..\02_assets\dake_icon.ico ^
--name Dake_AllStaff_Mail ^
main.py

pause
