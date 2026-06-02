@echo off

rmdir /s /q build
rmdir /s /q dist
del *.spec

pyinstaller ^
--name DakeVideo_Shorts_Cut ^
--onefile ^
--noconsole ^
--clean ^
--icon=..\..\02_assets\dake_icon.ico ^
main.py

pause
