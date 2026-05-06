@echo off

rmdir /s /q build
rmdir /s /q dist
del *.spec

pyinstaller ^
--onefile ^
--noconsole ^
--clean ^
--name DakeYear_Notice ^
--icon=..\..\02_assets\dake_icon.ico ^
main.py

del *.spec

pause
