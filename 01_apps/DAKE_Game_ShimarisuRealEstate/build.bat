@echo off

rmdir /s /q build
rmdir /s /q dist
del *.spec

pyinstaller ^
--onefile ^
--noconsole ^
--clean ^
--name DakeShimarisuRealEstate ^
--add-data "assets;assets" ^
main.py

pause
