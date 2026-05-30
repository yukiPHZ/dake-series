@echo off

rmdir /s /q build
rmdir /s /q dist
del *.spec
del /q version_info.txt 2>nul

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
--name DakeFixedTax ^
--paths=..\..\00_core ^
--icon=..\..\02_assets\dake_icon.ico ^
--version-file version_info.txt ^
main.py

pause
