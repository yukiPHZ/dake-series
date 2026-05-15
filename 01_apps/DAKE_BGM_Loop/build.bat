@echo off
chcp 65001 > nul
cd /d "%~dp0"

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
del /q *.spec 2>nul

python -m pip install -r requirements.txt
if errorlevel 1 exit /b 1

python -m PyInstaller ^
  --onefile ^
  --noconsole ^
  --clean ^
  --name DakeBGM_Loop ^
  --icon=..\..\02_assets\dake_icon.ico ^
  main.py
if errorlevel 1 exit /b 1

echo.
echo Build complete: dist\DakeBGM_Loop.exe
