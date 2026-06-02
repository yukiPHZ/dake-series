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
--name DakeGame_Diver_Catch ^
main.py

if errorlevel 1 (
  echo Build failed.
  exit /b 1
)

if not exist dist\DakeGame_Diver_Catch.exe (
  echo dist\DakeGame_Diver_Catch.exe was not created.
  exit /b 1
)

echo Build OK: dist\DakeGame_Diver_Catch.exe
