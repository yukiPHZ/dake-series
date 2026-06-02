@echo off
setlocal

cd /d "%~dp0"

set APP_NAME=DakeNote_Inbox
set ENTRY=main.py
set ICON=..\..\02_assets\dake_icon.ico

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist "%APP_NAME%.spec" del "%APP_NAME%.spec"
if exist version_info.txt del version_info.txt

pyinstaller --noconfirm --onefile --windowed --name "%APP_NAME%" --icon "%ICON%" --add-data "%ICON%;assets" "%ENTRY%"

if errorlevel 1 (
  echo build failed
  exit /b 1
)

echo dist\%APP_NAME%.exe
endlocal
