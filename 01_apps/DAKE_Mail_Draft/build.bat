@echo off
chcp 65001 > nul
cd /d "%~dp0"

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
del *.spec 2>nul

if exist "..\..\02_assets\dake_icon.ico" (
  pyinstaller --onefile --noconsole --clean --name DAKE_Mail_Draft --icon=..\..\02_assets\dake_icon.ico main.py
) else (
  echo Common icon was not found. Building without icon.
  pyinstaller --onefile --noconsole --clean --name DAKE_Mail_Draft main.py
)

if errorlevel 1 (
  echo.
  echo Build failed.
  exit /b %ERRORLEVEL%
)

echo.
echo dist\DAKE_Mail_Draft.exe was created.
