@echo off
chcp 65001 > nul
setlocal

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

set "PYTHON_CMD="

if defined PYTHON_EXE (
  if exist "%PYTHON_EXE%" set "PYTHON_CMD=%PYTHON_EXE%"
)

if not defined PYTHON_CMD (
  where py > nul 2>&1
  if %errorlevel%==0 set "PYTHON_CMD=py"
)

if not defined PYTHON_CMD (
  where python > nul 2>&1
  if %errorlevel%==0 set "PYTHON_CMD=python"
)

if not defined PYTHON_CMD (
  echo Python executable was not found.
  echo Example: set PYTHON_EXE=C:\Path\To\python.exe
  pause
  exit /b 1
)

if exist ".vendor" (
  set "PYTHONPATH=%cd%\.vendor;%PYTHONPATH%"
)

%PYTHON_CMD% -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --onefile ^
  --windowed ^
  --name=DakePDF_Split_One ^
  --icon=..\..\02_assets\dake_icon.ico ^
  --collect-all tkinterdnd2 ^
  main.py

echo.
echo Build complete: dist\DakePDF_Split_One.exe
pause
