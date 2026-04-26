@echo off
chcp 65001 > nul
setlocal
cd /d "%~dp0"

set "APP_EXE=DakeImageToPDF.exe"
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
  echo [ERROR] Python was not found.
  echo Example: set PYTHON_EXE=C:\Path\To\python.exe
  goto :error
)

echo.
echo ========================================
echo Cleanup
echo ========================================
if exist build (
  rmdir /s /q build
  if errorlevel 1 (
    echo [ERROR] Failed to remove the build folder.
    goto :error
  )
)
if exist dist (
  rmdir /s /q dist
  if errorlevel 1 (
    echo [ERROR] Failed to remove the dist folder.
    goto :error
  )
)

echo.
echo ========================================
echo Upgrade pip
echo ========================================
%PYTHON_CMD% -m pip install --upgrade pip
if errorlevel 1 (
  echo [ERROR] Failed to upgrade pip.
  goto :error
)

echo.
echo ========================================
echo Install requirements
echo ========================================
%PYTHON_CMD% -m pip install -r requirements.txt
if errorlevel 1 (
  echo [ERROR] Failed to install requirements.txt.
  goto :error
)

echo.
echo ========================================
echo Run PyInstaller
echo ========================================
%PYTHON_CMD% -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --onefile ^
  --windowed ^
  --name=DakeImageToPDF ^
  --icon=..\..\02_assets\dake_icon.ico ^
  --hidden-import tkinterdnd2 ^
  --collect-all tkinterdnd2 ^
  main.py
if errorlevel 1 (
  echo [ERROR] PyInstaller failed.
  goto :error
)

echo.
echo ========================================
echo Verify output
echo ========================================
if not exist "dist\%APP_EXE%" (
  echo [ERROR] dist\%APP_EXE% was not found.
  goto :error
)

echo.
echo Build complete: dist\%APP_EXE%
pause
exit /b 0

:error
echo.
echo Build failed.
pause
exit /b 1
