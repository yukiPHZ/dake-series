@echo off
chcp 65001 > nul
setlocal
cd /d "%~dp0"

set "APP_EXE=DakeImage_BatchPDF.exe"
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
  if errorlevel 1 goto :error
)
if exist dist (
  rmdir /s /q dist
  if errorlevel 1 goto :error
)
for %%F in (*.spec) do del /q "%%~fF"
del /q version_info.txt 2>nul

echo.
echo ========================================
echo Install requirements
echo ========================================
%PYTHON_CMD% -m pip install -r requirements.txt
if errorlevel 1 goto :error

echo.
echo ========================================
echo Generate version info
echo ========================================
%PYTHON_CMD% ..\..\tools\generate_version_info.py --app-dir . --out version_info.txt
if errorlevel 1 goto :error

echo.
echo ========================================
echo Run PyInstaller
echo ========================================
%PYTHON_CMD% -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --onefile ^
  --windowed ^
  --noconsole ^
  --name=DakeImage_BatchPDF ^
  --paths=..\..\00_core ^
  --icon=..\..\02_assets\dake_icon.ico ^
  --version-file version_info.txt ^
  --hidden-import tkinterdnd2 ^
  --collect-all tkinterdnd2 ^
  --collect-all=pillow_heif ^
  main.py
if errorlevel 1 goto :error

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
