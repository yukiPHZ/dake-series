@echo off
chcp 65001 > nul
cd /d "%~dp0"

set "APP_EXE=DakeImage_PasteA4.exe"
set "PYINSTALLER_CMD="
set "PYTHON_CMD="

where pyinstaller > nul 2>&1
if %errorlevel%==0 set "PYINSTALLER_CMD=pyinstaller"

if not defined PYINSTALLER_CMD (
  if defined PYTHON_EXE (
    if exist "%PYTHON_EXE%" set "PYTHON_CMD=%PYTHON_EXE%"
  )
)

if not defined PYINSTALLER_CMD (
  if not defined PYTHON_CMD (
    where py > nul 2>&1
    if %errorlevel%==0 set "PYTHON_CMD=py"
  )
)

if not defined PYINSTALLER_CMD (
  if not defined PYTHON_CMD (
    where python > nul 2>&1
    if %errorlevel%==0 set "PYTHON_CMD=python"
  )
)

if not defined PYINSTALLER_CMD (
  if defined PYTHON_CMD set "PYINSTALLER_CMD=%PYTHON_CMD% -m PyInstaller"
)

if not defined PYTHON_CMD (
  where python > nul 2>&1
  if %errorlevel%==0 set "PYTHON_CMD=python"
)

rmdir /s /q build 2>nul
rmdir /s /q dist 2>nul
del /q *.spec 2>nul
del /q version_info.txt 2>nul

if not defined PYINSTALLER_CMD (
  echo [ERROR] PyInstaller was not found.
  goto :error
)

%PYTHON_CMD% ..\..\tools\generate_version_info.py --app-dir . --out version_info.txt
if errorlevel 1 (
    echo VersionInfo generation failed.
    pause
    exit /b 1
)

%PYINSTALLER_CMD% ^
--onefile ^
--noconsole ^
--clean ^
--paths=..\..\00_core ^
--icon=..\..\02_assets\dake_icon.ico ^
--version-file version_info.txt ^
--name DakeImage_PasteA4 ^
main.py

if errorlevel 1 goto :error

if not exist "dist\%APP_EXE%" (
  echo [ERROR] dist\%APP_EXE% was not found.
  goto :error
)

echo Build complete: dist\%APP_EXE%
pause
exit /b 0

:error
echo Build failed.
pause
exit /b 1
