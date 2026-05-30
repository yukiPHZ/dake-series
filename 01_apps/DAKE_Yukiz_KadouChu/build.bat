@echo off
chcp 65001 > nul
cd /d "%~dp0"

set "APP_NAME=DakeYukiz_KadouChu"
set "PYTHON_EXE="
set "PYTHON_ARGS="

where python >nul 2>nul
if not errorlevel 1 set "PYTHON_EXE=python"

if not defined PYTHON_EXE (
  where py >nul 2>nul
  if not errorlevel 1 (
    set "PYTHON_EXE=py"
    set "PYTHON_ARGS=-3"
  )
)

if not defined PYTHON_EXE (
  if exist "%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" (
    set "PYTHON_EXE=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
  )
)

if not defined PYTHON_EXE (
  if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" (
    set "PYTHON_EXE=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
  )
)

if not defined PYTHON_EXE (
  echo [ERROR] Python was not found.
  pause
  exit /b 1
)

"%PYTHON_EXE%" %PYTHON_ARGS% -m PyInstaller --version >nul 2>nul
if errorlevel 1 (
  echo [ERROR] PyInstaller was not found.
  echo Install requirements, then run build.bat again.
  pause
  exit /b 1
)

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
del /q *.spec 2>nul
del /q version_info.txt 2>nul

python ..\..\tools\generate_version_info.py --app-dir . --out version_info.txt
if errorlevel 1 (
    echo VersionInfo generation failed.
    pause
    exit /b 1
)

"%PYTHON_EXE%" %PYTHON_ARGS% -m PyInstaller ^
--onefile ^
--noconsole ^
--clean ^
--name %APP_NAME% ^
--collect-all customtkinter ^
--hidden-import PIL._tkinter_finder ^
--paths=..\..\00_core ^
--icon=..\..\02_assets\dake_icon.ico ^
--version-file version_info.txt ^
main.py

if errorlevel 1 (
  echo.
  echo [ERROR] PyInstaller build failed.
  pause
  exit /b 1
)

echo.
echo Build complete: dist\DakeYukiz_KadouChu.exe
pause
