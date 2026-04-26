@echo off
setlocal EnableExtensions

cd /d "%~dp0"

set "APP_EXE_NAME=DakePDF_to_Images"
set "ENTRY_FILE=main.py"
set "ICON_FILE=..\..\02_assets\dake_icon.ico"
set "VENV_DIR=.venv"
set "VENV_PYTHON=%VENV_DIR%\Scripts\python.exe"
set "BOOTSTRAP_EXE="
set "BOOTSTRAP_ARGS="

echo [INFO] Working directory: %CD%

if not exist "%ENTRY_FILE%" (
    echo [ERROR] Missing %ENTRY_FILE%
    exit /b 1
)

if not exist "%ICON_FILE%" (
    echo [ERROR] Missing %ICON_FILE%
    exit /b 1
)

where py >nul 2>nul
if %errorlevel%==0 (
    set "BOOTSTRAP_EXE=py"
    set "BOOTSTRAP_ARGS=-3"
)

if not defined BOOTSTRAP_EXE (
    where python >nul 2>nul
    if %errorlevel%==0 (
        set "BOOTSTRAP_EXE=python"
        set "BOOTSTRAP_ARGS="
    )
)

if not defined BOOTSTRAP_EXE (
    if exist "%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" (
        set "BOOTSTRAP_EXE=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
        set "BOOTSTRAP_ARGS="
    )
)

if not defined BOOTSTRAP_EXE (
    echo [ERROR] Python was not found. Install Python 3.10+ or py launcher first.
    exit /b 1
)

echo [INFO] Cleaning previous build output...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
for %%F in (*.spec) do del /q "%%~fF"

if not exist "%VENV_PYTHON%" (
    echo [INFO] Creating virtual environment...
    if defined BOOTSTRAP_ARGS (
        "%BOOTSTRAP_EXE%" %BOOTSTRAP_ARGS% -m venv "%VENV_DIR%"
    ) else (
        "%BOOTSTRAP_EXE%" -m venv "%VENV_DIR%"
    )
    if errorlevel 1 (
        echo [ERROR] Failed to create virtual environment.
        exit /b 1
    )
)

if not exist "%VENV_PYTHON%" (
    echo [ERROR] Virtual environment python not found: %VENV_PYTHON%
    exit /b 1
)

echo [INFO] Upgrading pip...
"%VENV_PYTHON%" -m pip install --upgrade pip
if errorlevel 1 (
    echo [ERROR] pip upgrade failed.
    exit /b 1
)

echo [INFO] Installing requirements...
"%VENV_PYTHON%" -m pip install -r "requirements.txt"
if errorlevel 1 (
    echo [ERROR] Dependency installation failed.
    exit /b 1
)

echo [INFO] Building exe with PyInstaller...
"%VENV_PYTHON%" -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --windowed ^
  --onefile ^
  --name=DakePDF_to_Images ^
  --icon=..\..\02_assets\dake_icon.ico ^
  --hidden-import fitz ^
  --collect-all fitz ^
  "%ENTRY_FILE%"

if errorlevel 1 (
    echo [ERROR] PyInstaller build failed.
    exit /b 1
)

if not exist "dist\%APP_EXE_NAME%.exe" (
    echo [ERROR] Build finished, but exe was not found.
    exit /b 1
)

echo [SUCCESS] dist\%APP_EXE_NAME%.exe
start "" explorer.exe "%CD%\dist"
exit /b 0
