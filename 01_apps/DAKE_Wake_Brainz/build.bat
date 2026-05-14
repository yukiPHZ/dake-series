@echo off
chcp 65001 > nul
setlocal
cd /d "%~dp0"

set "APP_NAME=DakeWake_Brainz"
set "DIST_DIR=dist"
set "BUILD_DIR=build"
set "SPEC_FILE=%APP_NAME%.spec"
set "OUTPUT_EXE=%DIST_DIR%\DakeWake_Brainz.exe"
set "ICON_FILE=assets\peakheadz_logo.ico"

echo [1/3] Cleaning previous build artifacts...
if exist "%BUILD_DIR%" rmdir /s /q "%BUILD_DIR%"
if exist "%DIST_DIR%" rmdir /s /q "%DIST_DIR%"
if exist "%SPEC_FILE%" del /q "%SPEC_FILE%"
if exist "main.spec" del /q "main.spec"

if exist "%BUILD_DIR%" goto :clean_error
if exist "%DIST_DIR%" goto :clean_error

if not exist "%ICON_FILE%" (
    echo.
    echo Missing icon: %ICON_FILE%
    goto :build_error
)

set "PYTHON_EXE=python"
where python >nul 2>nul
if errorlevel 1 (
    if exist "%LocalAppData%\Programs\Python\Python312\python.exe" (
        set "PYTHON_EXE=%LocalAppData%\Programs\Python\Python312\python.exe"
    )
)

set "PYINSTALLER_EXE="
where pyinstaller >nul 2>nul
if not errorlevel 1 (
    set "PYINSTALLER_EXE=pyinstaller"
)
if not defined PYINSTALLER_EXE (
    if exist "%LocalAppData%\Programs\Python\Python312\Scripts\pyinstaller.exe" (
        set "PYINSTALLER_EXE=%LocalAppData%\Programs\Python\Python312\Scripts\pyinstaller.exe"
    )
)

echo [2/3] Building DakeWake_Brainz.exe...
if defined PYINSTALLER_EXE (
    "%PYINSTALLER_EXE%" ^
     --clean ^
     --noconfirm ^
     --onefile ^
     --noconsole ^
     --name DakeWake_Brainz ^
     --icon=%ICON_FILE% ^
     --exclude-module numpy ^
     --add-data "assets;assets" ^
     main.py
) else (
    "%PYTHON_EXE%" -m PyInstaller ^
     --clean ^
     --noconfirm ^
     --onefile ^
     --noconsole ^
     --name DakeWake_Brainz ^
     --icon=%ICON_FILE% ^
     --exclude-module numpy ^
     --add-data "assets;assets" ^
     main.py
)
if errorlevel 1 goto :build_error

echo [3/3] Verifying output...
if exist "%OUTPUT_EXE%" goto :success

echo.
echo Build finished but %OUTPUT_EXE% was not found.
goto :build_error

:clean_error
echo.
echo Cleanup failed. A previous build folder or exe may still be in use.
exit /b 1

:build_error
echo.
echo Build failed.
echo Expected output: %OUTPUT_EXE%
exit /b 1

:success
echo.
echo Build success.
echo Output: %OUTPUT_EXE%
exit /b 0
