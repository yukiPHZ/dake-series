@echo off
chcp 65001 > nul
setlocal
cd /d "%~dp0"

set "APP_NAME=DAKE_Wake_Brainz"
set "DIST_DIR=dist"
set "BUILD_DIR=build"
set "SPEC_FILE=%APP_NAME%.spec"
set "OUTPUT_EXE=%DIST_DIR%\DAKE_Wake_Brainz.exe"
set "ICON_FILE=..\..\02_assets\dake_icon.ico"

set "PYTHON_EXE=python"
where python >nul 2>nul
if errorlevel 1 (
    if exist "%LocalAppData%\Programs\Python\Python312\python.exe" (
        set "PYTHON_EXE=%LocalAppData%\Programs\Python\Python312\python.exe"
    )
)

set "PYINSTALLER_EXE="
where pyinstaller >nul 2>nul
if not errorlevel 1 set "PYINSTALLER_EXE=pyinstaller"
if not defined PYINSTALLER_EXE (
    if exist "%LocalAppData%\Programs\Python\Python312\Scripts\pyinstaller.exe" (
        set "PYINSTALLER_EXE=%LocalAppData%\Programs\Python\Python312\Scripts\pyinstaller.exe"
    )
)

echo [1/3] Cleaning previous build artifacts...
if exist "%BUILD_DIR%" rmdir /s /q "%BUILD_DIR%"
if exist "%DIST_DIR%" rmdir /s /q "%DIST_DIR%"
if exist "%SPEC_FILE%" del /q "%SPEC_FILE%"
if exist "main.spec" del /q "main.spec"

if exist "%BUILD_DIR%" goto :clean_error
if exist "%DIST_DIR%" goto :clean_error

echo [2/3] Building %APP_NAME%.exe...
if defined PYINSTALLER_EXE (
    "%PYINSTALLER_EXE%" ^
     --clean ^
     --noconfirm ^
     --onefile ^
     --name %APP_NAME% ^
     --icon=..\..\02_assets\dake_icon.ico ^
     --add-data "templates;templates" ^
     --add-data "static;static" ^
     --add-data "config.example.json;." ^
     main.py
) else (
    "%PYTHON_EXE%" -m PyInstaller ^
     --clean ^
     --noconfirm ^
     --onefile ^
     --name %APP_NAME% ^
     --icon=..\..\02_assets\dake_icon.ico ^
     --add-data "templates;templates" ^
     --add-data "static;static" ^
     --add-data "config.example.json;." ^
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
