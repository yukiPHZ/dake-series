@echo off
setlocal EnableExtensions

cd /d "%~dp0"

set "APP_EXE_NAME=Dake_AllStaff_Mail"
set "ENTRY_FILE=main.py"
set "ICON_FILE=..\..\02_assets\dake_icon.ico"
set "VERSION_FILE=version_info.txt"
set "PYTHON_EXE="

if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
del /q *.spec 2>nul

if not exist "%ENTRY_FILE%" (
    echo [ERROR] Missing %ENTRY_FILE%
    pause
    exit /b 1
)

if not exist "%ICON_FILE%" (
    echo [ERROR] Missing %ICON_FILE%
    pause
    exit /b 1
)

where pyinstaller >nul 2>nul
if %errorlevel%==0 (
    pyinstaller ^
    --onefile ^
    --noconsole ^
    --clean ^
    --icon=..\..\02_assets\dake_icon.ico ^
    --version-file=version_info.txt ^
    --name Dake_AllStaff_Mail ^
    main.py
    goto AFTER_BUILD
)

where python >nul 2>nul
if %errorlevel%==0 set "PYTHON_EXE=python"

if not defined PYTHON_EXE (
    where py >nul 2>nul
    if %errorlevel%==0 set "PYTHON_EXE=py -3"
)

if not defined PYTHON_EXE (
    if exist "%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" (
        set "PYTHON_EXE=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
    )
)

if not defined PYTHON_EXE (
    echo [ERROR] Python or PyInstaller was not found.
    pause
    exit /b 1
)

%PYTHON_EXE% -m PyInstaller ^
--onefile ^
--noconsole ^
--clean ^
--icon=..\..\02_assets\dake_icon.ico ^
--version-file=version_info.txt ^
--name Dake_AllStaff_Mail ^
main.py

:AFTER_BUILD
if errorlevel 1 (
    echo [ERROR] PyInstaller build failed.
    pause
    exit /b 1
)

if not exist "dist\%APP_EXE_NAME%.exe" (
    echo [ERROR] Build finished, but exe was not found.
    pause
    exit /b 1
)

echo [SUCCESS] dist\%APP_EXE_NAME%.exe

pause
