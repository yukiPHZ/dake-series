@echo off
setlocal EnableExtensions

cd /d "%~dp0"
del /q version_info.txt 2>nul

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist *.spec del *.spec

set "PYINSTALLER_CMD="
set "PYINSTALLER_ARGS="

where pyinstaller >nul 2>nul
if %errorlevel%==0 (
    set "PYINSTALLER_CMD=pyinstaller"
)

if not defined PYINSTALLER_CMD (
    where py >nul 2>nul
    if %errorlevel%==0 (
        py -3 -m PyInstaller --version >nul 2>nul
        if %errorlevel%==0 (
            set "PYINSTALLER_CMD=py"
            set "PYINSTALLER_ARGS=-3 -m PyInstaller"
        )
    )
)

if not defined PYINSTALLER_CMD (
    where python >nul 2>nul
    if %errorlevel%==0 (
        python -m PyInstaller --version >nul 2>nul
        if %errorlevel%==0 (
            set "PYINSTALLER_CMD=python"
            set "PYINSTALLER_ARGS=-m PyInstaller"
        )
    )
)

if not defined PYINSTALLER_CMD (
    if exist "%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" (
        "%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" -m PyInstaller --version >nul 2>nul
        if %errorlevel%==0 (
            set "PYINSTALLER_CMD=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
            set "PYINSTALLER_ARGS=-m PyInstaller"
        )
    )
)

if not defined PYINSTALLER_CMD (
    echo PyInstaller was not found.
    echo Install PyInstaller, then run build.bat again.
    pause
    exit /b 1
)
python ..\..\tools\generate_version_info.py --app-dir . --out version_info.txt
if %errorlevel% neq 0 (
    echo VersionInfo generation failed.
    pause
    exit /b 1
)

"%PYINSTALLER_CMD%" %PYINSTALLER_ARGS% ^
--onefile ^
--noconsole ^
--clean ^
--name DakeWork_Calendar ^
--paths=..\..\00_core ^
--icon=..\..\02_assets\dake_icon.ico ^
--version-file version_info.txt ^
main.py

pause
