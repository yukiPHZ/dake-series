@echo off
chcp 65001 > nul
cd /d "%~dp0"

set "APP_EXE_NAME=DakeBackup"
set "PYTHON_EXE=python"

where python >nul 2>nul
if errorlevel 1 (
    if exist "%LocalAppData%\Programs\Python\Python312\python.exe" (
        set "PYTHON_EXE=%LocalAppData%\Programs\Python\Python312\python.exe"
    )
)

rmdir /s /q build 2>nul
rmdir /s /q dist 2>nul
del /q *.spec 2>nul

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

if defined PYINSTALLER_EXE goto RUN_PYINSTALLER
goto RUN_PYTHON_MODULE

:RUN_PYINSTALLER
"%PYINSTALLER_EXE%" ^
--onefile ^
--noconsole ^
--clean ^
--icon=..\..\02_assets\dake_icon.ico ^
--name %APP_EXE_NAME% ^
main.py
goto CHECK_BUILD

:RUN_PYTHON_MODULE
"%PYTHON_EXE%" -m PyInstaller ^
--onefile ^
--noconsole ^
--clean ^
--icon=..\..\02_assets\dake_icon.ico ^
--name %APP_EXE_NAME% ^
main.py

:CHECK_BUILD
if errorlevel 1 (
    echo.
    echo build failed
    exit /b 1
)

echo.
echo dist\%APP_EXE_NAME%.exe created.

