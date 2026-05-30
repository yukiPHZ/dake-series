@echo off
chcp 65001 > nul
cd /d "%~dp0"

rmdir /s /q build 2>nul
rmdir /s /q dist 2>nul
del /q *.spec 2>nul
del /q version_info.txt 2>nul

set "PYINSTALLER=pyinstaller"
where pyinstaller >nul 2>nul
if errorlevel 1 (
    if exist "%LocalAppData%\Programs\Python\Python312\Scripts\pyinstaller.exe" (
        set "PYINSTALLER=%LocalAppData%\Programs\Python\Python312\Scripts\pyinstaller.exe"
    )
)

python ..\..\tools\generate_version_info.py --app-dir . --out version_info.txt
if errorlevel 1 (
    echo VersionInfo generation failed.
    pause
    exit /b 1
)

"%PYINSTALLER%" ^
--onefile ^
--noconsole ^
--clean ^
--paths=..\..\00_core ^
--icon=..\..\02_assets\dake_icon.ico ^
--version-file version_info.txt ^
--name DakeYesterday_Task_Memo ^
main.py

if errorlevel 1 (
    echo Build failed.
    pause
    exit /b 1
)

pause
