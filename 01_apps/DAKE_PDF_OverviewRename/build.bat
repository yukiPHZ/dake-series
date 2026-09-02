@echo off
chcp 65001 > nul
cd /d "%~dp0"

rmdir /s /q build 2>nul
rmdir /s /q dist 2>nul
del /q *.spec 2>nul

where pyinstaller >nul 2>nul
if %errorlevel%==0 (
    set "PYINSTALLER=pyinstaller"
    goto build
)

where python >nul 2>nul
if %errorlevel%==0 (
    set "PYINSTALLER=python -m PyInstaller"
    goto build
)

where py >nul 2>nul
if %errorlevel%==0 (
    set "PYINSTALLER=py -m PyInstaller"
    goto build
)

set "CODEX_PY=%USERPROFILE%\AppData\Local\Programs\Python\Python312\python.exe"
if exist "%CODEX_PY%" (
    set PYINSTALLER="%CODEX_PY%" -m PyInstaller
    goto build
)

echo PyInstaller was not found.
echo Run pip install -r requirements.txt, then run build.bat again.
exit /b 1

:build
%PYINSTALLER% ^
--onefile ^
--noconsole ^
--clean ^
--name=DakePDF_OverviewRename ^
--icon=..\..\02_assets\dake_icon.ico ^
--add-data=..\..\02_assets\dake_icon.ico;. ^
main.py

exit /b %errorlevel%
