@echo off

rmdir /s /q build
rmdir /s /q dist
del *.spec
del /q version_info.txt 2>nul

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

set "CODEX_PY=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if exist "%CODEX_PY%" (
    set "PYINSTALLER=\"%CODEX_PY%\" -m PyInstaller"
    goto build
)

echo PyInstaller was not found.
echo Run pip install -r requirements.txt, then run build.bat again.
pause
exit /b 1

:build
python ..\..\tools\generate_version_info.py --app-dir . --out version_info.txt
if errorlevel 1 (
    echo VersionInfo generation failed.
    pause
    exit /b 1
)

%PYINSTALLER% ^
--onefile ^
--noconsole ^
--clean ^
--name=DakePDF_Viewer ^
--paths=..\..\00_core ^
--icon=..\..\02_assets\dake_icon.ico ^
--version-file version_info.txt ^
main.py

pause
