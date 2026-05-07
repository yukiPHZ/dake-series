@echo off
setlocal

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
del /q *.spec 2>nul

set "ICON_PATH=..\..\02_assets\dake_icon.ico"
set "BUNDLED_PY=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"

if exist "%BUNDLED_PY%" (
    "%BUNDLED_PY%" -m PyInstaller ^
    --onefile ^
    --noconsole ^
    --clean ^
    --icon="%ICON_PATH%" ^
    --exclude-module pandas ^
    --exclude-module numpy ^
    --exclude-module openpyxl ^
    --exclude-module lxml ^
    --exclude-module PIL ^
    --name DakePDF_Marker ^
    main.py
) else (
    pyinstaller ^
    --onefile ^
    --noconsole ^
    --clean ^
    --icon="%ICON_PATH%" ^
    --exclude-module pandas ^
    --exclude-module numpy ^
    --exclude-module openpyxl ^
    --exclude-module lxml ^
    --exclude-module PIL ^
    --name DakePDF_Marker ^
    main.py
)

exit /b %ERRORLEVEL%
