@echo off
setlocal

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
del /q *.spec 2>nul
del /q version_info.txt 2>nul

set "BUNDLED_PY=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
set "USE_BUNDLED_PY="

if exist "%BUNDLED_PY%" (
    "%BUNDLED_PY%" -m PyInstaller --version >nul 2>nul
    if not errorlevel 1 set "USE_BUNDLED_PY=1"
)

if defined USE_BUNDLED_PY (
python ..\..\tools\generate_version_info.py --app-dir . --out version_info.txt
if errorlevel 1 (
    echo VersionInfo generation failed.
    pause
    exit /b 1
)

    "%BUNDLED_PY%" -m PyInstaller ^
    --onefile ^
    --noconsole ^
    --clean ^
--paths=..\..\00_core ^
    --icon=..\..\02_assets\dake_icon.ico ^
--version-file version_info.txt ^
    --exclude-module pandas ^
    --exclude-module numpy ^
    --exclude-module openpyxl ^
    --exclude-module lxml ^
    --exclude-module PIL ^
    --name DakePDF_Marker ^
    main.py
) else (
python ..\..\tools\generate_version_info.py --app-dir . --out version_info.txt
if errorlevel 1 (
    echo VersionInfo generation failed.
    pause
    exit /b 1
)

    pyinstaller ^
    --onefile ^
    --noconsole ^
    --clean ^
--paths=..\..\00_core ^
    --icon=..\..\02_assets\dake_icon.ico ^
--version-file version_info.txt ^
    --exclude-module pandas ^
    --exclude-module numpy ^
    --exclude-module openpyxl ^
    --exclude-module lxml ^
    --exclude-module PIL ^
    --name DakePDF_Marker ^
    main.py
)

exit /b %ERRORLEVEL%
