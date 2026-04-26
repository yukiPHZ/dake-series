@echo off
setlocal
cd /d "%~dp0"

set "APP_NAME=DakePDF_Merge"
set "DIST_DIR=dist"
set "BUILD_DIR=build"
set "SPEC_FILE=%APP_NAME%.spec"
set "OUTPUT_EXE=%DIST_DIR%\%APP_NAME%.exe"

rem If the previous exe is still running, cleanup or overwrite may fail.
echo [1/3] Cleaning previous build artifacts...
if exist "%BUILD_DIR%" rmdir /s /q "%BUILD_DIR%"
if exist "%DIST_DIR%" rmdir /s /q "%DIST_DIR%"
if exist "%SPEC_FILE%" del /q "%SPEC_FILE%"
if exist "main.spec" del /q "main.spec"
if exist "DAKE_PDF_Merge.spec" del /q "DAKE_PDF_Merge.spec"

if exist "%BUILD_DIR%" goto :clean_error
if exist "%DIST_DIR%" goto :clean_error


echo [2/3] Building %APP_NAME%.exe...
pyinstaller --clean --noconfirm ^
 main.py ^
 --onefile ^
 --noconsole ^
 --name=DakePDF_Merge ^
 --icon=..\..\02_assets\dake_icon.ico ^
 --add-data "..\..\02_assets\dake_icon.ico;."
if errorlevel 1 goto :build_error


echo [3/3] Verifying output...
if exist "%OUTPUT_EXE%" goto :success

echo.
echo Build finished but %OUTPUT_EXE% was not found.
goto :build_error

:clean_error
echo.
echo Cleanup failed. A previous build folder or exe may still be in use.
pause
exit /b 1

:build_error
echo.
echo Build failed.
echo Expected output: %OUTPUT_EXE%
pause
exit /b 1

:success
echo.
echo Build success.
echo Output: %OUTPUT_EXE%
pause
exit /b 0
