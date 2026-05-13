@echo off
chcp 65001 > nul
setlocal
cd /d "%~dp0"

set "APP_NAME=DakeColumn_Memo"
set "DIST_DIR=dist"
set "BUILD_DIR=build"
set "SPEC_FILE=%APP_NAME%.spec"
set "OUTPUT_EXE=%DIST_DIR%\%APP_NAME%.exe"
set "DAKE_ICON=..\..\02_assets\dake_icon.ico"
set "PYINSTALLER_CMD=pyinstaller"

where pyinstaller > nul 2> nul
if errorlevel 1 (
  set "PYTHON_EXE=%LocalAppData%\Programs\Python\Python312\python.exe"
  if exist "%PYTHON_EXE%" (
    set "PYINSTALLER_CMD="%PYTHON_EXE%" -m PyInstaller"
  ) else (
    set "PYINSTALLER_CMD=python -m PyInstaller"
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
%PYINSTALLER_CMD% --clean --noconfirm ^
 main.py ^
 --onefile ^
 --noconsole ^
 --name=%APP_NAME% ^
 --icon=%DAKE_ICON% ^
 --add-data "%DAKE_ICON%;."
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
