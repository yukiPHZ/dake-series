@echo off
chcp 65001 > nul
setlocal
cd /d "%~dp0"

set "APP_NAME=DakeTwoPerson_Memo"
set "DIST_DIR=dist"
set "BUILD_DIR=build"
set "SPEC_FILE=%APP_NAME%.spec"
set "OUTPUT_EXE=%DIST_DIR%\%APP_NAME%.exe"
set "DAKE_ICON=..\..\02_assets\dake_icon.ico"
set "PYTHON_EXE=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"

echo [1/3] Cleaning previous build artifacts...
if exist "%BUILD_DIR%" rmdir /s /q "%BUILD_DIR%"
if exist "%DIST_DIR%" rmdir /s /q "%DIST_DIR%"
if exist "%SPEC_FILE%" del /q "%SPEC_FILE%"
if exist "main.spec" del /q "main.spec"

if exist "%BUILD_DIR%" goto :clean_error
if exist "%DIST_DIR%" goto :clean_error

echo [2/3] Building %APP_NAME%.exe...
if exist "%PYTHON_EXE%" (
  "%PYTHON_EXE%" -m PyInstaller --clean --noconfirm ^
   main.py ^
   --onefile ^
   --noconsole ^
   --name=%APP_NAME% ^
   --icon=%DAKE_ICON% ^
   --add-data "%DAKE_ICON%;."
) else (
  pyinstaller --clean --noconfirm ^
   main.py ^
   --onefile ^
   --noconsole ^
   --name=%APP_NAME% ^
   --icon=%DAKE_ICON% ^
   --add-data "%DAKE_ICON%;."
)
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
