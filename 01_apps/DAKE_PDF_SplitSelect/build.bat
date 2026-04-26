@echo off
setlocal
cd /d "%~dp0"

set "PYTHON_EXE="
set "PYTHON_ARGS="
set "APP_NAME=DakePDF_Split_Select"
set "ENTRY_FILE=main.py"
set "OUTPUT_EXE=dist\%APP_NAME%.exe"

if exist "%~dp0.venv\Scripts\python.exe" (
  set "PYTHON_EXE=%~dp0.venv\Scripts\python.exe"
)

if not defined PYTHON_EXE (
  where py >nul 2>nul
  if not errorlevel 1 (
    set "PYTHON_EXE=py"
    set "PYTHON_ARGS=-3"
  )
)

if not defined PYTHON_EXE (
  where python >nul 2>nul
  if not errorlevel 1 (
    set "PYTHON_EXE=python"
  )
)

if not defined PYTHON_EXE goto :python_missing
if not exist "%ENTRY_FILE%" goto :entry_missing

call :run %PYTHON_EXE% %PYTHON_ARGS% -m pip install --upgrade pip
if errorlevel 1 goto :fail

call :run %PYTHON_EXE% %PYTHON_ARGS% -m pip install -r requirements.txt
if errorlevel 1 goto :fail

call :run %PYTHON_EXE% %PYTHON_ARGS% -m PyInstaller --noconfirm --clean --onefile --windowed --name=DakePDF_Split_Select --icon=..\..\02_assets\dake_icon.ico --hidden-import=tkinterdnd2 --collect-all=tkinterdnd2 --hidden-import=fitz --collect-all=fitz "%ENTRY_FILE%"
if errorlevel 1 goto :fail

if not exist "%OUTPUT_EXE%" goto :output_missing

echo.
echo Build completed.
echo Output: %OUTPUT_EXE%
pause
exit /b 0

:python_missing
echo.
echo Python 3 was not found.
echo Install Python or create .venv in this folder.
pause
exit /b 1

:entry_missing
echo.
echo Entry file was not found.
echo Expected: %ENTRY_FILE%
pause
exit /b 1

:run
echo.
echo [RUN] %*
%*
exit /b %errorlevel%

:fail
echo.
echo Build failed.
pause
exit /b 1

:output_missing
echo.
echo Build finished but the exe was not found.
echo Expected: %OUTPUT_EXE%
pause
exit /b 1
