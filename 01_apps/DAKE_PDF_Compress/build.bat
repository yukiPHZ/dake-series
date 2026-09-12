@echo off
setlocal
cd /d "%~dp0"
set "PYTHON_CMD=python"
set "PYTHON_ARGS="
if defined DAKE_PYTHON set "PYTHON_CMD=%DAKE_PYTHON%"
"%PYTHON_CMD%" --version >nul 2>&1
if errorlevel 1 (
    set "PYTHON_CMD=py"
    set "PYTHON_ARGS=-3"
)
"%PYTHON_CMD%" %PYTHON_ARGS% -c "import pymupdf,tkinterdnd2; from adaptive import VERIFIED_PYMUPDF; assert pymupdf.__version__ == VERIFIED_PYMUPDF; assert callable(pymupdf.Document.rewrite_images)"
if errorlevel 1 exit /b 1
"%PYTHON_CMD%" %PYTHON_ARGS% ..\..\tools\generate_version_info.py --app-dir . --out version_info.txt
if errorlevel 1 exit /b 1
rem Do not recursively delete unrelated dist files.
"%PYTHON_CMD%" %PYTHON_ARGS% -m PyInstaller ^
--onefile ^
--noconsole ^
--clean ^
--noconfirm ^
--paths=..\..\00_core ^
--icon=..\..\02_assets\dake_icon.ico ^
--add-data "..\..\02_assets\dake_icon.ico;." ^
--version-file version_info.txt ^
--collect-data=tkinterdnd2 ^
--exclude-module pandas ^
--exclude-module numpy ^
--exclude-module PIL ^
--exclude-module openpyxl ^
--exclude-module lxml ^
--exclude-module matplotlib ^
--exclude-module scipy ^
--name DakePDF_Compress ^
main.py
set "BUILD_RESULT=%ERRORLEVEL%"
if not "%~1"=="--no-pause" pause
exit /b %BUILD_RESULT%
