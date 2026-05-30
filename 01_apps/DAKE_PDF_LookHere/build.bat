@echo off
setlocal

rmdir /s /q build
rmdir /s /q dist
del *.spec
del /q version_info.txt 2>nul

set "PYINSTALLER_CMD=pyinstaller"
where pyinstaller > nul 2>&1
if errorlevel 1 (
  if defined PYTHON_EXE (
    set "PYINSTALLER_CMD=%PYTHON_EXE% -m PyInstaller"
  ) else (
    where py > nul 2>&1
    if not errorlevel 1 (
      set "PYINSTALLER_CMD=py -m PyInstaller"
    ) else (
      set "PYINSTALLER_CMD=python -m PyInstaller"
    )
  )
)

python ..\..\tools\generate_version_info.py --app-dir . --out version_info.txt
if errorlevel 1 (
    echo VersionInfo generation failed.
    pause
    exit /b 1
)

%PYINSTALLER_CMD% ^
--onefile ^
--noconsole ^
--clean ^
--paths=..\..\00_core ^
--icon=..\..\02_assets\dake_icon.ico ^
--version-file version_info.txt ^
--name DakePDF_LookHere ^
main.py

pause
