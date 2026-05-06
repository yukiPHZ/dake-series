@echo off
setlocal

rmdir /s /q build
rmdir /s /q dist
del *.spec

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

%PYINSTALLER_CMD% ^
--onefile ^
--noconsole ^
--clean ^
--icon=..\..\02_assets\dake_icon.ico ^
--name DakePDF_LookHere ^
main.py

pause
