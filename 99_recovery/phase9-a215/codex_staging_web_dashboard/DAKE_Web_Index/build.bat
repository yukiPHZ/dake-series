@echo off
setlocal

rmdir /s /q build 2>nul
rmdir /s /q dist 2>nul
del /q *.spec 2>nul
del /q version_info.txt 2>nul

set "SERIES_ROOT=C:\Users\yukiz\devlop\DAKE_series"
set "VERSION_TOOL=%SERIES_ROOT%\tools\generate_version_info.py"
set "ICON_PATH=%SERIES_ROOT%\02_assets\dake_icon.ico"
set "ICON_OPTION="
set "VERSION_OPTION="

if exist "%VERSION_TOOL%" (
  python "%VERSION_TOOL%" --app-dir . --out version_info.txt
  if errorlevel 1 (
    echo VersionInfo generation failed.
    exit /b 1
  )
  set "VERSION_OPTION=--version-file version_info.txt"
)

if exist "%ICON_PATH%" set "ICON_OPTION=--icon=%ICON_PATH%"

pyinstaller ^
--onefile ^
--noconsole ^
--clean ^
--name DakeWeb_Index ^
%ICON_OPTION% ^
%VERSION_OPTION% ^
main.py

exit /b %ERRORLEVEL%
