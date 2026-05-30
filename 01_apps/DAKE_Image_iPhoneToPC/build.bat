@echo off
cd /d "%~dp0"

python -c "import sys, qrcode; print(sys.executable); print(qrcode.__file__)"
if errorlevel 1 (
    python -m pip install "qrcode[pil]"
)
python -c "import sys, qrcode; print(sys.executable); print(qrcode.__file__)"
if errorlevel 1 exit /b 1

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
del /q *.spec 2>nul
del /q version_info.txt 2>nul

python ..\..\tools\generate_version_info.py --app-dir . --out version_info.txt
if errorlevel 1 (
    echo VersionInfo generation failed.
    pause
    exit /b 1
)

python -m PyInstaller ^
--onefile ^
--noconsole ^
--clean ^
--name DakeImage_iPhoneToPC ^
--paths=..\..\00_core ^
--icon=..\..\02_assets\dake_icon.ico ^
--version-file version_info.txt ^
--collect-all qrcode ^
--collect-submodules qrcode ^
--collect-data qrcode ^
--collect-all=pillow_heif ^
main.py

pause
