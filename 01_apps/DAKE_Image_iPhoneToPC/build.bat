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

python -m PyInstaller ^
--onefile ^
--noconsole ^
--clean ^
--name DakeImage_iPhoneToPC ^
--icon=..\..\02_assets\dake_icon.ico ^
--collect-all qrcode ^
--collect-submodules qrcode ^
--collect-data qrcode ^
--collect-all=pillow_heif ^
main.py

pause
