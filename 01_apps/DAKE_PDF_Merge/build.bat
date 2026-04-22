@echo off
cd /d %~dp0
pyinstaller --noconfirm --onefile --windowed --name DAKE_PDF_Merge --icon app.ico --add-data "app.ico;." --add-data "icon.png;." --hidden-import=tkinterdnd2 --collect-all tkinterdnd2 main.py
exit
