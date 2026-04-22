@echo off
cd /d %~dp0
pyinstaller --noconfirm --onefile --windowed --name ST_PDF_Merge_Only --icon st_pdf_merge_only.ico --add-data "st_pdf_merge_only.ico;." --add-data "s512_f_object_154_2bg.png;." --hidden-import=tkinterdnd2 --collect-all tkinterdnd2 st_pdf_merge_only_v7.py
exit
