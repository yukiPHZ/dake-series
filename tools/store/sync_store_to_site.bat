@echo off
cd /d %~dp0\..\..
python tools\store\sync_store_to_site.py %*
