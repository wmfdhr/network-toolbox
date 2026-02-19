@echo off
chcp 65001 >nul
cd /d "%~dp0"
"D:\python all\python 3.13.3\pythonw.exe" "updater.py"
start "" "D:\python all\python 3.13.3\pythonw.exe" "app_main.py"
