@echo off
rem Launcher for Document Formatter (ASCII only, avoids codepage issues)
cd /d "%~dp0"
if not exist "venv\Scripts\pythonw.exe" goto setup
start "" "venv\Scripts\pythonw.exe" main.py
exit /b 0

:setup
echo [Setup] Python venv not found. Run these commands first:
echo     python -m venv venv
echo     venv\Scripts\python -m pip install -r requirements.txt
pause
exit /b 1
