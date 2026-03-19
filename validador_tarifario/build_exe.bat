@echo off
setlocal
cd /d "%~dp0"
pyinstaller --noconsole --onefile --name "Validador Tarifario" --paths src run.py
pause
