@echo off
cd /d "%~dp0"
if not exist index.html (
  echo ERROR: index.html is missing. Extract the entire ZIP before running start.bat.
  pause
  exit /b 1
)
python -m pip install -r requirements.txt
if errorlevel 1 (
  pause
  exit /b 1
)
start "" http://127.0.0.1:5050/
python app.py
pause
