@echo off
setlocal
cd /d "%~dp0"
echo Findcontracts launcher v2
echo Folder: %CD%
if not exist launch.py (
  echo ERROR: launch.py is missing. Extract the ENTIRE latest ZIP into a new folder.
  pause
  exit /b 1
)
if not exist app.py (
  echo ERROR: app.py is missing. Extract the ENTIRE ZIP.
  pause
  exit /b 1
)
if not exist index.html (
  echo ERROR: index.html is missing. Extract the ENTIRE ZIP.
  pause
  exit /b 1
)
python -m pip install -r requirements.txt
if errorlevel 1 (
  echo Dependency installation failed.
  pause
  exit /b 1
)
python "%~dp0launch.py"
pause
