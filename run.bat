@echo off
cd /d "%~dp0"
if not exist ".venv" (
  echo Creating virtual environment...
  python -m venv .venv
  call .venv\Scripts\activate.bat
  pip install --upgrade pip
  pip install -r requirements.txt
) else (
  call .venv\Scripts\activate.bat
)
echo.
echo Opening http://127.0.0.1:5050
start "" "http://127.0.0.1:5050"
python app.py
pause
