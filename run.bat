@echo off
cd /d "%~dp0"
echo.
echo   ASK MY CFO — M1 Automation (Flask)
echo.
pip install -r requirements.txt -q
python app.py
pause
