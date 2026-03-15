@echo off
cd /d "%~dp0"
echo Stoppar eventuell tidigare ASK CSV listener pa port 8010...
for /f "tokens=5" %%a in ('netstat -aon ^| findstr ":8010 "') do (
    taskkill /F /PID %%a >nul 2>&1
)
timeout /t 1 >nul
echo Startar ASK CSV listener pa http://127.0.0.1:8010
echo Tryck Ctrl+C for att stoppa.
python -m uvicorn ask_csv_listener:app --host 127.0.0.1 --port 8010
pause

