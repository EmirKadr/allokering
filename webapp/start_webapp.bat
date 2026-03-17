@echo off
cd /d "%~dp0"
if "%ALLOK_DATABASE_URL%"=="" if "%DATABASE_URL%"=="" (
    echo Saknar ALLOK_DATABASE_URL eller DATABASE_URL.
    echo Satt PostgreSQL-anslutningen innan du startar webappen.
    pause
    exit /b 1
)
echo Stoppar eventuell tidigare instans pa port 8000...
for /f "tokens=5" %%a in ('netstat -aon ^| findstr ":8000 "') do (
    taskkill /F /PID %%a >nul 2>&1
)
timeout /t 1 >nul
echo Startar Allokering WebApp pa http://localhost:8000
echo Tryck Ctrl+C for att stoppa.
python -m uvicorn backend.main:app --host 0.0.0.0 --port 8000
pause
