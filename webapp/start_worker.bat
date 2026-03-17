@echo off
cd /d "%~dp0"
if "%ALLOK_DATABASE_URL%"=="" if "%DATABASE_URL%"=="" (
    echo Saknar ALLOK_DATABASE_URL eller DATABASE_URL.
    echo Satt PostgreSQL-anslutningen innan du startar workern.
    pause
    exit /b 1
)
echo Startar Allokering Worker...
echo Tryck Ctrl+C for att stoppa.
python -m backend.worker
pause
