@echo off
cd /d "%~dp0"
echo Startar Allokering WebApp pa http://localhost:8000
echo Tryck Ctrl+C for att stoppa.
python -m uvicorn backend.main:app --host 0.0.0.0 --port 8000
pause
