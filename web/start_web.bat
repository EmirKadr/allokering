@echo off
REM Bygger React-frontenden och startar allokerings-demon i ett pywebview-fonster.
cd /d "%~dp0"

echo === Installerar/bygger frontend ===
cd frontend
if not exist node_modules (
  call npm install
)
call npm run build
cd ..

echo === Startar appen ===
python backend\desktop.py
pause
