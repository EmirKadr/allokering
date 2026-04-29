@echo off
setlocal

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1"
if errorlevel 1 (
    echo.
    echo Installationen misslyckades.
    pause
    exit /b 1
)

echo.
echo Klar. Du kan starta Allokering fran skrivbordet eller Start-menyn.
pause
