@echo off
setlocal enabledelayedexpansion

cd /d "%~dp0"

set "FOUND="
for /f "delims=" %%F in ('dir /b /a:-d /o:-d "v_ask_article_buffertpallet*" 2^>nul') do (
    if not defined FOUND set "FOUND=%%F"
)

if not defined FOUND (
    echo Hittade ingen fil som borjar med v_ask_article_buffertpallet i %~dp0
    pause
    exit /b 1
)

echo Anvander indatafil: %FOUND%

python "%~dp0artikel_max.py" "%FOUND%"


if errorlevel 1 (
    echo Python-skriptet misslyckades.
    pause
    exit /b 1
)

pause
