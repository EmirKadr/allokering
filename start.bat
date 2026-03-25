@echo off
setlocal

cd /d "%~dp0"

python allokering12.1.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Could not start with "python". Trying "py -3"...
    py -3 allokering12.1.py
)

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Failed to start allokering12.1.py
    pause
)

endlocal
