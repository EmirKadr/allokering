@echo off
setlocal

cd /d "%~dp0"

python allokera11.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Could not start with "python". Trying "py -3"...
    py -3 allokera11.py
)

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Failed to start allokera11.py.
    pause
)

endlocal
