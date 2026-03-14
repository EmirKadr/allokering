@echo off
setlocal

cd /d "%~dp0"

python allokera11_pyqt6.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Could not start with "python". Trying "py -3"...
    py -3 allokera11_pyqt6.py
)

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Failed to start allokera11_pyqt6.py.
    pause
)

endlocal
