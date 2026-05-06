@echo off
setlocal

cd /d "%~dp0"

python -m streamlit run analytics_dashboard.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo If Streamlit is missing, install analytics dependencies first:
    echo   python -m pip install -r requirements-analytics.txt
    echo.
    echo Could not start with "python". Trying "py -3"...
    py -3 -m streamlit run analytics_dashboard.py
)

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Failed to start analytics_dashboard.py
    pause
)

endlocal
