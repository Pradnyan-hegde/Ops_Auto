@echo off
title Ops_Auto Reconciliation Dashboard Server
cd /d "%~dp0"

echo ======================================================================
echo           Ops_Auto Reconciliation Dashboard Server
echo ======================================================================
echo.

set "VENV_PYTHON=C:\Users\HP\OneDrive - SwinkPay Fintech Pvt Ltd\Pradnyan\Testing_Tool\backend\.venv\Scripts\python.exe"

if exist "%VENV_PYTHON%" (
    echo [INFO] Found virtual environment Python.
    echo [INFO] Launching dashboard at http://127.0.0.1:8000 ...
    start "" http://127.0.0.1:8000
    echo [INFO] Server running. Press Ctrl+C in this window to stop.
    echo.
    "%VENV_PYTHON%" -m uvicorn dashboard.server:app --host 127.0.0.1 --port 8000 --reload
) else (
    echo [INFO] Using default system Python.
    echo [INFO] Launching dashboard at http://127.0.0.1:8000 ...
    start "" http://127.0.0.1:8000
    echo [INFO] Server running. Press Ctrl+C in this window to stop.
    echo.
    python -m uvicorn dashboard.server:app --host 127.0.0.1 --port 8000 --reload
)

pause
