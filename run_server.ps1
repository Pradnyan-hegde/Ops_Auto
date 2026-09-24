# Ops_Auto Reconciliation Dashboard Server Runner
Set-Location $PSScriptRoot

Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host "           Ops_Auto Reconciliation Dashboard Server" -ForegroundColor Cyan
Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host ""

$venvPython = "C:\Users\HP\OneDrive - SwinkPay Fintech Pvt Ltd\Pradnyan\Testing_Tool\backend\.venv\Scripts\python.exe"

Start-Process "http://127.0.0.1:8000"

if (Test-Path $venvPython) {
    Write-Host "[INFO] Starting server with Python virtualenv on http://127.0.0.1:8000..." -ForegroundColor Green
    & $venvPython -m uvicorn dashboard.server:app --host 127.0.0.1 --port 8000 --reload
} else {
    Write-Host "[INFO] Starting server with default python on http://127.0.0.1:8000..." -ForegroundColor Green
    python -m uvicorn dashboard.server:app --host 127.0.0.1 --port 8000 --reload
}
