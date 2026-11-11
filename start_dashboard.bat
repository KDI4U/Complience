@echo off
title Compliance Automation Suite - ESIC and EPFO
setlocal enabledelayedexpansion

cd /d "d:\Complience\dashboard"

echo ================================================================
echo         COMPLIANCE AUTOMATION SUITE (ESIC and EPFO)
echo ================================================================
echo  [1/4] Checking Python environment...

:: Detect Python executable
set "PY_CMD=python"
if exist "C:\Users\PC\AppData\Local\Programs\Python\Python313\python.exe" (
    set "PY_CMD=C:\Users\PC\AppData\Local\Programs\Python\Python313\python.exe"
)

:: Check and start Ollama if not running
echo  [2/4] Checking Ollama AI Service (CAPTCHA Reader)...
powershell -NoProfile -Command "try { $r = Invoke-WebRequest -Uri 'http://127.0.0.1:11434/api/tags' -UseBasicParsing -TimeoutSec 1; if ($r.StatusCode -eq 200) { Write-Host '       [OK] Ollama is online.' } } catch { Write-Host '       [i] Starting Ollama service...'; Start-Process 'ollama' -ArgumentList 'serve' -WindowStyle Hidden; Start-Sleep -Seconds 2 }"

:: Free Port 8008 if lingering from previous run or unclean shutdown
echo  [3/4] Checking Port 8008 availability...
powershell -NoProfile -Command "Get-NetTCPConnection -LocalPort 8008 -ErrorAction SilentlyContinue | ForEach-Object { Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue }"

echo  [4/4] Starting Server on http://127.0.0.1:8008 ...
echo ================================================================
echo  Server is starting up. Browser will open automatically once ready...
echo  Please keep this window open while using the dashboard
echo ================================================================

:: Background browser launcher: waits until server returns HTTP 200 before opening browser
start "" powershell -NoProfile -WindowStyle Hidden -Command "for ($i=0; $i -lt 40; $i++) { Start-Sleep -Milliseconds 800; try { $r = Invoke-WebRequest -Uri 'http://127.0.0.1:8008' -UseBasicParsing -TimeoutSec 1; if ($r.StatusCode -eq 200) { Start-Process 'http://127.0.0.1:8008'; break } } catch {} }"

:: Start Uvicorn Server
"%PY_CMD%" -m uvicorn server:app --host 0.0.0.0 --port 8008

if %ERRORLEVEL% neq 0 (
    echo.
    echo ================================================================
    echo [ERROR] Server stopped with error code %ERRORLEVEL%.
    echo Please check the error messages above.
    echo ================================================================
)
pause
