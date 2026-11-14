@echo off
title WSDOT SMTP & TCRS Interactive Gantt Chart
cd /d "%~dp0"

echo ==============================================================
echo  Starting WSDOT SMTP & TCRS Interactive Gantt Chart
echo ==============================================================
echo.
echo Checking for Python sync server...

where python >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    echo Python detected. Starting live two-way Excel sync server...
    echo Browser will open automatically at http://localhost:8080
    echo (Leave this window open to keep live disk sync active)
    echo.
    python sync_server.py
) else (
    echo Python not found in PATH. Opening standalone Gantt Chart in browser...
    start "" "%~dp0index.html"
)
pause

