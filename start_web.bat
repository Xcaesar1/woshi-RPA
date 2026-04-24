@echo off
setlocal
title Lingxing Web Server
cd /d "%~dp0"
echo [Lingxing] Starting web server...
echo [Lingxing] Open in browser:
echo [Lingxing]   Local:  http://127.0.0.1:8000/tasks/new
echo [Lingxing]   LAN:    http://192.168.31.101:8000/tasks/new
echo.
python -m uvicorn app.main:app --host 0.0.0.0 --port 8000
echo.
echo [Lingxing] Web server stopped.
pause
