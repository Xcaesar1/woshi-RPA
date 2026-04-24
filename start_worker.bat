@echo off
setlocal
title Lingxing Task Worker
cd /d "%~dp0"
echo [Lingxing] Starting task worker...
echo [Lingxing] Keep this window open while tasks are running.
echo.
python -m app.workers.task_worker
echo.
echo [Lingxing] Task worker stopped.
pause
