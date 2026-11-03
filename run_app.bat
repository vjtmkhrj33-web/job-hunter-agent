@echo off
title Autonomous Job Hunter & Application Tracker
cd /d "%~dp0"

echo ====================================================================
echo   Starting Autonomous Job Hunter & Application Tracker Dashboard...
echo   Localhost Address: http://localhost:8080
echo ====================================================================

start http://localhost:8080
python server.py 8080
pause

