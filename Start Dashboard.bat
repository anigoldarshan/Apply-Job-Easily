@echo off
REM Opens the Job Bot dashboard in your browser. Keep this window open while you use it.
cd /d "%~dp0"
title Job Bot Dashboard
python dashboard.py
if errorlevel 1 (
  echo.
  echo  The dashboard couldn't start. If it says "No module named ...", run Install.bat first.
  pause
)
