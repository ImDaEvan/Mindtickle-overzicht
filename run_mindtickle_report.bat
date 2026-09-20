@echo off
cd /d "%~dp0"
python mindtickle_report.py
if errorlevel 1 (
    echo.
    echo Failed to run mindtickle_report.py.
    echo Make sure Python is installed and available on PATH.
    pause
)
