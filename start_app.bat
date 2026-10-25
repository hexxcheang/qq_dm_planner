@echo off
rem Double-click to start the Invoice & DO generator on Windows.
cd /d "%~dp0"
title DM Optical - Invoice and DO generator

if not exist ".venv\Scripts\python.exe" (
    echo First run: setting up, this takes a few minutes...
    py -3 -m venv .venv 2>nul || python -m venv .venv
    if not exist ".venv\Scripts\python.exe" (
        echo.
        echo Python was not found. Install it from https://www.python.org/downloads/
        echo and tick "Add python.exe to PATH" on the first screen of the installer.
        pause
        exit /b 1
    )
)
".venv\Scripts\python.exe" -m pip install --disable-pip-version-check -q -r requirements.txt
if errorlevel 1 (
    echo.
    echo Installing the required packages failed. Check your internet connection and try again.
    pause
    exit /b 1
)

echo.
echo The app is starting and will open in your browser at http://localhost:8501
echo Keep this window open while you use the app. Close it to stop the app.
echo.
start "" cmd /c "timeout /t 6 >nul & start http://localhost:8501"
".venv\Scripts\python.exe" -m streamlit run app.py
pause
