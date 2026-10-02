@echo off
cd /d "%~dp0"
title Sweet
echo Checking Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo Python is not installed or not on PATH.
    echo Install it from https://www.python.org/downloads/ and tick "Add Python to PATH",
    echo then double-click this file again.
    echo.
    pause
    exit /b 1
)
echo Installing needed packages (first time takes a minute)...
python -m pip install --quiet google-genai Pillow keyboard pyautogui pyperclip
if errorlevel 1 (
    echo.
    echo Package install failed. Copy the message above and send it to Claude.
    pause
    exit /b 1
)
echo.
echo Starting Sweet. Keep this window open to see messages. Ctrl+Shift+Q quits.
echo.
python sweet.py --taskbar
echo.
echo Sweet stopped. If there is an error above, send it to Claude.
pause
