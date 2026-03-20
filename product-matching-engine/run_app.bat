@echo off
title Product Matching Engine

echo.
echo ========================================
echo    Product Matching Engine
echo ========================================
echo.
echo Starting application...
echo.

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.7+ and try again
    pause
    exit /b 1
)

REM Check if virtual environment exists
if not exist ".venv\Scripts\activate.bat" (
    echo Setting up virtual environment...
    python -m venv .venv
    call .venv\Scripts\activate.bat
    pip install -r requirements.txt
) else (
    call .venv\Scripts\activate.bat
)

echo.
echo Opening Product Matching Engine in your browser...
echo Close this window to stop the application.
echo.

REM Start Streamlit
streamlit run app.py --server.headless=true

pause
