@echo off
cd /d "%~dp0"

echo Starting Product Matching Engine...
echo.

REM Check if virtual environment exists
if not exist ".venv\Scripts\activate.bat" (
    echo ERROR: Virtual environment not found!
    echo Please run setup.bat first to create the virtual environment.
    pause
    exit /b 1
)

REM Activate virtual environment
call .venv\Scripts\activate.bat

REM Set environment variables to disable telemetry
set STREAMLIT_BROWSER_GATHER_USAGE_STATS=false
set STREAMLIT_TELEMETRY_OPT_OUT=true

REM Start Streamlit and open browser
echo Starting Streamlit server...
start "" "http://localhost:8501"
streamlit run app.py --server.port=8501 --global.developmentMode=false --browser.gatherUsageStats=false --server.headless=true

pause
