@echo off
title Brochure Hub 2027 - Bakalia
cd /d "%~dp0"

set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8
set STREAMLIT_BROWSER_GATHER_USAGE_STATS=false
set STREAMLIT_SERVER_HEADLESS=false

set MISTRAL_STREAMLIT=%USERPROFILE%\.mistral_env\Scripts\streamlit.exe

if not exist "%MISTRAL_STREAMLIT%" (
    echo [ERROR] Streamlit executable not found at:
    echo %MISTRAL_STREAMLIT%
    pause
    exit /b 1
)

echo ===============================================================================
echo   BAKALIA - BROCHURE HUB 2027
echo ===============================================================================
echo   Starting application at: http://localhost:8502
echo ===============================================================================

start "" "http://localhost:8502"
"%MISTRAL_STREAMLIT%" run "%~dp0brochure_hub_app.py" --server.port=8502 --server.address=0.0.0.0 --browser.gatherUsageStats=false

if errorlevel 1 (
    echo [ERROR] An error occurred while running the application.
    pause
)
