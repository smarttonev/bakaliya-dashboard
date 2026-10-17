@echo off
chcp 65001 >nul
title Mistral ERP - Web Price Importer (Bakalia)
cd /d "%~dp0"

set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8
set STREAMLIT_BROWSER_GATHER_USAGE_STATS=false
set STREAMLIT_SERVER_HEADLESS=false

echo ===============================================================================
echo   MISTRAL ERP - WEB PRICE IMPORTER (BAKALIA)
echo ===============================================================================
echo   Preparing environment and starting application...
echo ===============================================================================
echo.

:: 1. Ensure Streamlit credentials exist so it never prompts for email
if not exist "%USERPROFILE%\.streamlit" mkdir "%USERPROFILE%\.streamlit"
if not exist "%USERPROFILE%\.streamlit\credentials.toml" (
    echo [general]> "%USERPROFILE%\.streamlit\credentials.toml"
    echo email = "">> "%USERPROFILE%\.streamlit\credentials.toml"
)

:: 2. Check if local virtual environment already exists on this PC
set MISTRAL_STREAMLIT=%USERPROFILE%\.mistral_env\Scripts\streamlit.exe
if exist "%MISTRAL_STREAMLIT%" goto launch

:: 3. Locate uv.exe
set UV_EXE=
if exist "%~dp0bin\uv.exe" set UV_EXE=%~dp0bin\uv.exe
if not defined UV_EXE if exist "%USERPROFILE%\.local\bin\uv.exe" set UV_EXE=%USERPROFILE%\.local\bin\uv.exe
if not defined UV_EXE (
    where uv >nul 2>&1
    if not errorlevel 1 set UV_EXE=uv
)

:: 4. If uv.exe is not found, auto-download it
if not defined UV_EXE (
    echo [i] First-time setup: Downloading portable launcher...
    powershell -NoProfile -ExecutionPolicy Bypass -Command "irm https://astral.sh/uv/install.ps1 | iex"
    if exist "%USERPROFILE%\.local\bin\uv.exe" set UV_EXE=%USERPROFILE%\.local\bin\uv.exe
)

if not defined UV_EXE (
    echo [ERROR] Could not find or download portable launcher.
    pause
    exit /b 1
)

:: 5. Create local Python 3.12 environment on this PC (one-time setup)
echo [i] First-time setup on this PC: preparing Python 3.12 and dependencies...
echo     (This will take about 10-20 seconds only the first time)
"%UV_EXE%" venv "%USERPROFILE%\.mistral_env" --python 3.12
if errorlevel 1 (
    echo [ERROR] Failed to create virtual environment.
    pause
    exit /b 1
)

echo [i] Installing application packages (streamlit, pandas, openpyxl, firebirdsql, passlib)...
"%UV_EXE%" pip install --python "%USERPROFILE%\.mistral_env\Scripts\python.exe" streamlit pandas openpyxl firebirdsql passlib
if errorlevel 1 (
    echo [ERROR] Failed to install packages.
    pause
    exit /b 1
)

:launch
if not exist "%MISTRAL_STREAMLIT%" (
    echo [ERROR] Streamlit executable not found at %MISTRAL_STREAMLIT%
    pause
    exit /b 1
)

echo [OK] Starting application in default browser...
echo ===============================================================================
echo   Приложението е достъпно в мрежата за всички компютри на адрес:
echo   http://192.168.5.99:8501
echo ===============================================================================
start "" http://localhost:8501
"%MISTRAL_STREAMLIT%" run "%~dp0mistral_importer_app.py" --server.port=8501 --server.address=0.0.0.0 --browser.gatherUsageStats=false

if errorlevel 1 (
    echo.
    echo ===============================================================================
    echo   [ERROR] An error occurred while running the application.
    echo ===============================================================================
    pause
)
