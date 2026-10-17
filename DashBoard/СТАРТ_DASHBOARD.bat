@echo off
chcp 65001 >nul
title Бакалия — Главен оперативен Dashboard
cd /d "%~dp0"

set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8
set STREAMLIT_BROWSER_GATHER_USAGE_STATS=false
set STREAMLIT_SERVER_HEADLESS=false

echo ===============================================================================
echo   ВЕРИГА СУПЕРМАРКЕТИ „БАКАЛИЯ“ — ЕДИНЕН ОПЕРАТИВЕН DASHBOARD
echo ===============================================================================
echo   Подготовка на работната среда и стартиране на уеб таблото...
echo ===============================================================================
echo.

:: 1. Проверка за Streamlit credentials
if not exist "%USERPROFILE%\.streamlit" mkdir "%USERPROFILE%\.streamlit"
if not exist "%USERPROFILE%\.streamlit\credentials.toml" (
    echo [general]> "%USERPROFILE%\.streamlit\credentials.toml"
    echo email = "">> "%USERPROFILE%\.streamlit\credentials.toml"
)

:: 2. Локализиране на Streamlit в .mistral_env
set MISTRAL_STREAMLIT=%USERPROFILE%\.mistral_env\Scripts\streamlit.exe

if not exist "%MISTRAL_STREAMLIT%" (
    echo [i] Подготовка на Python среда...
    if exist "%~dp0..\bin\uv.exe" (
        "%~dp0..\bin\uv.exe" venv "%USERPROFILE%\.mistral_env" --python 3.12
        "%~dp0..\bin\uv.exe" pip install --python "%USERPROFILE%\.mistral_env\Scripts\python.exe" streamlit pandas openpyxl firebirdsql dulwich requests
    )
)

if not exist "%MISTRAL_STREAMLIT%" (
    echo [ERROR] Не беше открит Streamlit в %MISTRAL_STREAMLIT%
    pause
    exit /b 1
)

echo [OK] Стартиране на Dashboard в браузъра...
echo ===============================================================================
echo   Приложението е достъпно на адрес:
echo   Локално: http://localhost:8501
echo   Вътрешна мрежа: http://192.168.5.99:8501 или http://192.168.5.254:8501
echo ===============================================================================

start "" "http://localhost:8501"
"%MISTRAL_STREAMLIT%" run "%~dp0dashboard_app.py" --server.port=8501 --server.address=0.0.0.0 --browser.gatherUsageStats=false

if errorlevel 1 (
    echo.
    echo [ERROR] Възникна грешка при изпълнението.
    pause
)
