@echo off
chcp 65001 >nul
title BAKALIA - Bank Reconciliation DSK (Montevideo)

REM 1. Set working directory to project folder
set "WORK_DIR=g:\My Drive\AntiG_test"
if not exist "%WORK_DIR%" set "WORK_DIR=%~dp0"
cd /d "%WORK_DIR%"

set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8

REM 2. Locate Python executable in .mistral_env
set "PYTHON_EXE=%USERPROFILE%\.mistral_env\Scripts\python.exe"

if not exist "%PYTHON_EXE%" (
    where python >nul 2>&1
    if not errorlevel 1 (
        set "PYTHON_EXE=python"
    ) else (
        echo [ERROR] Python environment was not found on this computer.
        pause
        exit /b 1
    )
)

REM 3. Run reconciliation script (supports drag and drop)
"%PYTHON_EXE%" "%WORK_DIR%\run_bank_reconciliation.py" %*

if errorlevel 1 (
    echo.
    echo ===============================================================================
    echo   [ERROR] An error occurred while executing the script.
    echo ===============================================================================
    pause
)
