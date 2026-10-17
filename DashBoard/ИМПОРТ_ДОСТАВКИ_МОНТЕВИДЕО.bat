@echo off
chcp 65001 > nul
echo ============================================================================
echo   ИМПОРТ НА ДОСТАВКИ В МИСТРАЛ ERP — МОНТЕВИДЕО
echo   Верига магазини „Бакалия“ | ЦЕНТРАЛЕН ОФИС
echo ============================================================================
echo.

set PYTHON_CMD=python
where python >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    set PYTHON_CMD="C:\Users\Tonev\AppData\Local\Programs\Python\Python313\python.exe"
)

echo [1] Стартиране на предварителен преглед (Dry-Run)...
echo ----------------------------------------------------------------------------
%PYTHON_CMD% "g:\My Drive\AntiG_test\import_deliveries_montevideo.py"
echo ----------------------------------------------------------------------------
echo.

set /p CONFIRM="Желаете ли да изпълните РЕАЛЕН ЗАПИС (Commit) в базата Монтевидео? (Y/N): "
if /i "%CONFIRM%"=="Y" (
    echo.
    echo [2] Изпълнение на реален запис (Live Commit)...
    %PYTHON_CMD% "g:\My Drive\AntiG_test\import_deliveries_montevideo.py" --commit
    echo.
    echo Резултатът е записан и в одитния Excel файл.
) else (
    echo.
    echo Операцията е отказана. Базата данни остава непокътната.
)

echo.
pause
