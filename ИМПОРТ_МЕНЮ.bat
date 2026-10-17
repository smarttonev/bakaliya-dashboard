@echo off
chcp 65001 >nul
python "%~dp0mistral_promotions_importer.py"
pause
