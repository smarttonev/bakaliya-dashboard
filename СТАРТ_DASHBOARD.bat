@echo off
chcp 65001 >nul
title Бакалия — Главен оперативен Dashboard
cd /d "%~dp0DashBoard"

call "%~dp0DashBoard\СТАРТ_DASHBOARD.bat"
