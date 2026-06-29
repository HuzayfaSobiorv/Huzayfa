@echo off
chcp 65001 >nul
title NEJAVIYKA — Ishga tushirilmoqda...

cd /d "%~dp0"
set PYTHONIOENCODING=utf-8

echo ============================================================
echo  NEJAVIYKA tizimi ishga tushirilmoqda
echo ============================================================
echo.

:: 1. main.py — Power BI yangilanadi
echo [1/2] main.py ishga tushirilmoqda (Power BI yangilanmoqda)...
python main.py
if %errorlevel% neq 0 (
    echo.
    echo [XATO] main.py da xatolik yuz berdi!
    echo Bot ishga tushirilmaydi. Xatoni tuzating va qayta urinib ko'ring.
    pause
    exit /b 1
)

echo.
echo [OK] Power BI yangilandi!
echo.

:: 2. Bot.py — Telegram bot ishga tushadi
echo [2/2] Bot.py ishga tushirilmoqda...
echo.
title NEJAVIYKA — Bot ishlayapti
python Bot.py

echo.
echo Bot to'xtatildi.
pause
