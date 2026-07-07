@echo off
:: Bu skript FAQAT SERVERDA (Mega Max 7002, D:\Nerjaveyka_Project) ishga tushiriladi.
:: Vazifasi: eng oxirgi kodni GitHub'dan olib, botni qayta ishga tushiradi.

cd /d "D:\Nerjaveyka_Project"

echo ============================================
echo  NEJAVIYKA server yangilanmoqda...
echo ============================================
git pull origin main

echo.
echo ============================================
echo  Bot qayta ishga tushirilmoqda (pm2)...
echo ============================================
pm2 restart nejaviyka-bot

echo.
timeout /t 3 /nobreak >nul
echo ============================================
echo  Oxirgi loglar (xato bor-yo'qligini tekshiring):
echo ============================================
pm2 logs nejaviyka-bot --lines 20 --nostream

echo.
echo ============================================
echo  Tayyor.
echo ============================================
pause
