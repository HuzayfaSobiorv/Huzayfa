@echo off
cd /d "C:\Users\PC_HP\Desktop\NEJAVIYKA"

:restart
echo.
echo ============================================
echo  NEJAVIYKA Bot ishga tushirilmoqda...
echo ============================================

:: Eski bot jarayonlarini to'xtatish
taskkill /F /FI "IMAGENAME eq python*" >nul 2>&1
taskkill /F /FI "IMAGENAME eq python3.14*" >nul 2>&1
timeout /t 2 /nobreak >nul

:: .pyc kesh tozalash (eski kod yuklanmasligi uchun)
if exist __pycache__ (
    rmdir /s /q __pycache__
    echo [OK] Kesh tozalandi
)

:: Botni ishga tushirish (start yo'q - to'g'ridan-to'g'ri, to'xtagan vaqtda loop ishlaydi)
echo [OK] Bot ishga tushdi -- %date% %time%
python Bot.py

:: Bot to'xtadi (xato yoki WiFi uzildi)
echo.
echo [!!] Bot to'xtadi. Sabab: xato yoki tarmoq uzilishi.
echo [..] 10 soniyada qayta ishga tushadi...
echo      (To'xtatish uchun Ctrl+C bosing)
echo.
timeout /t 10 /nobreak >nul
goto restart
