@echo off
title DANG NHAP / THAY DOI TAI KHOAN CANTEEN KIOSK
cd /d "%~dp0"
echo ===================================================================
echo     DANG NHAP / THAY DOI TAI KHOAN XAC THUC CANTEEN KIOSK
echo ===================================================================
echo.
if exist "%LocalAppData%\Programs\Python\Python311\python.exe" (
    "%LocalAppData%\Programs\Python\Python311\python.exe" "tao_token_moi.py"
) else (
    py "tao_token_moi.py"
)
pause
