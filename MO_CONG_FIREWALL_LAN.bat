@echo off
title SEAZEN CANTEEN - MO CONG TUONG LUA MANG LAN (PORT 5000)
chcp 65001 >nul
cd /d "%~dp0"

echo ===================================================================
echo    CONG CU MO CONG TUONG LUA WINDOWS CHO HE THONG CANTEEN KIOSK
echo    (Cho phep cac may khac trong mang Wi-Fi/LAN truy cap vao Kiosk)
echo ===================================================================
echo.

:: Kiem tra quyen Administrator
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo [*] Dang yeu cau quyen Administrator de cau hinh tuong lua...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

echo [*] Dang them quy tac cho phep ket noi vao Port 5000...
netsh advfirewall firewall delete rule name="SEAZEN_Canteen_Kiosk_Port_5000" >nul 2>&1
netsh advfirewall firewall add rule name="SEAZEN_Canteen_Kiosk_Port_5000" dir=in action=allow protocol=TCP localport=5000 >nul

if %errorlevel% equ 0 (
    echo.
    echo ===================================================================
    echo  [+] THANH CONG: Da mo cong 5000 tren Windows Firewall!
    echo ===================================================================
    echo.
    echo  Bay gio bat ky thiet bi nao (may tinh, dien thoai, may tinh bang)
    echo  trong cung mang Wi-Fi / LAN deu co the truy cap vao Kiosk nay.
    echo.
) else (
    echo.
    echo [-] Khong the them quy tac Firewall. Vui long kiem tra lai quyen Admin!
    echo.
)

pause
