@echo off
title SEAZEN CANTEEN KIOSK
cd /d "%~dp0"
echo ===================================================================
echo             SEAZEN CANTEEN CHECK-IN KIOSK SYSTEM
echo ===================================================================
echo.
if exist "SEAZEN_Canteen_Kiosk.exe" (
    "SEAZEN_Canteen_Kiosk.exe"
) else if exist "%LocalAppData%\Programs\Python\Python311\python.exe" (
    "%LocalAppData%\Programs\Python\Python311\python.exe" "Canteen_Kiosk_Runner.py"
) else (
    python "Canteen_Kiosk_Runner.py"
)
pause
