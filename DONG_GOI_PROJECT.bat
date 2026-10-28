@echo off
title DONG GOI SEAZEN CANTEEN KIOSK STANDALONE
cd /d "%~dp0"
echo ===================================================================
echo   DONG GOI SEAZEN CANTEEN KIOSK - BAN DOC LAP 1 FILE EXE DUY NHAT
echo ===================================================================
echo.

if exist "%LocalAppData%\Programs\Python\Python311\python.exe" (
    set PY_EXE="%LocalAppData%\Programs\Python\Python311\python.exe"
) else (
    set PY_EXE=py
)

echo [*] Dang bien dich va dong goi bang PyInstaller qua package_release.py...
%PY_EXE% package_release.py

REM Don dep thu muc build tam
if exist "build" rmdir /s /q "build"

echo.
echo ===================================================================
echo  HOAN TAT! Ban chi can gui duy nhat 1 file EXE:
echo  %~dp0SEAZEN_Canteen_Kiosk.exe
echo ===================================================================
pause
