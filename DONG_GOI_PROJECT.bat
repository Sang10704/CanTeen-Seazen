@echo off
title DONG GOI SEAZEN CANTEEN KIOSK
cd /d "%~dp0"
echo ===================================================================
echo             DONG GOI SEAZEN CANTEEN KIOSK DE GUI MAY KHAC
echo ===================================================================
echo.

if exist "%LocalAppData%\Programs\Python\Python311\python.exe" (
    set PY_EXE="%LocalAppData%\Programs\Python\Python311\python.exe"
) else (
    set PY_EXE=py
)

echo [1/2] Dang bien dich file chay SEAZEN_Canteen_Kiosk.exe qua PyInstaller...
%PY_EXE% -m PyInstaller --clean --noconfirm --onefile --name "SEAZEN_Canteen_Kiosk" --icon "canteen_icon.ico" --add-data "Canteen_Kiosk_Production.html;." --add-data "token_cache.bin;." --add-data "mp3;mp3" Canteen_Kiosk_Runner.py

echo.
echo [2/2] Dang dong goi tat ca file va nen thanh SEAZEN_Canteen_Kiosk_Full_Package.zip...
%PY_EXE% package_release.py

REM Don dep thu muc build tam cua PyInstaller de thu muc luon sach se
if exist "build" rmdir /s /q "build"

echo.
echo ===================================================================
echo  HOAN TAT! Ban chi can gui file:
echo  D:\SeaZen APP\CanTeen-Seazen\SEAZEN_Canteen_Kiosk_Full_Package.zip
echo ===================================================================
pause
