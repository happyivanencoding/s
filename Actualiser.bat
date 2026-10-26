@echo off
chcp 65001 >nul
set PYTHONUTF8=1
cd /d "%~dp0"
python actualiser_excel.py
if errorlevel 1 (
    echo La mise a jour a echoue. Voir le message ci-dessus.
    pause
    exit /b 1
)
start "" "Score_Sectoriel_EU_NEW.xlsm"
start "" "Score_Sectoriel_US_NEW.xlsm"
