@echo off
title GLOBTOUR GPS SERVER
cd /d "%~dp0"

echo ==========================================
echo       GLOBTOUR GPS SERVER
echo ==========================================
echo.

where py >nul 2>nul
if errorlevel 1 (
    echo GRESKA: Python nije pronadjen.
    echo Instaliraj Python 3.11 ili noviji i ponovo pokreni.
    echo.
    pause
    exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
    echo [1/3] Kreiram Python okruzenje...
    py -3 -m venv .venv
    if errorlevel 1 (
        echo GRESKA pri kreiranju virtualnog okruzenja.
        pause
        exit /b 1
    )
)

echo [2/3] Provjeravam potrebne pakete...
".venv\Scripts\python.exe" -m pip install -q --disable-pip-version-check -r requirements.txt
if errorlevel 1 (
    echo.
    echo GRESKA pri instalaciji paketa.
    pause
    exit /b 1
)

echo [3/3] Pokrecem Globtour GPS Server...
echo.
echo Web stranica: http://127.0.0.1:8000
echo GPS TCP port: 9000
echo.
echo Ovaj prozor mora ostati otvoren dok server radi.
echo Za zaustavljanje pritisni CTRL+C.
echo.

start "" http://127.0.0.1:8000
".venv\Scripts\python.exe" server.py

echo.
echo Server je zaustavljen.
pause
