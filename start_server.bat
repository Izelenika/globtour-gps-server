@echo off
cd /d "%~dp0"
if not exist .venv\Scripts\python.exe (
  echo Virtual environment nije napravljen.
  echo Pokreni:
  echo py -3 -m venv .venv
  echo .venv\Scripts\activate
  echo pip install -r requirements.txt
  pause
  exit /b 1
)
.venv\Scripts\python.exe server.py
pause
