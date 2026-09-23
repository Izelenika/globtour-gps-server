@echo off
title GLOBTOUR GPS SERVER - LOCAL
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    py -3 -m venv .venv
)

".venv\Scripts\python.exe" -m pip install -q --disable-pip-version-check -r requirements.txt
start "" http://127.0.0.1:8000
".venv\Scripts\python.exe" server.py
pause
