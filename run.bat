@echo off
cd /d "%~dp0"
if not exist .venv ( python -m venv .venv && .venv\Scripts\pip install -q -r requirements.txt )
start "" http://localhost:5055
.venv\Scripts\python app.py
