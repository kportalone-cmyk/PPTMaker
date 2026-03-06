@echo off
cd /d "%~dp0server"
call venv\Scripts\activate
uvicorn main:app --host 0.0.0.0 --port 8000 --reload
pause
