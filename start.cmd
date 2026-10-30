@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo Python environment not found. See README.md.
  pause
  exit /b 1
)
".venv\Scripts\python.exe" report_generator.py
if errorlevel 1 pause
