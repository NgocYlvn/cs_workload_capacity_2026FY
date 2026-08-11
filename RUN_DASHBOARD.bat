@echo off
cd /d "%~dp0"
streamlit run app.py
if errorlevel 1 (
  echo.
  echo Dashboard could not start. Please install dependencies first:
  echo pip install -r requirements.txt
  pause
)
