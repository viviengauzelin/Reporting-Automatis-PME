@echo off
cd /d "%~dp0"

call venv\Scripts\activate
python generate_demo_data.py

exit