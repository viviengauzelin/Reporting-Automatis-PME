@echo off
cd /d "%~dp0"

call venv\Scripts\activate
start "" streamlit run app.py --server.address localhost
exit
