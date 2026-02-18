@echo off
cd /d "%~dp0"

call venv\Scripts\activate
python main.py

REM enlever les 3 lignes suivantes dans le cas d'une utilisation du scipt batch planifiée

echo.
echo Traitement termine.
pause
exit
