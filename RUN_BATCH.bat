@echo off
cd /d "%~dp0"

call venv\Scripts\activate
python main.py

REM Vérification du code de sortie :
REM   0 = succès
REM   2 = alerte d'intégrité des données (écart détecté par la réconciliation)
REM
REM Enlever les lignes ci-dessous (jusqu'à "exit") en cas d'utilisation planifiée.
REM Le planificateur de tâches Windows détectera lui-même le code de sortie 2.

IF ERRORLEVEL 2 (
    echo.
    echo ============================================================
    echo  ALERTE : ecart de valeur des données detecte. Consultez le log avant
    echo  de diffuser le reporting. Code de sortie : 2
    echo ============================================================
)

echo.
echo Traitement termine.
pause
exit