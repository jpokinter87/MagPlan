@echo off
REM Script pour générer facilement un planning mensuel
REM Usage: generer_planning.bat <mois> <annee>

if "%1"=="" (
    echo Usage: generer_planning.bat ^<mois^> ^<annee^>
    echo Exemple: generer_planning.bat 2 2025
    echo.
    echo Ce script genere un planning pour le mois et l'annee specifies.
    exit /b 1
)

if "%2"=="" (
    echo Usage: generer_planning.bat ^<mois^> ^<annee^>
    echo Exemple: generer_planning.bat 2 2025
    echo.
    echo Ce script genere un planning pour le mois et l'annee specifies.
    exit /b 1
)

set MOIS=%1
set ANNEE=%2

REM Nom de fichier avec zero padding
if %MOIS% LSS 10 (
    set MOIS_PADDED=0%MOIS%
) else (
    set MOIS_PADDED=%MOIS%
)

set OUTPUT=planning_%MOIS_PADDED%_%ANNEE%.xlsx

echo Generation du planning pour %MOIS%/%ANNEE%...
python generer_planning_mensuel.py %MOIS% %ANNEE% tableau_répartition_audiences.xlsx %OUTPUT%

if %ERRORLEVEL% EQU 0 (
    echo.
    echo Planning genere avec succes: %OUTPUT%
) else (
    echo.
    echo Erreur lors de la generation du planning.
)
