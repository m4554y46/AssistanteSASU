@echo off
setlocal enabledelayedexpansion

:: ============================================================
:: LANCEUR AVEC ALEATOIRE (Entre 0 et 3600 secondes = 1 heure)
:: Le rapport sera lancé entre 14h30 et 15h30 si la tâche est
:: déclenchée à 14h30.
:: ============================================================

:: Génère un nombre aléatoire entre 0 et 3600
set /a delay=%random% %% 3601

echo [%time:~0,8%] Lancement du script. Attente de %delay% secondes avant de generer le rapport...

:: Attend le nombre de secondes calculé (ne bloque pas si delay=0)
if %delay% gtr 0 (
    timeout /t %delay% /nobreak >nul
)

:: Une fois l'attente terminée, on lance le vrai rapport
echo [%time:~0,8%] Fin de l'attente. Lancement du rapport...
cd /d "%~dp0"
python run_weekly_report.py

:: Petite pause pour voir le résultat dans la console (si exécuté en manuel)
pause