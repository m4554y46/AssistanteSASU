@echo off
setlocal enabledelayedexpansion

:: ============================================================
:: LANCEUR MERCREDI 15h — Assistante Virtuelle
:: Enchaîne : délai aléatoire → rapport hebdo → agents → email
:: À programmer dans le Task Scheduler vers 14h30
:: ============================================================

cd /d "%~dp0"

:: 1. DÉLAI ALÉATOIRE (900-2700s = 15-45 min)
set /a delay=900 + !random! %% 1801
echo [%time:~0,8%] Attente de %delay% secondes avant le rapport...
if %delay% gtr 0 (
    timeout /t %delay% /nobreak >nul
)

echo [%time:~0,8%] === DEBUT DU CYCLE HEBDOMADAIRE ===

:: 2. RAPPORT HEBDOMADAIRE (veille + prospection)
echo [%time:~0,8%] Lancement du rapport hebdo...
cd /d "%~dp0moteur_rapport"
python run_weekly_report.py
if errorlevel 1 echo [ERR] Rapport hebdo echoue

:: 3. AGENTS SPECIALISES
echo [%time:~0,8%] Lancement des agents...
cd /d "%~dp0agents"
python run_agent.py all
if errorlevel 1 echo [ERR] Agents echoues

:: 4. COPIE VERS DOSSIER PROPRE (sans le code IA)
echo [%time:~0,8%] Copie des livrables vers le dossier clean...
set LIVRABLES="C:\Users\micas\OneDrive\Bureau\ASTRA MOMENTUM - Livrables"
copy /Y "%~dp0Rapport_Astra_Momentum_*.docx" %LIVRABLES% >nul 2>&1
copy /Y "%~dp0agents\output\*.docx" %LIVRABLES% >nul 2>&1
echo [%time:~0,8%] Livrables copies dans %LIVRABLES%

:: 5. OUVERTURE DU DOSSIER PROPRE
explorer %LIVRABLES%

echo [%time:~0,8%] === CYCLE TERMINE ===
