$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -File `"C:\Users\micas\OneDrive\Bureau\Rapports Assistant SASU\pull_and_deploy.ps1`""

# Trigger 1 : à chaque ouverture de session
$trigger1 = New-ScheduledTaskTrigger -AtLogOn

# Trigger 2 : au démarrage du PC (même sans ouverture de session)
$trigger2 = New-ScheduledTaskTrigger -AtStartup

# Trigger 3 : quotidien 18h (si PC déjà allumé depuis le matin)
$trigger3 = New-ScheduledTaskTrigger -Daily -At 18:00

# Exécuter dès que possible si un déclenchement a été manqué (PC éteint/BSOD)
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit ([TimeSpan]::Zero)

$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType S4U -RunLevel Highest
Register-ScheduledTask -TaskName "ASTRA_DEPLOIEMENT_LIVRABLES" -Action $action -Trigger @($trigger1, $trigger2, $trigger3) -Settings $settings -Principal $principal -Description "Pull GitHub + copie livrables ASTRA MOMENTUM (au login + demarrage + 18h)" -Force
Write-Host "Tache 'ASTRA_DEPLOIEMENT_LIVRABLES' creee : declenchee au login, au demarrage et a 18h (rattrapage si manque)"
