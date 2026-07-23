$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -File `"C:\Users\micas\OneDrive\Bureau\Rapports Assistant SASU\pull_and_deploy.ps1`""

# Trigger 1 : à chaque ouverture de session
$trigger1 = New-ScheduledTaskTrigger -AtLogOn

# Trigger 2 : quotidien 18h (si PC déjà allumé depuis le matin)
$trigger2 = New-ScheduledTaskTrigger -Daily -At 18:00

$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType S4U -RunLevel Highest
Register-ScheduledTask -TaskName "ASTRA_DEPLOIEMENT_LIVRABLES" -Action $action -Trigger @($trigger1, $trigger2) -Principal $principal -Description "Pull GitHub + copie livrables ASTRA MOMENTUM (au login + 18h)" -Force
Write-Host "Tache 'ASTRA_DEPLOIEMENT_LIVRABLES' creee : declenchee a chaque ouverture de session + 18h"
