# Appel quotidien à 18h (PC allumé ou non au moment du trigger)
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -File `"C:\Users\micas\OneDrive\Bureau\Rapports Assistant SASU\pull_and_deploy.ps1`""
$trigger = New-ScheduledTaskTrigger -Daily -At 18:00
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType S4U -RunLevel Highest
Register-ScheduledTask -TaskName "ASTRA_DEPLOIEMENT_LIVRABLES" -Action $action -Trigger $trigger -Principal $principal -Description "Pull GitHub + copie livrables ASTRA MOMENTUM" -Force
Write-Host "Tache 'ASTRA_DEPLOIEMENT_LIVRABLES' creee : tous les jours a 18h00"
