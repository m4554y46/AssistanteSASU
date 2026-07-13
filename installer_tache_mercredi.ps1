# Installer la tâche planifiée "ASTRA_MOMENTUM_RapportHebdo" tous les mercredis à 14h30
$action = New-ScheduledTaskAction -Execute "C:\Users\micas\OneDrive\Bureau\Rapports Assistant SASU\lancer_mercredi.bat"
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Wednesday -At 14:30
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel Limited
Register-ScheduledTask -TaskName "ASTRA_MOMENTUM_RapportHebdo" -Action $action -Trigger $trigger -Principal $principal -Description "Rapport hebdomadaire ASTRA MOMENTUM - Veille + Agents + Email" -Force
Write-Host "Tache planifiee 'ASTRA_MOMENTUM_RapportHebdo' creee : mercredi 14h30"
