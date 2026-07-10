# Installer la tâche planifiée "AssistanteSASU" tous les mercredis à 14h30
$action = New-ScheduledTaskAction -Execute "C:\Users\micas\OneDrive\Bureau\Rapports Assistant SASU\lancer_mercredi.bat"
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Wednesday -At 14:30
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel Limited
Register-ScheduledTask -TaskName "AssistanteSASU" -Action $action -Trigger $trigger -Principal $principal -Description "Assistante virtuelle - rapport hebdo mercredi 15h" -Force
Write-Host "Tache planifiee 'AssistanteSASU' creee : mercredi 14h30"
