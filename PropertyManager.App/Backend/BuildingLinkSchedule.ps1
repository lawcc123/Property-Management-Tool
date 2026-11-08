param([ValidateSet('Get','Save','Disable')][string]$Action='Get')
$ErrorActionPreference='Stop'
$inputData=[Console]::In.ReadToEnd() | ConvertFrom-Json
$identity=[Security.Principal.WindowsIdentity]::GetCurrent()
$taskName='PropertyManagerDemo - BuildingLink Data Download - '+$identity.User.Value
$marker='PropertyManagerDemo.BuildingLink.v1'
$task=Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
if($task -and $task.Description -ne $marker) { throw 'A different task uses this name. It has not been changed.' }
if($Action -eq 'Save') {
    $clock=[datetime]::ParseExact($inputData.Time,'HH:mm',[Globalization.CultureInfo]::InvariantCulture)
    if(-not (Test-Path -LiteralPath $inputData.Executable -PathType Leaf)) { throw 'The app executable was not found.' }
    $days=@($inputData.Days | ForEach-Object { if($_ -lt 0 -or $_ -gt 6) { throw 'Invalid schedule day.' }; [DayOfWeek][int]$_ })
    if($days.Count -eq 0) { throw 'Choose at least one day.' }
    $start=[datetime]::Today.Add($clock.TimeOfDay)
    $trigger=if($days.Count -eq 7) { New-ScheduledTaskTrigger -Daily -At $start } else { New-ScheduledTaskTrigger -Weekly -DaysOfWeek $days -At $start }
    $actionDef=New-ScheduledTaskAction -Execute $inputData.Executable -Argument '--buildinglink-scheduled' -WorkingDirectory ([IO.Path]::GetDirectoryName($inputData.Executable))
    $principal=New-ScheduledTaskPrincipal -UserId $identity.Name -LogonType Interactive -RunLevel Limited
    $settings=New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Hours 1)
    Register-ScheduledTask -TaskName $taskName -Description $marker -Action $actionDef -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null
} elseif($Action -eq 'Disable' -and $task) {
    Disable-ScheduledTask -TaskName $taskName | Out-Null
}
$task=Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
$other=@(Get-ScheduledTask | Where-Object { $_.TaskName -ne $taskName -and $_.State -ne 'Disabled' -and ($_.TaskName -match 'BuildingLink.*(Backup|Download)|Resident Email.*Download|(Occupant|Unit).*Download') } | ForEach-Object { $_.TaskName })
if($task) {
    $info=Get-ScheduledTaskInfo -TaskName $taskName
    @{Exists=$true;Enabled=($task.State -ne 'Disabled');State=[string]$task.State;NextRun=$info.NextRunTime.ToString('g');LastResult=$info.LastTaskResult;OtherTasks=$other;TriggersXml=(Export-ScheduledTask -TaskName $taskName)} | ConvertTo-Json -Compress
} else { @{Exists=$false;Enabled=$false;OtherTasks=$other} | ConvertTo-Json -Compress }
