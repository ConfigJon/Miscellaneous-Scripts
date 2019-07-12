#Set variables =========================================================
$appName = "Example App"
$scriptName = "Create_Firewall_Rules.ps1"
$folderPath = "$Env:public\"
#=======================================================================
 
#Set paths
$scriptFolder = "$folderPath$appname"
$scriptFile = "$scriptFolder\$scriptName"
 
#Create the public app directory if it does not already exist
if (!(Test-Path -PathType Container "$scriptFolder")){
    New-Item -ItemType Directory -Path "$scriptFolder"
}
   
#Copy the firewall rules script to the public app directory if it does not already exist
if (!(Test-Path "$scriptFile")){
    Copy-Item "$PSScriptRoot\$scriptName" -Destination "$scriptFolder"
}
   
#Create a scheduled task to run the firewall rules script at user logon
$taskName = "Create $appName Firewall Rules"
$task = Get-ScheduledTask | Where-Object -Property TaskName -eq $taskName | Select-Object -ExpandProperty TaskName -ErrorAction SilentlyContinue
if (!($task)){
    $action = New-ScheduledTaskAction -Execute 'Powershell.exe' `
    -Argument "-ExecutionPolicy Bypass -File $scriptFile"
    $trigger =  New-ScheduledTaskTrigger -AtLogon
    $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit 00:05:00
    Register-ScheduledTask -Action $action -Trigger $trigger -TaskName "$taskName" -User "System" -Settings $settings -Description "Creates firewall rules to allow $appName for all domain users."
}
