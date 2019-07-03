$taskName = "Create Example App Firewall Rules"
$task = Get-ScheduledTask | Where-Object -Property TaskName -eq $taskName | Select-Object -ExpandProperty TaskName
$file = Test-Path "C:\Users\Public\Example App\Create_Firewall_Rules.ps1"
 
if ($task -and $file){
    Write-Host "Installed"
}
