#Set seach domain (Case Sensitive)
$domain = 'LAB'
#Set application name
$appName = "Example App"
#Set the AppData path to the executable
$appDataPath = "AppData\Local\Example App\App.exe"
 
#Find all user profiles matching the search domain
$profiles = 'Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion\ProfileList\*'
$accounts = Get-ItemProperty -path $profiles
Foreach ($account in $accounts) {
    $objUser = New-Object System.Security.Principal.SecurityIdentifier($account.PSChildName)
    $objName = $objUser.Translate([System.Security.Principal.NTAccount])
    $account.PSChildName = $objName.value
}
$users = $accounts | Where-Object {$_.PSChildName -like "*$domain*"} | Select-Object -ExpandProperty PSChildName
$users = $users.Replace("$domain\","")
 
#Create a firewall exception for each user profile found
Foreach ($user in $users) {
    if (!(Get-NetFirewallRule -DisplayName "$appname $user" -ErrorAction SilentlyContinue)){
        New-NetFirewallRule -DisplayName "$appName $user" -Direction Inbound -Protocol TCP -LocalPort Any -RemoteAddress Any -Program "C:\Users\$user\$appDataPath" -Action Allow
    }
}
