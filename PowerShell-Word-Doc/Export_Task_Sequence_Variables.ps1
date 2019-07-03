#Set account credentials to connect to network share
$pass="Password"|ConvertTo-SecureString -AsPlainText -Force
$Cred = New-Object System.Management.Automation.PsCredential('user@domain.local',$pass)
 
#Set network share used for CSV storage
$netpath = '\\server.domain.local\csv'
 
#Connect to network drive
New-PSDrive -name R -Root $netpath -Credential $cred -PSProvider filesystem -ErrorAction SilentlyContinue
 
#Set variables
$sn = Get-WmiObject Win32_ComputerSystemProduct | Select-Object -ExpandProperty IdentifyingNumber
$asset = Get-WmiObject Win32_SystemEnclosure | Select-Object -ExpandProperty SmbiosAssetTag
$model = Get-WmiObject Win32_ComputerSystem | Select-Object -ExpandProperty Model
 
#Write data to the CSV and export to a file
$data = @(
  [pscustomobject]@{
    sn  = $sn
    asset  = $asset
    model = $model
  }
)
$data | Export-Csv -Path "$netpath\$sn.csv" -NoTypeInformation
 
#Remove quotes from the CSV
(Get-Content "$netpath\$sn.csv") | ForEach-Object {$_ -Replace '"', ""} | Out-File "$netpath\$sn.csv" -Force -Encoding ascii