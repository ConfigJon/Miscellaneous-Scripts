Function New-RegistryValue
{
<#
    .DESCRIPTION
        Create or modify a registry value
    
    .PARAMETER RegKey
        Specify the registry key path

    .PARAMETER Name
        Specify the name of the registry value to modify

    .PARAMETER PropertyType
        Specify the type of the registry value

    .PARAMETER Value
        Specify the data the value will be set to

    .EXAMPLE
        #Create or modify a registry value
            New-RegistryValue -RegKey "HKLM:\SOFTWARE\Example" -Name "ExampleValue" -PropertyType Dword -Value 2

        #Create of modify a default registry value
            New-RegistryValue -RegKey "HKLM:\SOFTWARE\Example" -Name '(Default)' -PropertyType String -Value "abcd"

    .NOTES
        Created by: Jon Anderson (@ConfigJon)
        Modified: 1/17/2020
#>
    [CmdletBinding()]
    param(   
        [String][parameter(Mandatory=$true)][ValidateNotNullOrEmpty()]$RegKey,
        [String][parameter(Mandatory=$true)][ValidateNotNullOrEmpty()]$Name,
        [String][parameter(Mandatory=$true)][ValidateSet('String','ExpandString','Binary','DWord','MultiString','Qword','Unknown')]$PropertyType,
        [String][parameter(Mandatory=$true)][ValidateNotNullOrEmpty()]$Value
    )
        
    #Create the registry key if it does not exist
    if (!(Test-Path $RegKey))
    {
        try{New-Item -Path $RegKey -Force | Out-Null}
        catch{throw "Failed to create $RegKey"}
    }

    #Create the registry value
    try{New-ItemProperty -Path $RegKey -Name $Name -PropertyType $PropertyType -Value $Value -Force | Out-Null}
    catch{throw "Failed to set $RegKey\$Name to $Value"}

    #Check if the registry value was successfully created
    $KeyCheck = Get-ItemProperty $RegKey
    if ($KeyCheck.$Name -eq $Value){Write-Output "Successfully set $RegKey\$Name to $Value"}
    else{throw "Failed to set $RegKey\$Name to $Value"}
}