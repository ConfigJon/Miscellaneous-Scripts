<#
    .DESCRIPTION
        Dynamically search the registry for software and uninstall it. Only works for software that has a GUID
    
    .PARAMETER DisplayName
        The name of the software to be uninstalled

    .EXAMPLE
        Uninstall 7-Zip
            .\Uninstall-Application.ps1 -DisplayName "7-Zip"

    .NOTES
        Created by: Jon Anderson (@ConfigJon)
        Modified: 10/20/2020
#>

#Parameters ===================================================================================================================

param(
    [parameter(Mandatory = $false)][ValidateNotNullOrEmpty()][String]$DisplayName
)

#Functions ====================================================================================================================

Function Find-App
{
    param(
        [parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][String]$UninstallKey
    )

    #Get all apps under the specified key
    $Apps = Get-ChildItem -Path $UninstallKey | Select-Object -ExpandProperty Name
    
    #Find any matching apps
    ForEach($App in $Apps){
        $TempPath = $App -replace "HKEY_LOCAL_MACHINE","HKLM:"
        $TempApp = Get-ItemProperty -Path $TempPath
        if(($TempApp).DisplayName -match $DisplayName)
        {
            #If the matching app has a GUID, run MsiExec.exe to uninstall it
            if($TempApp.PSChildName -Match ("^(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}$"))
            {
                Invoke-MsiExec -Uninstall -Guid $TempApp.PSChildName -Arguments '/qn,/norestart'   
            }
            #if the matching app does not have a GUID, output the contents of the UninstallString value
            else
            {
                throw "No GUID found. UninstallString is: $($TempApp.UninstallString)"
            }
        }
    }
}

Function Uninstall-App
{
    param(
        [parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][String]$DisplayName
    )

    #Search the standard uninstall key
    Find-App -UninstallKey "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    
    #Search the WOW64 uninstall key
    if(Test-Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall")
    {
        Find-App -UninstallKey "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    }
}

Function Invoke-MsiExec
{
    param(
    [ValidateScript({
        if (!($_ | Test-Path))
        {
            throw "The specified file does not exist"
        }
        if (!($_ | Test-Path -PathType Leaf))
        {
            throw "The FilePath argument must be a file. Folder paths are not allowed."
        }
        if (($_ -notmatch "(\.msi)") -and ($_ -notmatch "(\.msp)"))
        {
            throw "The specified file must be a .msi or .msp file"
        }
        return $true 
    })]
    [System.IO.FileInfo]$FilePath,
    [Parameter(Mandatory = $false)][ValidateNotNullOrEmpty()][String]$Arguments,
    [Parameter(Mandatory = $false)][ValidateNotNullOrEmpty()][String]$ExitCodes,
    [Parameter(Mandatory = $false)][ValidateNotNullOrEmpty()][Switch]$Install,
    [Parameter(Mandatory = $false)][ValidateNotNullOrEmpty()][Switch]$Uninstall,
    [Parameter(Mandatory = $false)][ValidateNotNullOrEmpty()][Switch]$Patch,
    [Parameter(Mandatory = $false)][ValidateNotNullOrEmpty()][String]$Guid
    )

    #Parameter validation
    if(!($Install) -and !($Uninstall) -and !($Patch)){throw "One of the Install or Uninstall or Patch parameters must be specified"}
    if($Install -and ($Uninstall -or $Patch)){throw "Only one of the Install or Uninstall or Patch parameters can be specified"}
    if($Uninstall -and ($Install -or $Patch)){throw "Only one of the Install or Uninstall or Patch parameters can be specified"}
    if($Patch -and ($Install -or $Uninstall)){throw "Only one of the Install or Uninstall or Patch parameters can be specified"}
    if($Install -and !($FilePath)){throw "The FilePath parameter must be specified when using the Install parameter"}
    if($Patch -and !($FilePath)){throw "The FilePath parameter must be specified when using the Patch parameter"}
    if($Install -and $Guid){throw "The Guid parameter should not be specified when using the Install parameter"}
    if($Patch -and $Guid){throw "The Guid parameter should not be specified when using the Patch parameter"}

    #Create a list to store arguments
    $ArgumentList = New-Object 'System.Collections.Generic.List[string]'
	
    #Convert the exit codes to a list
    if($ExitCodes){$ExitSplit = $ExitCodes.Split(',')}

    #Add the install, uninstall, or patch argument to the list
    if($Install){$ArgumentList.Add('/i')}
    if($Uninstall){$ArgumentList.Add('/x')}
    if($Patch){$ArgumentList.Add('/p')}

    #Add the FilePath argument to the list
    if($FilePath)
    {
        $StringFilePath = $FilePath.ToString() #Convert the FilePath argument to a string
        $StringFilePath = $StringFilePath.insert(0,'"') #Add a quote at the start of the path
        $StringFilePath+='"' #Add a quote at the end of the path
        $ArgumentList.Add($StringFilePath)
    }

    #Add the Guid argument to the list
    if($Guid){$ArgumentList.Add($Guid)}

    #Add any additional arguments to the list
    if($Arguments)
    {
        $Arguments = $ExecutionContext.InvokeCommand.ExpandString($Arguments) #Expand any variables passed in the arguments list
        $ArgumentsSplit = $Arguments.Split(',')
        $Count = 0

        while($Count -lt $ArgumentsSplit.Count)
        {
            $ArgumentList.Add($ArgumentsSplit[$Count].Trim())
            $Count++
        }
    }
    
    #Run MsiExec
    if($FilePath){Write-Output "Running Command: MsiExec.exe $ArgumentList"}
    if($Guid){Write-Output "Uninstalling $Guid"}
    $ExitCode = (Start-Process -FilePath "MsiExec.exe" -ArgumentList $ArgumentList -Wait -PassThru).ExitCode

    #Report the Exit Code
    Write-Output "The exit code is $ExitCode"

    #Terminate the script if an error occurs
    if(($ExitCode -ne 0) -and ($ExitCode -ne 3010) -and !($ExitSplit -contains $ExitCode)){throw "MsiExec terminated with error code $ExitCode"}
}

#Main program =================================================================================================================

Uninstall-App -DisplayName $DisplayName