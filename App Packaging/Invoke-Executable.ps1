Function Invoke-Executable
{
<#
    .DESCRIPTION
        Install a .exe file
    
    .PARAMETER FilePath
        Specify the path to the .exe file to be installed

    .PARAMETER Arguments
        Specify additional arguments to pass to the .exe file (Comma seperated list)

	.PARAMETER ExitCodes
		Specify non-standard success exit codes (Comma seperated list)

    .EXAMPLE
        Install a .exe file
            Invoke-Executable -FilePath "$PSScriptRoot\Setup.exe" -Arguments '/S /v"/qn REBOOT=reallysuppress"'

        Install a .exe file with non-standard exit codes 2 and 8
            Invoke-Executable -FilePath "$PSScriptRoot\Setup.exe" -Arguments '/S' -ExitCodes '2,8'

    .NOTES
        Created by: Jon Anderson (@ConfigJon)
        Modified: 1/17/2020
#>
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
        if ($_ -notmatch "(\.exe)")
        {
            throw "The specified file must be a .exe file"
        }
        return $true 
    })]
    [Parameter(Mandatory = $true)][System.IO.FileInfo]$FilePath,
    [Parameter(Mandatory = $false)][ValidateNotNullOrEmpty()][String]$Arguments,
    [Parameter(Mandatory = $false)][ValidateNotNullOrEmpty()][String]$ExitCodes
    )

    #Create a list to store arguments
    $ArgumentList = New-Object 'System.Collections.Generic.List[string]'

	#Convert the exit codes to a list
    if($ExitCodes){$ExitSplit = $ExitCodes.Split(',')}

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

    #Add quotes to the FilePath
    $StringFilePath = $FilePath.ToString() #Convert the FilePath argument to a string
    $StringFilePath = $StringFilePath.insert(0,'"') #Add a quote at the start of the path
    $StringFilePath+='"' #Add a quote at the end of the path
    
    #Run the executable
    Write-Output "Running Command: $StringFilePath $ArgumentList"
    if($Arguments)
    {
        $ExitCode = (Start-Process -FilePath $StringFilePath -ArgumentList $ArgumentList -Wait -PassThru).ExitCode
    }
    else
    {
        $ExitCode = (Start-Process -FilePath $StringFilePath -Wait -PassThru).ExitCode
    }

    #Report the Exit Code
    Write-Output "The exit code is $ExitCode"

    #Terminate the script if an error occurs
    if(($ExitCode -ne 0) -and ($ExitCode -ne 3010) -and !($ExitSplit -contains $ExitCode)){throw "The executable terminated with error code $ExitCode"}
}