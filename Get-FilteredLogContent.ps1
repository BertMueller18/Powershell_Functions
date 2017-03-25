<#PSScriptInfo

.VERSION 1.0.0

.GUID cc7f100d-5fee-4e5d-8963-0b461fbcfda1

.AUTHOR Mike Hendrickson

#>

<# 
.SYNOPSIS
 Used to filter out log entries from one or more file paths on one or more target computers.

.DESCRIPTION 
 Used to filter out log entries from one or more file paths on one or more target computers.
 
.PARAMETER Path
 The path(s) to look for log files in. Paths are searched recursively.

.PARAMETER Include
 Optional Include strings to pass to Get-ChildItem when looking for files in specifieds Path(s).

.PARAMETER Exclude
 Optional Exclude strings to pass to Get-ChildItem when looking for files in specifieds Path(s).

.PARAMETER SearchStrings
 One or more non-empty strings to use when searching for content in log files. 

.PARAMETER ComputerName
 Optional list of one or more target computers to send remote log filtering jobs to jobs to. If nothing is specified, the log filtering will be done locally.

.PARAMETER StartTime
 The minimum Creation or Last Write date and time for logs to inspect. Defaults to the start of time.

.PARAMETER EndTime
 The maximum Creation or Last Write date and time for logs to inspect. Defaults to the end of time.

.PARAMETER AppendComputerNameToMatches
 Whether the computer name where the log match was found should be appended to each match

.PARAMETER ComputerNameSeparator
 If AppendComputerNameToMatches is $true, specifies the character to use to separate the match from the computer name.

.PARAMETER SimpleMatch
 Whether SearchStrings should be treated literally (the default), or treated as a Regular Expression.

.EXAMPLE
 >Inspects the IIS log directory on an Exchange Server and looks for all entries containing user DOMAIN\username (backslash must be escaped) against the EWS virtual directory
 PS> .\Get-FilteredLogContent.ps1 -Path "C:\inetpub\logs\LogFiles\W3SVC1" -SearchStrings "DOMAIN\\username","/EWS/Exchange.asmx" -Verbose

.EXAMPLE
 >Inspects the EWS log directory on an Exchange Server and looks for all entries containing user user1@contoso.com
 PS> .\Get-FilteredLogContent.ps1 -Path "C:\Program Files\Microsoft\Exchange Server\V15\Logging\Ews" -SearchStrings "user1@contoso.com" -ComputerNameSeparator "," -Verbose

.INPUTS
None. You cannot pipe objects to Get-FilteredLogContent.ps1.

.OUTPUTS
[System.Collections.Generic.List[System.Object]]. Get-FilteredLogContent.ps1 returns a List of Object's containing the discovered log matches
#>

[CmdletBinding()]
[OutputType([System.Collections.Generic.List[System.Object]])]
param
(
    [parameter(Mandatory = $true)] 
    [String[]]
    $Path,

    [String[]]
    $Include = @(),

    [String[]]
    $Exclude = @(),

    [parameter(Mandatory = $true)]
    [String[]]
    $SearchStrings = @(),

    [String[]]
    $ComputerName = @(),

    [DateTime]
    $StartTime = [DateTime]::MinValue,
    
    [DateTime]
    $EndTime = [DateTime]::MaxValue,

    [Bool]
    $AppendComputerNameToMatches = $true,

    [String]
    $ComputerNameSeparator = " ",

    [Boolean]
    $SimpleMatch = $true
)

function Get-LogContentLocal
{
    [CmdletBinding()]
    [OutputType([System.Collections.Generic.List[System.Object]])]
    param
    ( 
        [String[]]
        $Path,

        [String[]]
        $Include,

        [String[]]
        $Exclude,

        [String[]]
        $SearchStrings,

        [DateTime]
        $StartTime,
    
        [DateTime]
        $EndTime,

        [Bool]
        $AppendComputerNameToMatches,

        [String]
        $ComputerNameSeparator,

        [Boolean]
        $SimpleMatch = $true
    )

    #Return value
    [System.Collections.Generic.List[System.Object]]$allMatches = New-Object System.Collections.Generic.List[System.Object]

    #Find all logs in the IIS log directory
    $logs = Get-ChildItem -Path $Path -Include $Include -Exclude $Exclude -Recurse | Where-Object {($_.LastWriteTime -gt $StartTime -and $_.LastWriteTime -lt $EndTime) -or ($_.CreationTime -gt $StartTime -and $_.CreationTime -lt $EndTime)}

    Write-Verbose "$([DateTime]::Now): Found $($logs.Count) logs to search."

    foreach ($log in $logs)
    {
        $matchCount = 0

        Write-Verbose "$([DateTime]::Now): Searching log $($log.FullName)."

        $logFiltered = Select-String -Path $log.FullName -Pattern $SearchStrings[0] -SimpleMatch:$SimpleMatch

        for ($i = 1; $i -lt $SearchStrings.Count; $i++)
        {
            $logFiltered = $logFiltered | Select-String -Pattern $SearchStrings[$i] -SimpleMatch:$SimpleMatch
        }

        foreach ($match in $logFiltered)
        {
            $matchCount++

            if ($AppendComputerNameToMatches)
            {
                $allMatches.Add($match.Line + $ComputerNameSeparator + $env:COMPUTERNAME)
            }
            else
            {
                $allMatches.Add($match.Line)
            }
        }

        Write-Verbose "$([DateTime]::Now): Found $($matchCount) matches in log $($log.FullName)."
    }

    Write-Verbose "$([DateTime]::Now): Found $($allMatches.Count) matches in all logs."

    return $allMatches
}

### Begin script execution

$SearchStrings = $SearchStrings | Where-Object {$_.Trim().Length -gt 0}

if ($SearchStrings.Count -eq 0)
{
    Write-Error "Either no Search Strings, or only empty Search Strings, were specified. Exiting script."
    return
}

if ($ComputerName.Count -eq 0)
{
    $logMatches = Get-LogContentLocal -Path $Path -Include $Include -Exclude $Exclude -SearchStrings $SearchStrings -StartTime $StartTime -EndTime $EndTime -AppendComputerNameToMatches $AppendComputerNameToMatches -ComputerNameSeparator $ComputerNameSeparator -SimpleMatch $SimpleMatch
}
else
{
    $jobs = Invoke-Command -ComputerName $ComputerName -ScriptBlock ${function:Get-LogContentLocal} -ArgumentList $Path, $Include, $Exclude, $SearchStrings, $StartTime, $EndTime, $AppendComputerNameToMatches, $ComputerNameSeparator, $SimpleMatch -AsJob

    Wait-Job $jobs | Out-Null

    $logMatches = Receive-Job $jobs
}

$logMatches = $logMatches | Sort-Object

return $logMatches