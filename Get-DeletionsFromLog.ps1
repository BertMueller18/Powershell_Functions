<#
.Synopsis
   extract baditem lines from Exchange Moverequest-Reports as CSV Data
.DESCRIPTION
   If there are Exchange 2010/2013 Mailboxmoves started with New-Moverequest -baditemlimit xx, reports are generated as part of the moverequest.
   It is possible to extract this reports with powershell the command
   
   get-moverequest "John, Doe" | get-moverequeststatistics -includereport | select report | fl * | out-file report-John-Doe.log

   in this reports are lines with messageclass, subject, sender and so on. To extract these line this funtioin uses regular expressions to
   extract the data to screen or csv file

   TODO: Umlaut issues in regex "Entwürfe"

.EXAMPLE
   get-deletionsFromLog -logfile test1.log -outfile test2.log
   generates CSV file
.EXAMPLE
   get-deletionsFromLog -logfile test1.log
   shows data as csv on screen

#>
function get-DeletionsFromLog
{    [CmdletBinding()]
    [OutputType([int])]
    Param
    (
        # Eingabedatei
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [string]$Logfile,

        # Ausgabedatei
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        [string]$Outfile
    )
    Begin
    {
        $file = $null
        $Matches = $null
        $ergebnis = $null
        $ergebnis = @()
        $fullpath = (Split-Path -Resolve -Path $Logfile) +"\"+ (Split-Path -Leaf $Logfile)
        $ergebnis += "Ordner|sender|recipient|Subject|Messageclass|size|datesent|datereceived"
        $regex =  @"
(<baditem.*
.*
.*
.*
.*
.*
.*
.*
.*
.*</dateReceived>)
"@

        $k= $null
        $k = @()
        $re = @()
        $re += "<folder id.*>(.*)</folder>.*","<sender>(.*)</sender>.*","<recipient>(.*)</recipient>.*","<subject>(.*)</subject>.*","<messageClass>(IPM.*)</messageClass>.*","<size>(.*)</size>.*","<dateSent>(.*)<+/.*>.*","<dateReceived>(.*)<+/.*>"

    }
    Process
    { 
    Write-Verbose "Processing $Logfile" 
      $file = [System.IO.File]::ReadAllText($fullpath)
    $k = [regex]::matches($file, $regex, 'MultiLine') | % { $_.Value }
    Write-Verbose "Found $($k.count) Baditem errors in File `"$Logfile`" " 
    foreach ($item in $k) {
    $ncontent = $item -replace "`n",""
    $Matches = $null
    if ($ncontent -match $re) { 
    $ergebnis += ($Matches[1..8] -join "|")
    }
    }
   
        if ($outfile -like "*.*" ) {
                $ergebnis | ConvertFrom-Csv -Delimiter "|"| Export-Csv -Path $Outfile -Delimiter ";" -NoClobber -NoTypeInformation -Encoding utf8
                Write-Verbose "`"$Logfile`" Exported to `"$Outfile`"" 
                }
        else { $ergebnis | ConvertFrom-Csv -Delimiter "|"| ft -AutoSize
            }
    }
    End
    {
    }
}
