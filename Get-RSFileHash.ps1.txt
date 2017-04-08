function Get-RSFileHash {
<#
.Synopsis
   Returns an MD5 filehash when given a file path
.DESCRIPTION
   Returns an MD5 filehash when given a file path
.EXAMPLE
   Get-RSFileHash -Filename c:\temp\filetohash.txt
.EXAMPLE
   Get-ChildItem c:\temp\*.txt | get-rsfilehash
#>

    [CmdletBinding()]
    [OutputType([string])]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0,
                   ParameterSetName="ChildItem")]
        [Alias('Filename')]
        [System.IO.FileInfo]$InputObject
    )

    Process
    {
        foreach ($object in $InputObject) {
            $file = $object.fullname
            if (test-path -Path $file -PathType Leaf) {
                try { 
                    $md5 = New-Object -TypeName System.Security.Cryptography.MD5CryptoServiceProvider
                    $stream = [System.IO.File]::Open($file, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)
                    $hash = [System.BitConverter]::ToString($md5.ComputeHash($stream)) 
                    $stream.Close()
                    Write-Debug "$file - $hash"

                    Write-Output $hash.Replace('-',$null)
                }
                catch {
                    throw "Unable to generate hash for $file - $($_)"
                }
            }
        }
    }
}

$someFilePath = "C:\temp\users.csv"

get-childitem -Path "C:\temp\Validation.ps1" | Get-RSFileHash

Get-RSFileHash -Filename C:\temp\Validation.ps1