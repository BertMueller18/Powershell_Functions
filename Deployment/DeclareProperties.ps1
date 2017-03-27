param($deploymentshare =".", $inifile="", $ztigatherfile="$deploymentshare\scripts\ZTIGather.xml", $outputFile=".\Variables.xml")

#' *    Aly Shivji 1-20-2012
#      =======================================================================================
#    '*
#    '* Disclaimer
#    '*
#    '* This script is not supported under any Microsoft standard support program or service. This 
#    '* script is provided AS IS without warranty of any kind. Microsoft further disclaims all 
#    '* implied warranties including, without limitation, any implied warranties of merchantability 
#    '* or of fitness for a particular purpose. The entire risk arising out of the use or performance 
#    '* of this script remains with you. In no event shall Microsoft, its authors, or anyone else 
#    '* involved in the creation, production, or delivery of this script be liable for any damages 
#    '* whatsoever (including, without limitation, damages for loss of business profits, business 
#    '* interruption, loss of business information, or other pecuniary loss) arising out of the use 
#    '* of or inability to use this script, even if Microsoft has been advised of the possibility 
#    '* of such damages.
#    '*
#    '*=======================================================================================

$globalvars = @()
$globallists = @()

if ($outputFile -ne $null)
    {
    Set-Content -Path $outputfile -Value "<?xml version=""1.0"" encoding=""utf-8""?>"    
    Add-Content -Path $outputfile -Value "<?xml-stylesheet type=""text/xsl"" href=""varDocumentor.xsl""?>"
    Add-Content -Path $outputfile -Value "<ListofVariables>"
    }

# Load the ZTIGather.xml file if found or specified to find properties already declared
if (test-path($ztigatherfile))
    {
    write-host "===================Found a ZTIGather.xml file - parsing already declared variables====================="
    $xml = [xml] (Get-Content $ztigatherfile)
    $alreadydeclared = $xml.properties.property | Where-Object { $_.type -notmatch "list" } | foreach {$_.getAttribute("id") }
    if ($alreadydeclared -ne $null) {$alreadydeclared = $alreadydeclared | foreach { $_.trim().toUpper() } | Sort-Object | Get-Unique}
    foreach ($var in  $alreadydeclared)
        {
        Add-Content -Path $outputFile -Value "<Variable Type=""Variable"" Name=""$var"" Reference=""$ztigatherfile""/>"
        }
    $lists = $xml.properties.property | Where-Object { $_.type -match "list" } | foreach {$_.getAttribute("id") }
    if ($lists-ne $null) {$lists = $lists | foreach { $_.trim().toUpper() } | Sort-Object | Get-Unique}
    foreach ($list in  $lists)
        {
        Add-Content -Path $outputFile -Value "<Variable Type=""List"" Name=""$var"" Reference=""$ztigatherfile""/>"
        }
    $alreadydeclared = $alreadydeclared + $lists
    }

# Load the ini file if specified to find properties already declared

if (($inifile -ne "") )
    {
    if (test-path($inifile))
        {
        $ini = Get-Content $inifile | select-string "^Properties=(.*)"  | Select -Expand Matches | Foreach {$_.Groups[1]} | Select -Expand Value | foreach { $_.trim().toUpper() } | Sort-Object | Get-Unique
        if (($ini -ne $null) -and ($ini -ne "")) 
            {
            write-host "===================Found an INI file - parsing already declared variables=========================="
            foreach ($var in  $ini.Split(","))
                { 
                Add-Content -Path $outputFile -Value "<Variable Type=""Variable"" Name=""$var"" Reference=""$inifile"" />"                
                }
            
            $alreadydeclared = $alreadydeclared + $ini.Split(",")
            }
        }
    }

#trim the already declared list
if ($alreadydeclared -ne $null) {$alreadydeclared = $alreadydeclared | foreach { $_.trim().toUpper() }}

$found = $alreadydeclared.length
if ($found -eq "") {$found=0}

write-host "===================Found $found already declared variables=========================="

if (test-path("$deploymentshare\Control"))
{
    #find all task sequences
    $tasksequences = Get-ChildItem $deploymentshare\Control -recurse | where{$_.Name -match "^ts\.xml$"} 

    foreach ($ts in $tasksequences)
        {
        $vars = (Get-Content $ts.Fullname | select-string "<variable name=\""Variable\"">(\w*)</variable>|%(\w*)%" | Select -Expand Matches | Foreach {$_.Groups[1..2]} | Select -Expand Value | foreach { $_.trim().toUpper() } | Sort-Object | Get-Unique )
        foreach ($var in  $vars)
            {
            if ($var -ne "")
            {
                $tsname = $ts.fullName 
                Add-Content -Path $outputFile -Value "<Variable Type=""Variable"" Name=""$var"" Reference=""$tsname"" />"
            }                
            }
    }
    $globalvars = $vars
}


# Find all scripts    
$scripts = Get-ChildItem $deploymentshare -recurse |where{$_.extension -match "wsf|vbs|ps1|bat|cmd"}
$numscripts = $scripts.length
write-host "===================Found $numscripts scripts======================================="
foreach ($script in $scripts)
    {
    
    #Seach for 1 VBScript (oEnvironment.Item("variable")) or 2 Batch (%variable%) or 3 Powershell (tsenv:variable) references to variables, match any of those 3 scenarios
    $vars = (Get-Content $script.Fullname | select-string "oEnvironment.Item\(\""(\w*)|%(\w*)%|tsenv:(\w*)" | Select -Expand Matches | Foreach {$_.Groups[1..3]} | Select -Expand Value | foreach { $_.trim().toUpper() } | Sort-Object | Get-Unique  )    
    $globalvars = $globalvars + $vars

    foreach ($var in  $vars)
        {
            if ($var -ne "")
                {
                $scriptname = $script.fullName 
                Add-Content -Path $outputFile -Value "<Variable Type=""Variable"" Name=""$var"" Reference=""$scriptname"" />"
                }                
        }


    #search for 1 Vbscript (oEnvironment.ListItem("variable")) or 2 Powershell (tsenvlist:variable) references to variables, match any of those 2 scenarios
    $lists = (Get-Content $script.Fullname | select-string "oEnvironment.ListItem\(\""(\w*)|tsenvlist:(\w*)" | Select -Expand Matches | Foreach {$_.Groups[1..2]} | Select -Expand Value | foreach { $_.trim().toUpper() } | Sort-Object | Get-Unique)    
    $globallists = $globallists + $lists

    foreach ($var in  $lists)
        {
            if ($var -ne "")
                {
                $scriptname = $script.fullName 
                Add-Content -Path $outputFile -Value "<Variable Type=""List"" Name=""$var"" Reference=""$scriptname"" />"
                }                
        }


    }

Add-Content -Path $outputfile -Value "</ListofVariables>"


# Remove any null or empty values
$globalvars = $globalvars | foreach { if (($_ -ne $null) -and ($_ -ne "")){$_.toupper()} }
$globallists = $globallists | foreach { if (($_ -ne $null) -and ($_ -ne "")) {$_.toupper()} }

# Find unique undeclared variable lists
$uniquelists = $globallists |  ? {$alreadydeclared -notcontains $_} | Sort-Object | Get-Unique 

# Find unique undeclared variables without the reserved _SMSTS prefix
$uniquevars = $globalvars  | ? {$alreadydeclared -notcontains $_} | ?{ $_ -notmatch "^_SMSTS"} | ?{ $_ -notmatch "[0-9]{3}$" } | Sort-Object | Get-Unique

$line = ($uniquevars -join ",")

# Add any lists with the (*) declaration
if ($uniquelists.length -gt 0) 
    { 
    $line = $line + "," + ($uniquelists -join "(*),") 
    $line = $line + "(*)"
    }

$newvars = $uniquelists.length + $uniquevars.length
write-host "===================Found $newvars new variables=========================="
write-host $line
