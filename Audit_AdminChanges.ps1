     <#   
    .SYNOPSIS   
        Retrieves Admin Changes made within the date range of 7 days and builds HTML page from this data.
         
    .DESCRIPTION   
        Launches Exchange Management Shell from Powershell and search all admin changes made in the ladt 7 days. Groups these changes and counts to display a total number of times each cmdlet was ran. 
            Builds out a HTML page for each cmdlet that was found and displays admin names, date ran, the object modified, and the parameters used.
        
         
    .NOTES   
        Author: Josh Lavely
        Version: 1.0
     
    #> 
$HTMLPath=Read-Host -Prompt "Please Enter Path of where you want this saved (Example: '\\Server\Share\Folder\')"
$ExServer = Read-Host -Prompt "Please enter FQDN of an Exchange Server"

## Load Exchange Shell into Powershell
$ExPath="$env:ExchangeInstallPath" + "bin\RemoteExchange.ps1"
. $ExPath ; Connect-ExchangeServer -Server $ExServer -clientapplication:ManagementShell 
$GenerateXMLtoFolder = "$htmlPath"

## Clear Variables ##
$logs=$null
$LogHTMLDoc=$null
$path=$Null
$target=$null
$HTMLContent=$null
$ServerTableEntrys = $null
$ServerTableE = $null
$mastertable=$null

## Get Run Dates ##
$startDate=(Get-Date).AddDays(-7)
$Enddate=Get-Date

## Gather Log data ##
$LOGS=Search-AdminAuditLog -ResultSize 100000 -StartDate $startDate -EndDate $Enddate 
$data=$logs | Group-Object cmdletname | select count,name | sort count -Descending

## Count CMDLet's Ran and Build Main HTML Data ##
ForEach ($d in ($data | select name,count)){
                      ## Remove old data from directory

                      Remove-Item "$HTMLPath$($d.name).html" -Force -ErrorAction Ignore

                      ## Build Main HTML Page Information

                      $target = $($d.name) + '.html'
                      $ServerTableEntrys += $("<tr><td><a href=" + $target + " target='_Top'>$($d.name)</td></a> <td>" +  $d.count + "</td></tr>")

                      }

## Build page and fill data for each log ##
ForEach ($Log in $LOGS){
$table=$null
      ForEach ($lo in $log){
                              ## Build Log Data
                              $Table=[pscustomobject][ordered]@{
                                                "Command Ran" = $log.CmdletName
                                                "Date Ran" = $log.RunDate
                                                "Admin Name" = $log.Caller
                                                "Object Modified" = $log.ObjectModified
                                                "Parameters Used" = [string]$Log.CmdletParameters -replace " ",","
                                                "Succeeded" = $log.Succeeded
            
                                                                }
                           }
## Add Built Data to be displayed on the HTML page ##
$log2 = $table | select -Property * | ConvertTo-Html
                       $LogHTMLDoc = @"
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"  "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<style>BODY{background-color:white;}TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;background-color:LightGreen;}TH{border-width: 1px;padding: 5px;border-style: solid;border-color: black;background-color:Gray;}TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;white-space: nowrap;text-align: center;}</style>
</head><body>
<H3>
~~~~~~~~~~~~~~~~~~~~~~
</H3>
<P>
$($Log2)
</P>
</body></html>
"@

## Output HTML page to your destination ##
$path="$HTMLPath$($log.CmdletName + '.html')"
$LogHTMLDoc | out-file $path  -Append
                    }
## Build HTML Page ##                   
$HTMLHeader = @"
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"  "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<style>BODY{background-color:white;}TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;background-color:LightGreen;}TH{border-width: 1px;padding: 5px;border-style: solid;border-color: black;background-color:Orange;}TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;white-space: nowrap;text-align: center;}</style>
</head><body>
<H2>Exchange Admin Changes
($startDate - $Enddate)
</H2>
"@
$HTMLTableHeader = @"
<table>
<colgroup><col/><col/></colgroup>
<tr><th> Command </th><th> Count </th></tr>
"@
$HTMLEnd = "</body></html>"
$HTMLContent = @"
<table border="2">
<td>
$($HTMLHeader)
$($HtmlTableHeader)
$($ServerTableEntrys)

$($HTMLEnd)</td>
"@

## Dump Main WebPage to your Path ##
$HTMLContent | Out-File "$HTMLPath" + "AdminChanges.html" -Force

