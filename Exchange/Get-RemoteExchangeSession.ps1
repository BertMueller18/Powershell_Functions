<#PSScriptInfo

.VERSION 1.0.0

.GUID 9f12808e-b640-40cb-9fed-e39ef3c9b83e

.AUTHOR Mike Hendrickson

#>

<# 
.SYNOPSIS
 Used to establish a Remote PowerShell session to Exchange with the same capabilities as Exchange Management Shell.
 Exchange Management tools must be installed on the machine using this module.

.DESCRIPTION 
 Used to establish a Remote PowerShell session to Exchange with the same capabilities as Exchange Management Shell.
 Exchange Management tools must be installed on the machine using this module.
 
.PARAMETER ExchangeVersion
 The version of Exchange installed on this machine. Must be 2010, 2013, or 2016.
 Defaults to 2013.

.PARAMETER SuppressBanner
 Whether to suppress the default welcome banner and tips normally displayed when opening Exchange Management Shell.
 Default to $true.

.PARAMETER ViewEntireForest
 Whether Set-ADServerSettings -ViewEntireForest $true should be executed within the remote Exchange session.
 Defaults to $true.

.EXAMPLE
 >Gets a Remote PowerShell Session to Exchange. Suppresses the banner, and defaults to ViewEntireForest.
 PS> .\Get-RemoteExchangeSession.ps1

.INPUTS
None. You cannot pipe objects to Get-RemoteExchangeSession.ps1.

.OUTPUTS
None.
#>

[CmdletBinding()]
param
(
    [ValidateSet("2010","2013","2016")]
    [String]
    $ExchangeVersion = "2013",

    [Bool]
    $SuppressBanner = $true,

    [Bool]
    $ViewEntireForest = $true,

    [Bool]
    $AllowClobber = $false
)

$existingPSSessions = Get-PSSession

#Attempt to reuse the session if we found one
if ($null -ne ($existingPSSessions | Where-Object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.State -like "Opened"} | Select-Object -First 1) -and $AllowClobber -eq $false)
{
    Write-Verbose "Reusing existing Remote Powershell Session to Exchange"
    return
}
#There is an existing session, but it's in a bad state. Remove it.
elseif(($existingPSSessions | Where-Object {$_.ConfigurationName -like "Microsoft.Exchange"}).Count -gt 0 -and $AllowClobber -eq $true)
{
    Write-Verbose "Removing existing Remote Powershell Sessions to Exchange"
    $existingPSSessions | Where-Object {$_.ConfigurationName -like "Microsoft.Exchange"} | Remove-PSSession -Confirm:$false
}

Write-Verbose "Creating new Remote Powershell session to Exchange"

if ($SuppressBanner -eq $true)
{
    New-Alias Get-ExBanner Out-Null
    New-Alias Get-Tip Out-Null
}

if ($null -ne $env:ExchangeInstallPath)
{
	$remoteExchangeScriptPath = Join-Path $env:ExchangeInstallPath "bin\RemoteExchange.ps1"
}
else
{
	throw "Unable to determine installation path of Exchange"
}

#Load shell the same way Exchange Management Shell does
if ($ExchangeVersion -eq "2010")
{
    . $remoteExchangeScriptPath; Connect-ExchangeServer -auto
}
elseif ($ExchangeVersion -eq "2013" -or $ExchangeVersion -eq "2016")
{
    . $remoteExchangeScriptPath; Connect-ExchangeServer -auto -ClientApplication:ManagementShell
}
else
{
    throw "Unsupported server version"
}
        
if ($SuppressBanner -eq $true)
{
    Remove-Item Alias:Get-ExBanner
    Remove-Item Alias:Get-Tip
}

$exchangeSession = Get-PSSession | Where-Object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.State -like "Opened"} | Select-Object -First 1

if ($null -ne $exchangeSession)
{
    $moduleInfo = Import-PSSession $exchangeSession -WarningAction SilentlyContinue -DisableNameChecking -AllowClobber -Verbose:0
    Import-Module $moduleInfo -Global -DisableNameChecking

    if ($ViewEntireForest -eq $true -and $null -ne (Get-Command -Name Set-ADServerSettings -ErrorAction SilentlyContinue))
    {
        Set-ADServerSettings -ViewEntireForest $true
    }
}
