<# 
    .SYNOPSIS 
    Add IP address(es) to an existing receive connector on selected or all Exchange 2013 Servers

    Thomas Stensitzki 

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE  
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER. 

    Version 1.0, 2014-12-03

    Please send ideas, comments and suggestions to support@granikos.eu 

    .LINK 
    More information can be found at http://www.granikos.eu/en/scripts

    .DESCRIPTION 
    This script adds a given IP address to an existing Receive Connector or reads an
    input file to add more than one IP address to the an existing Receive Connector.
    The script creates a new child directory under the current location of the script.
    The script utilizes the directory as a log directory to store the current remote
    IP address ranges prior modification.
 
    .NOTES 
    Requirements 
    - Windows Server 2008 R2 SP1, Windows Server 2012 or Windows Server 2012 R2  
    - A txt file containing new remote IP address ranges, one per line
      Example:
      192.168.1.1
      192.168.2.10-192.168.2.20
      192.168.3.0/24
    
    Revision History 
    -------------------------------------------------------------------------------- 
    1.0 Initial community release 

    .PARAMETER ConnectorName  
    Name of the connector the new IP addresses should be added to  

    .PARAMETER FileName
    Name of the input file name containing IP addresses

    .PARAMETER ViewEntireForest
    View entire Active Directory forest (default FALSE)
    
    .EXAMPLE 
    .\Add-ReceiveConnectorIpAddress.ps1 -ConnectorName -FileName D:\Scripts\ip.txt

    .EXAMPLE 
    .\Add-ReceiveConnectorIpAddress.ps1 -ConnectorName -FileName .\ip-new.txt -ViewEntireForest $true

#> 
param(
	[parameter(Mandatory=$true,HelpMessage='Name of the Receive Connector',ParameterSetName="RC")]
		[string] $ConnectorName,
	[parameter(Mandatory=$true,HelpMessage='Name of the input file name containing IP addresses',ParameterSetName="RC")]
		[string] $FileName,
    [parameter(Mandatory=$false,HelpMessage='View entire Active Directory forest (default FALSE)',ParameterSetName="RC")]
        [boolean] $ViewEntireForest = $false
)

Set-StrictMode -Version Latest

$tmpFileFolderName = "ReceiveConnectorIpAddresses"
$tmpFileLocation = ""
# Timestamp for use in filename, adjust formatting to your regional requirements
$timeStamp = Get-Date -Format "yyyy-MM-dd hhmmss"

# FUNCTIONS --------------------------------------------------

function CheckLogPath {
    $script:tmpFileLocation = Join-Path -Path $PSScriptRoot -ChildPath $tmpFileFolderName
    if(-not (Test-Path $script:tmpFileLocation)) {
        Write-Verbose "New file folder created"
        New-Item -ItemType Directory -Path $script:tmpFileLocation -Force | Out-Null
    }
}

function CheckReceiveConnector([string]$Server) {

    Write-Verbose "Checking Server: $Server"

    # Fetch receive connector from server
    $targetRC = Get-ReceiveConnector -Server $Server | ?{$_.name -eq $ConnectorName} -ErrorAction SilentlyContinue

    if($targetRC -ne $null) {
        Write-Verbose "Found connector $ConnectorName on server $Server"
	    SaveConnectorIpRanges $targetRC
    }
    else {
        Write-Output "INFO: Connector $ConnectorName NOT found on server $Server"
    }
}

function SaveConnectorIpRanges($ReceiveConnector) {
    # Create a list of currently configured IP ranges 
    $tmpRemoteIpRanges = ""
    foreach ( $remoteIpRange in ($ReceiveConnector).RemoteIPRanges ) {
        $tmpRemoteIpRanges += "`n$remoteIpRange"			
	}

    Write-Verbose $tmpFileLocation
    
    # Save current remote IP ranges for connector to disk
    $fileIpRanges = ("$tmpFileLocation\$ConnectorName-$timeStamp-Export.txt").ToUpper()
    Write-Verbose "Saving current IP ranges to: $fileIpRanges"
    Write-Output $tmpRemoteIpRanges | Out-File -FilePath $fileIpRanges -Force

    # Fetch new IP ranges from disk
    $newIpRangesFileContent = ""
    if(Test-Path $FileName) {
	    Write-Verbose "Reading file $FileName"
	    $newIpRangesFileContent = Get-Content -Path $FileName
    }

    # add new IP ranges, if file exsists and had content
    if($newIpRangesFileContent -ne ""){
        foreach ($newIpRange in $newIpRangesFileContent ){
	        Write-Verbose "Checking new Remote IP range $newIpRange in $fileIpRanges"
            # Check if new remote IP range already exists in configure remote IP range of connector
	        $ipSearch = (Select-String -Pattern $newIpRange -Path $fileIpRanges )
	        if ($ipSearch -ne $null ){
                # Remote IP range exists, nothing to do here
                Write-Output "Remote IP range [$newIpRange] already exists"
		    }
		    else {
                # Remote IP range does not exist, add range to receive connector object
	            Write-Output "Remote IP range [$newIpRange] will be added to $ConnectorName" 
	            $ReceiveConnector.RemoteIPRanges += $newIpRange
	        }
        }
        # save changes to receive connector
        Set-ReceiveConnector -RemoteIPRanges $ReceiveConnector.RemoteIPRanges -Identity $ReceiveConnector.Identity | Sort -Unique
    }    
}

# MAIN -------------------------------------------------------

if($ViewEntireForest) {
    Write-Verbose "Setting ADServerSettings -ViewEntireForest $true"
    Set-ADServerSettings -ViewEntireForets $true
}

CheckLogPath

# Fetch all Exchange 2013 Servers
$allExchangeServers = Get-ExchangeServer | ?{($_.AdminDisplayVersion.Major -eq 15) -and ([string]$_.ServerRole).Contains("ClientAccess")} 

foreach($Server in $AllExchangeServers) {
    Write-Output "Checking receive connector $ConnectorName on server $Server"
    CheckReceiveConnector $Server
}