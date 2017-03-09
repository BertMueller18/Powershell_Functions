<# 
.SYNOPSIS
	Examines certificates on local machine, and emails admin if any are expired or about to expire.

.DESCRIPTION
	Examines certificates on local machine, and emails admin if any are expired or about to expire.
		
.NOTES
	Version    	      		: v1.3 - See changelog at https://www.ucunleashed.com/1358
	Wish list							: 
	Rights Required				: Local admin for installation
  Sched Task Required		: Yes - install option will automatically create scheduled task
	Lync/Skype4B Version  : N/A
	Exchange Version			: N/A
  Author/Copyright      : © Pat Richard, Office Servers and Services (Skype for Business) MVP - All Rights Reserved
  Email/Blog/Twitter    : pat@innervation.com  https://ucunleashed.com  @patrichard
  Donations             : https://www.paypal.me/PatRichard
	Dedicated Post				: https://www.ucunleashed.com/1360
	Disclaimer            : You running this script means you won't blame author(s) if this breaks your stuff. This script is
                          provided AS IS without warranty of any kind. Author(s) disclaim all implied warranties including,
                          without limitation, any implied warranties of merchantability or of fitness for a particular
                          purpose. The entire risk arising out of the use or performance of the sample scripts and
                          documentation remains with you. In no event shall author(s) be liable for any damages whatsoever
                          (including, without limitation, damages for loss of business profits, business interruption,
                          loss of business information, or other pecuniary loss) arising out of the use of or inability
                          to use the script or documentation. Neither this script, nor any part of it other than those
                          parts that are explicitly copied from others, may be republished without author(s) express written 
                          permission.
	Acknowlegements 			: http://www.dimensionit.tv/check-for-certificate-expiration-with-powershell/#comment-5
	Assumptions						: ExecutionPolicy of AllSigned (recommended), RemoteSigned or Unrestricted (not recommended)
	Limitations						: 
	Known issues					: None yet, but I'm sure you'll find some!

.LINK     
	https://www.ucunleashed.com/1360

.EXAMPLE 
	.\New-ExpiringCertificatesReminder.ps1
		
	Description
	-----------
	Runs the script normally.

.EXAMPLE 
	.\New-ExpiringCertificatesReminder.ps1 -Install
	
	Description
	-----------
	This will create a scheduled task in Windows to run the script once per day. The details of the scheduled task can be tweaked manually in Windows.

.INPUTS
	None. You cannot pipe objects to this script.
#> 
#Requires -Version 3.0

[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = "Default")]
param(
	[parameter(ValueFromPipelineByPropertyName, ParameterSetName = "Install")] 
	[switch] $Install,
	
	[parameter(ValueFromPipelineByPropertyName, ParameterSetName = "Default")]
	[ValidateNotNullOrEmpty()] 
	[string] $Company = "Contoso Ltd",
	
	[parameter(ValueFromPipelineByPropertyName, ParameterSetName = "Default")]
	[ValidateNotNullOrEmpty()]
	[string] $PSEmailServer = "10.9.0.11",
	
	[parameter(ValueFromPipelineByPropertyName, ParameterSetName = "Default")]
	[ValidateNotNullOrEmpty()]
	[string] $EmailFrom = "Help Desk <helpdesk@contoso.com>",
	
	[parameter(ValueFromPipelineByPropertyName, ParameterSetName = "Default")]
	[ValidateNotNullOrEmpty()]
	[string] $EmailTo = "pat@contoso.com",
	
	[parameter(ValueFromPipelineByPropertyName, ParameterSetName = "Default")] 
	[string] $HelpDeskPhone = "(586) 555-1010",	
	
	[parameter(ValueFromPipelineByPropertyName, ParameterSetName = "Default")]
	[ValidatePattern("^http")]
	[string] $HelpDeskURL = "https://intranet.contoso.com/",
	
	[parameter(ValueFromPipelineByPropertyName, ParameterSetName = "Default")] 
	[ValidatePattern("^http")]
	[string] $ImagePath = "http://www.contoso.com/images",
	
	[parameter(ParameterSetName = "Default")]
	[ValidateNotNullOrEmpty()]
	[string] $ScriptName = $MyInvocation.MyCommand.Name,
	
	[parameter(ParameterSetName = "Default")]
	[ValidateNotNullOrEmpty()]
	[string] $ScriptPathAndName = $MyInvocation.MyCommand.Definition,
	
	[parameter(ParameterSetName = "Default")]
  [int] $threshold = 15,
  
  [parameter(ParameterSetName = "Default")]
	[DateTime] $deadline = (Get-Date).AddDays($threshold),
  
  [parameter(ParameterSetName = "Default")]
	[string] $tables = $null,
  
  [parameter(ParameterSetName = "Default")]
	[bool] $expiring = $false,
  
  [parameter(ParameterSetName = "Default")]
  [int] $totalcerts = 0
)

#region functions
function Install	{
	[CmdletBinding(SupportsShouldProcess)]
	param()
	Write-Verbose "Begin install mode"
	$error.clear()
	Write-Output "Creating scheduled task `"$ScriptName`"..."
	$TaskCreds = Get-Credential("$env:userdnsdomain\$env:username")
	$TaskPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($TaskCreds.Password))
	schtasks /create /tn $ScriptName /tr "$env:windir\system32\windowspowershell\v1.0\powershell.exe -noprofile -command $ScriptPathAndName" /sc Daily /st 07:30 /ru $TaskCreds.UserName /rp $TaskPassword | Out-Null
	if (-not($error)){
		Write-Host "Installation complete!" -ForegroundColor green
	}else{
		Write-Host "Installation failed!" -ForegroundColor red
	}
	Write-Verbose "Completed install mode"
	exit
} # end function Install

function Remove-ScriptVariables {
	<#
  .SYNOPSIS
    Removed variables defined by a script.

  .DESCRIPTION
    Removed variables defined by a script. Beneficial if the PowerShell session is going to remain open, as removing the variables frees up the memory used.

  .NOTES
    Version               : 1.1
    Wish list             : Better error trapping
    Rights Required       : N/A
    Sched Task Required   : No
    Lync/Skype4B Version  : N/A
    Author/Copyright      : © Pat Richard, Office Servers and Services (Skype for Business) MVP - All Rights Reserved
    Email/Blog/Twitter    : pat@innervation.com  https://ucunleashed.com  @patrichard
    Donations             : https://www.paypal.me/PatRichard
    Dedicated Post        :	https://www.ucunleashed.com/247
    Disclaimer            : You running this script means you won't blame author(s) if this breaks your stuff. This script is
                            provided AS IS without warranty of any kind. Author(s) disclaim all implied warranties including,
                            without limitation, any implied warranties of merchantability or of fitness for a particular
                            purpose. The entire risk arising out of the use or performance of the sample scripts and
                            documentation remains with you. In no event shall author(s) be liable for any damages whatsoever
                            (including, without limitation, damages for loss of business profits, business interruption,
                            loss of business information, or other pecuniary loss) arising out of the use of or inability
                            to use the script or documentation. Neither this script, nor any part of it other than those
                            parts that are explicitly copied from others, may be republished without author(s) express written 
                            permission.
    Acknowledgements      :
    Assumptions           : ExecutionPolicy of AllSigned (recommended), RemoteSigned, or Unrestricted (not recommended)
    Limitations           :
    Known issues          : None yet, but I'm sure you'll find some!

  .LINK
		https://www.ucunleashed.com/247

  .EXAMPLE
    Remove-ScriptVariables

    Description
    -----------
    Removes the variables in memory that were created when a .ps1 script is run.

  .INPUTS
    You can pipe file path objects to this script.
  #>
	[cmdletBinding(SupportsShouldProcess)]
	param(
		[Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)]
		[ValidateNotNullOrEmpty()]
		[string] $path
	)
	process{
		Write-Verbose "Begin removing script variables from memory"
		$result = Get-Content $path |  
		ForEach { 
			if ($_ -match '(\$.*?)\s*=') {      
				$matches[1]  | Where-Object { $_ -notlike '*.*' -and $_ -notmatch 'result' -and $_ -notmatch 'env:'}  
			}  
		}  
		ForEach ($v in ($result | Sort-Object | Get-Unique)){		
			Remove-Variable ($v.replace("$","")) -EA 0
		}
	}
	end{
		Write-Verbose "Completed removing script variables from memory"
	}
} # end function Remove-ScriptVariables
#endregion functions

if ($install){	
	Install
}

$EmailSent = "Expiring certificates email sent by " + $ScriptName + " " + (Get-Date)
[object] $evt = New-Object System.Diagnostics.EventLog("Application")
[string] $evt.Source = $ScriptName
$infoevent = [System.Diagnostics.EventLogEntryType]::Information
[string] $EventLogText = "Beginning processing"
$evt.WriteEntry($EventLogText,$infoevent,70)

Write-Verbose "Getting local machine certificate information"
$mycerts = Get-ChildItem CERT:\LocalMachine\My

ForEach ($cert in $mycerts){
	# if ($cert.NotAfter -le $deadline -and ($cert.NotAfter -ge (Get-Date))) {
	if ($cert.NotAfter -le $deadline) {
		$totalcerts++
  	$expiring = $true
  	$Issuer = $cert.Issuer
  	$Subject = $cert.Subject
  	$expires = $cert.NotAfter
  	$NotAfter = ($cert.NotAfter - (Get-Date)).Days
  	if ($notafter -le 0){$extra = "[EXPIRED!]"}
  	$thumbprint = $cert.Thumbprint
  	$friendlyname = $cert.friendlyName

  	$Tables += @"
  	<table border="1" width="100%" style="font: 12px arial">
  		<tr><td nowrap>Server</td><td>$env:ComputerName</td></tr>
  		<tr><td nowrap>Friendly Name</td><td>$friendlyname</td></tr>
  		<tr><td nowrap>Issuer</td><td>$Issuer</td></tr>
  		<tr><td nowrap>Subject</td><td>$Subject</td></tr>
  		<tr><td nowrap>Expires</td><td>$expires ($NotAfter days) <span style="color: red; font-weight: bold">$extra</span></td></tr>
  		<tr><td nowrap>Thumbprint</td><td>$thumbprint</td></tr>
  	</table><br />
"@
	}
}

if ($expiring)	{
	[string]$EmailBody = @"
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
	</head>
<body>
	<table id="email" border="0" cellspacing="0" cellpadding="0" width="655" align="center">
"@

if ($HelpDeskURL){
$emailbody += @"
		<tr>
			<td align="left" valign="top"><img src="$ImagePath/spacer.gif" alt="" width="46" height="28" align="absMiddle"><font style="font-size: 10px; color: #000000; line-height: 16px; font-family: Verdana, Arial, Helvetica, sans-serif">If this e-mail does not appear properly, please <a href="$HelpDeskURL" style="font-weight: bold; font-size: 10px; color: #cc0000; font-family: verdana, arial, helvetica, sans-serif; text-decoration: underline" target="_blank">click here</a>.</font></td>
		</tr>
		<tr>
			<td height="121" align="left" valign="bottom"><a href="$HelpDeskURL" target="_blank"><img src="$ImagePath/header.gif" border="0" alt="Header" width="655" height="121"></a></td>
		</tr>
"@
}else{
$emailbody += @"
		<tr>
			<td height="121" align="left" valign="bottom"><img src="$ImagePath/header.gif" border="0" alt="Header" width="655" height="121"></td>
		</tr>
"@
}
$emailbody += @"
		<tr>
			<td>
				<table id="body" border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td width="1" align="left" valign="top" bgcolor="#a8a9ad"><img src="$ImagePath/spacer50.gif" alt="" width="1" height="50"></td>
						<td><img src="$ImagePath/spacer.gif" alt="" width="46" height="106"></td>
						<td id="text" width="572" align="left" valign="top" style="font-size: 12px; color: #000000; line-height: 17px; font-family: Verdana, Arial, Helvetica, sans-serif">
							<p style="font-weight: bold">Hello,</p>
							<p>We've noticed that the following certificate(s) have expired or are expiring within the next $threshold days. Please renew them as soon as possible to avoid a possible disruption in service.</p>
							$Tables
							<p>Thank you!<br />
								The $Company Help Desk<br />
								$HelpDeskPhone
							</p>
							</td>
							<td width="49" align="left" valign="top"><img src="$ImagePath/spacer50.gif" alt="" width="49" height="50"></td>
							<td width="1" align="left" valign="top" bgcolor="#a8a9ad"><img src="$ImagePath/spacer50.gif" alt="" width="1" height="50"></td>
						</tr>
					</table>
					<table id="footer" border="0" cellspacing="0" cellpadding="0" width="655">
						<tr>
							<td colspan="3"><img src="$ImagePath/footer.gif" alt="© 2010 $company" width="655" height="81"></td>
						</tr>
					</table>
					<table border="0" cellspacing="0" cellpadding="0" width="655" align="center">
						<tr>
							<td align="left" valign="top"><img src="$ImagePath/spacer.gif" alt="" width="36" height="1"></td>
							<td align="left" valign="top"><font face="Verdana" size="1" color="black">
"@
if ($HelpDeskURL){
$emailbody += @"
							<p>This email was sent by an automated process. If you would like to comment on it, please visit <a href="$HelpDeskURL"><font color="#ff0000"><u>click here</u></font></a>.</p>
"@
}
$emailbody += @"
								<p style="color: #009900;"><font face="Webdings" size="4">P</font> Please consider the environment before printing this email.</p></font>
							</td>
							<td align="left" valign="top"><img src="$ImagePath/spacer.gif" alt="" width="36" height="1"></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
	</body>
"@

	Send-MailMessage -To $EmailTo -Subject "$totalcerts certificate(s) expired or expiring soon!" -Body $emailbody -From $EmailFrom -Priority High -BodyAsHtml
}
$EventLogText = "Finished processing $totalcerts certificate(s). `n`nFor more information about this script, run Get-Help .\$ScriptName or see https://www.ucunleashed.com/1360."
$evt.WriteEntry($EventLogText,$infoevent,70)

Remove-ScriptVariables $ScriptPathAndName