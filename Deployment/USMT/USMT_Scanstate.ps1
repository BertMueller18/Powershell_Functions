
# USMT_SCANSTATE.ps1
#
$Credentials = Get-Credential -Credential SmartDeployTest\
$InvokingUser = $Credentials.UserName
$ElevatedPassword = $Credentials.GetNetworkCredential().Password

$SourcePCName = Read-Host -Prompt "Enter hostname of source PC"
$CapturedUser = Read-Host -Prompt "Enter the username for the profile you want to capture"
# must check for valid values

$sb = {
	Param(
		$SourcePCName,
		$CapturedUser
	)
	$scanstate = '\\smartdeploytest\USMT\amd64\scanstate.exe'
	$arglist = @(
		'/i:\\smartdeploytest\SAFE\$SourcePCName',
		'/i:\\smartdeploytest\USMT\amd64\MigApp.xml',
		'/i:\\smartdeploytest\USMT\amd64\MigDocs.xml',
		'/i:\\smartdeploytest\USMT\amd64\Testxml.xml',
		'/l:\\smartdeploytest\SAFE\$SourcePCNameSCANSTATE.log',
		'/localonly',
		'/o',
		'/c',
		'/vsc',
		'/ue:*\*',
		'/ui:example\$CapturedUser'
	)
	$process = Start-Process -FilePath $scanstate -ArgumentList $arglist -Wait -PassThru
	# check process status
}

Invoke-Command $sb -ComputerName $SourcePCName -ArgumentList $SourcePCName,$CapturedUser

Read-Host -Prompt "USMT-SAVESTATE Process has completed! Check USMT logs for details. (Press Enter to exit)"
