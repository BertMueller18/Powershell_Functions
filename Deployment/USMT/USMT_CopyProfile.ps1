# Prompt user for elevated credentials for remote access/UNC/PsExec privileges
$Credentials = Get-Credential -Credential Your.Domain\ # Make sure to change this to your domain so it includes it in the UserName field
$InvokingUser = $Credentials.UserName
$ElevatedPassword = $Credentials.GetNetworkCredential().Password

# Prompt user for source, destination, and profile information
$SourcePCName = $Read-Host -Prompt "Enter the old PC hostname to capture a user profile for"
$TargetPCName = $Read-Host -Prompt "`nEnter the new PC hostname to restore the user profile to"
$CapturedUser = $Read-Host -Prompt "`nEnter the username for the profile you want to capture"

# Build and store remote PsExec and USMT command arguments in a string
# to run on the OLD PC you are capturing FROM
#
# Make sure to replace the "\\YourServer" paths to your network share
# with USMT files (loadstate and scanstate and XML USMT config files)
#
$CaptureArgs = "-accepteula \\$SourcePCName -u "   $InvokingUser   " -p """   $ElevatedPassword   """ \\YourServer\USMTFiles\scanstate /localonly \\YourServer\USMTFiles\ProfileStoreLocation\"   $CapturedUser   " /ue:*\* /ui:Your.Domain\"   $CapturedUser   " /i:\\YourServer\USMTFiles\miguser.xml /i:\\YourServer\USMTFiles\migapp.xml /i:\\YourServer\USMTFiles\migsl.xml /l:C:\scanlog.log /o /c"

# Build and store remote PsExec and USMT command arguments in a string
# to run on the NEW PC you are restoring the profile TO
#
# Make sure to replace the "\\YourServer" paths to your network share
# with USMT files (loadstate and scanstate and XML USMT config files)
#
$RestoreArgs = "-accepteula \\$TargetPCName -u "   $InvokingUser   " -p """   $ElevatedPassword   """ \\YourServer\USMTFiles\loadstate.exe \\YourServer\USMTFiles\ProfileStoreLocation\"   $CapturedUser   " /i:\\YourServer\USMTFiles\miguser.xml /i:\\YourServer\USMTFiles\migapp.xml /i:\\YourServer\USMTFiles\migsl.xml"

# Start PsExec with the CaptureArgs to capture the user-profile on the old PC and wait for completion
$proc = Start-Process -Filepath "PsExec.exe" -ArgumentList $CaptureArgs -Credential $Credentials -WorkingDirectory "\\NetworkPath\To\Folder\Containing\PsExec\" -PassThru
do {start-sleep -Milliseconds 500}
until ($proc.HasExited)

# Start PsExec with the RestoreArgs to restore the user-profile to the new PC and wait for completion
$proc2 = Start-Process -Filepath "PsExec.exe" -ArgumentList $RestoreArgs -Credential $Credentials -WorkingDirectory "\\NetworkPath\To\Folder\Containing\PsExec\" -PassThru
do {start-sleep -Milliseconds 500}
until ($proc2.HasExited)

# Inform user of completion
Read-Host -Prompt "`nUSMT Process has completed! Check USMT logs for details. (Press Enter to exit) "
