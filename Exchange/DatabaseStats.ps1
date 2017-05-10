Write-Host ""
Write-Host "This tool will scan either one database or multiple databases and provide stats on:"
Write-Host ""
Write-Host " - Number of active mailboxes"
Write-Host " - Total size of active mailboxes (including dumpster)"
Write-Host " - Number of disconnected mailboxes"
Write-Host " - Total size of disconnected mailboxes"
Write-Host ""
Write-Host "For input, you will be able to do a single database, multiple databases, or an imported list (CSV) of databases"
Write-Host ""
Write-Host "At the end, you will have the opportunity to export the output to a CSV."
Write-Host ""
Write-Host ""

$DB_List = @()
$DB_Overview = @()

## Gather data from user

$DB_ListEntry = Read-Host "Enter database name (or type I to import CSV)"

## Import a CSV
	
If ($DB_ListEntry -eq "I")
	{
	$FilePath = Read-Host "Enter path for where the CSV is located"
	$DB_List = Get-Content $FilePath
	}

## Manually enter a database name or list of databases
	
Else
	{
	Do
		{
		$DB_List += $DB_ListEntry
		$DB_ListEntry = $null
		$DB_ListEntry = Read-Host "Enter another database or type C to continue"
		}
	While ($DB_ListEntry -ne "C")
	}

foreach ($DB in $DB_List)
	{
		## For active mailboxes

		$MB_List = Get-Mailbox -Database $DB
		$MB_Stats = $MB_List | measure
		$MB_List | foreach {$MB_Sum += (Get-MailboxStatistics $_.Alias).TotalItemSize.Value.ToMB()}
		$MB_List | foreach {$MB_Sum2 += (Get-MailboxStatistics $_.Alias).TotalDeletedItemSize.Value.ToMB()}

		## For disconnected mailboxes

		$MB_Discon = Get-MailboxStatistics -Database $DB | where {$_.disconnectdate -ne $null} | select DisplayName,MailboxGuid,{$_.TotalItemSize.value.ToMB()}
		$Discon_Sum = $MB_DisCon | measure {$_.TotalItemSize.value.ToMB()} -sum

		## Build array for results
			
		$ActiveMB_inGB = ($MB_Sum + $MB_Sum2)/1024
		$DisconnectedMB_inGB = $Discon_Sum.Sum/1024
		
		$DB_Array = New-Object System.Object
		$DB_Array | Add-Member -MemberType NoteProperty -Name DatabaseName -Value $DB
		$DB_Array | Add-Member -MemberType NoteProperty -Name ActiveMailboxCount -Value $MB_Stats.count
		$DB_Array | Add-Member -MemberType NoteProperty -Name ActiveMailboxSize -Value $ActiveMB_inGB
		$DB_Array | Add-Member -MemberType NoteProperty -Name DisconnectedMailboxCount -Value $Discon_Sum.count
		$DB_Array | Add-Member -MemberType NoteProperty -Name DisconnectedMailboxSize -Value $DisconnectedMB_inGB
			
		$DB_Overview += $DB_Array
		
		## Clear variables

		$MB_Sum = 0
		$MB_Sum2 = 0
	}

## Output results
	
$DB_Overview | ft -AutoSize

$ExportCSV = Read-Host "Would you like to export this to CSV (Y or N)"

If ($ExportCSV -eq "Y")
	{
	$FilePathCSV = Read-Host "Enter path for where you'd like the CSV to be stored"
	$DB_Overview | Export-Csv $FilePathCSV -NoTypeInformation
	}