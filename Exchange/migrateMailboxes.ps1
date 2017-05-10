Param (
	[Parameter(Position=0, Mandatory=$False, HelpMessage="The source server to move mailboxes from.")]
	[ValidateNotNullOrEmpty()]
	[String] $Source2007Server,
	
	[Parameter(Position=1, Mandatory=$False, HelpMessage="The source database to move mailboxes from in the format of `"Server\StorageGroup\Database`", e.g.: `"MMEM01\SG2\DB-AB`".")]
	[ValidateNotNullOrEmpty()]
	[String] $Source2007Database,
	
	[Parameter(Position=2, Mandatory=$False, HelpMessage="The file containing users to be moved.")]
	[ValidateNotNullOrEmpty()]
	[String] $SourceFile,
	
	[Parameter(Position=3, Mandatory=$False, HelpMessage="The destination server to move mailboxes to.")]
	[ValidateNotNullOrEmpty()]
	[String] $Target2010Server,

	[Parameter(Position=4, Mandatory=$False, HelpMessage="Specifies the number of bad items to skip if the move encounters corrupted items in the mailbox.")]
	[ValidateScript({$_ -ge 0})]
	[Int] $BadItemLimit = 0,

	[Parameter(Position=5, Mandatory=$False, HelpMessage="Specifies the number of bad items to skip if the move encounters corrupted items in the mailbox.")]
	[Switch] $SuspendWhenReadyToComplete,
	
	[Parameter(Position=6, Mandatory=$False, HelpMessage="The maximum number of mailboxes to be moved.")]
	[ValidateNotNullOrEmpty()]
	[Int] $NumberOfMoves,
	
	[Parameter(Position=7, Mandatory=$False, HelpMessage="Specifies if error messages should be sent by e-mail.")]
	[Switch] $SendErrorsByEmail
)



Function produceReport
{
	#$moveRequests = Get-MoveRequest -ResultSize Unlimited | Sort DisplayName
	
	# Much slower but will ensure we only get the requests for this job (could possibly use BatchName instead)
	$moveRequests = (Get-MoveRequest -ResultSize Unlimited) | ? {(Get-Date (Get-MoveRequestStatistics $_).StartTimeStamp) -ge $global:startDate} | Sort DisplayName

	$global:moveReport = "<HTML> <BODY>"

	# Header of the report (title and date)
	$global:moveReport = "<font size=""1"" face=""Arial,sans-serif"">
				<h3>Mailbox Move Report</h3>
				<h5>$((Get-Date).ToString())</h5>
				</font>"

	# Header of the first table with the Status and Count columns
	$global:moveReport += "<table border=""0"" cellpadding=""3"" style=""font-size:8pt;font-family:Arial,sans-serif"">	
						<tr align=""center"" bgcolor=""#FFD700"">
						<th>Status</th>
						<th>Count</tr>"

	# Group all the move requests based on their status and populate the table
	$AlternateRow = 1
	ForEach ($move in ($moveRequests | Group Status))
	{
		$global:moveReport += "<tr"

		If ($AlternateRow) {
			# If it is an alternate row, we set the background color to grey
			$global:moveReport += " style=""background-color:#dddddd"""
			$AlternateRow = 0
		} Else {
			$AlternateRow = 1
		}

		$global:moveReport += "><td>$($move.Name)</td>"
		$global:moveReport += "<td align=""center"">$($move.Count)</td>"
		$global:moveReport += "</tr>";
	}


	# Check how many mailboxes didn't get moved
	$notMoved = $global:intCount - $moveRequests.Count
	If ($notMoved -gt 0) {
		# Write how many skipped mailboxes
		$global:moveReport += "<tr"
		If ($AlternateRow) {
			$global:moveReport += " style=""background-color:#dddddd"""
			$AlternateRow = 0
		} Else {
			$AlternateRow = 1
		}
		
		$global:moveReport += "><td>Skipped</td>"
		$global:moveReport += "<td align=""center"">$($notMoved)</td>"
		$global:moveReport += "</tr>";
	}

	$global:moveReport += "</table><br/>"


	# If no users were migrated, don't generate the rest of the report (which would be empty)
	If ($moveRequests.Count -eq $null) { $global:moveReport += "</body></html>"; Return }
	
	$global:moveReport += "<table border=""0"" cellpadding=""3"" style=""font-size:8pt;font-family:Arial,sans-serif"">	
						<tr align=""center"" bgcolor=""#FFD700"">
						<th>User</th>
						<th>Status</th>
						<th>Bad Items</th></tr>"


	$AlternateRow = 1
	$moveRequests | ForEach {
		$global:moveReport += "<tr"
		
		If ($_.Status -match "Failed") {
			# If a mailbox move failed, we set the background color to Red
			$global:moveReport += " style=""background-color:#FF0000"""
		} Else {
			If ($AlternateRow) {
				$global:moveReport += " style=""background-color:#dddddd"""
				$AlternateRow = 0
			} Else {
				$AlternateRow = 1
			}
		}
		
		$global:moveReport += "><td>$($_.DisplayName)</td>"
		$global:moveReport += "<td align=""center"">$($_.Status)</td>"
		$global:moveReport += "<td align=""center"">$((Get-MoveRequestStatistics $_.Identity).BadItemsEncountered)</td>"

		$global:moveReport += "</tr>";
	}

	$global:moveReport += "</table><br/>"
	$global:moveReport += "</BODY></HTML>"
}


# Return the position of the smallest DB in the array
Function GetSmallestDB
{
	[Int] $intSmallestSize = -1
	
	# Go through the arrSizes array to get the smallest DB
	For ([Int] $x = 0; $x -lt $intNumberDBs; $x++) {
		If (($intSmallestSize -eq -1) -or ($arrDBSizes[$x, 1] -lt $intSmallestSize)) {
			$intSmallestSize = $arrDBSizes[$x, 1]
			$intSmallestPos = $x
		}
	}

	# Return the position on arrDBSizes array of the smallest DB
	Return $intSmallestPos
}


Function printDBSizesArray
{
	For ([Int] $x = 0; $x -lt $intNumberDBs; $x++) {
		Write-Host $arrDBSizes[$x, 0]  $arrDBSizes[$x, 1]
	}
	
	Write-Host "`n"
}


Function moveUser ($user)
{
	# Get the smallest DB to move the user to and update the arrDBSizes array
	$intSmallestPos = GetSmallestDB
	$arrDBSizes[$intSmallestPos, 1] += (Get-MailboxStatistics $user).TotalItemSize.Value.ToMB()

	Write-Host "Moving $user to", $arrDBSizes[$intSmallestPos, 0]
	
	# FIX!!!!!!!!!!!!
	If ($BadItemLimit -gt 50) {
		New-MoveRequest -Identity $user -TargetDatabase $arrDBSizes[$intSmallestPos, 0] -BadItemLimit $BadItemLimit -AcceptLargeDataLoss -SuspendWhenReadyToComplete:$SuspendWhenReadyToComplete -Confirm:$false
	} Else {
		New-MoveRequest -Identity $user -TargetDatabase $arrDBSizes[$intSmallestPos, 0] -BadItemLimit $BadItemLimit -SuspendWhenReadyToComplete:$SuspendWhenReadyToComplete -Confirm:$false
	}

	# If we have already moved as many users as we wanted, then stop
	$global:intCount++
	If ($global:intCount -eq $NumberOfMoves) { finishMigration }

	Start-Sleep -S 1
}



Function finishMigration
{
	#Write-Host "`nFinal Result:`n" -ForegroundColor Green
	#printDBSizesArray
	
	$moveRequests = Get-MoveRequest -ResultSize Unlimited
	While (($moveRequests | ? {$_.Status -match "InProgress"}) -or ($moveRequests | ? {$_.Status -eq "Queued"})) {
		# Sleep for 5 minutes
		Start-Sleep -Seconds 300
		$moveRequests = Get-MoveRequest -ResultSize Unlimited
	}

	produceReport
	# We can also export the report to a file
	#$global:moveReport | Out-File "move_report.html"

	Send-MailMessage -From "AdminMotaN@letsexchange.com" -To "nuno@letsexchange.com" -Subject "Move Complete!" -Body $global:moveReport -BodyAsHTML -SMTPserver "smtp.letsexchange.com" -DeliveryNotificationOption onFailure
	Exit
}


Function sendError ([String] $strError)
{
	Write-Warning $strError
	If ($SendErrorsByEmail) { Send-MailMessage -From "AdminMotaN@letsexchange.com" -To "nuno@letsexchange.com" -Subject "Move Error!" -Body $strError -SMTPserver "smtp.letsexchange.com" -DeliveryNotificationOption onFailure -Priority High }
	Exit
}


# Validate Input
If ($Source2007Database -and $Source2007Server) { sendError "Please use only -Source2007Server or -Source2007Database, not both." }

If ($Source2007Server) {
	# Verify that the mentioned server is actually a 2007 Mailbox server and that there are mailboxes in it to be moved
	$2007server = Get-ExchangeServer $Source2007Server
	If (!$2007server.IsMailboxServer -or $2007server.AdminDisplayVersion -notmatch "Version 8") { sendError "`"$Source2007Server`" is not a valid Exchange 2007 Mailbox server!" }

	If ((Get-Mailbox -ResultSize Unlimited -Server $Source2007Server | Measure-Object).Count -eq 0) { sendError "There are no users in `"$Source2007Server`" server." }
}

If ($Source2007Database) {
	If ($Source2007Database -notlike "*\*\*") { sendError "Please use `"Server\StorageGroup\Database`" format."	}
	
	$2007DB = Get-MailboxDatabase $Source2007Database
	If ($2007DB.ExchangeVersion -notlike "*(8*") { sendError "`"$Source2007Database`" is not a valid Exchange 2007 database." }
	
	If ((Get-Mailbox -ResultSize Unlimited -Database $Source2007Database | Measure-Object).Count -eq 0) { sendError "There are no users in `"$Source2007Database`" database." }
}

If ($Target2010Server) {
	$2010server = Get-ExchangeServer $Target2010Server
	If (!$2010server.IsMailboxServer -or $2010server.AdminDisplayVersion -notmatch "Version 14") {
		sendError "`"$Target2010Server`" is not a valid Exchange 2010 Mailbox server!"
	}
	
	$DBs = Get-MailboxDatabase -Server $Target2010Server -Status | ? {$_.IsExcludedFromProvisioning -eq $False -and $_.Mounted -eq $True} | Sort Name
} Else {
	# If we don't specify a target 2010 mailbox server, then get all available databases in 2010
	$DBs = Get-MailboxDatabase -Status | ? {$_.IsExcludedFromProvisioning -eq $False -and $_.Mounted -eq $True} | Sort Name
}

# Check if we found any databases to move the users to
If ($DBs.Count -lt 1) { sendError "No suitable target databases were found!" }


# Initialize variables
[String] $global:moveReport = $null
[DateTime] $global:startDate = Get-Date
[Int] $global:intCount = 0
[Int] $intNumberDBs = $DBs.Count
$arrDBSizes = New-Object 'Object[,]' $intNumberDBs,2
[Int] $x = 0


# Get the sizes of all DBs that will be used for the move and save them into the array
ForEach ($db in $DBs) {
	$arrDBSizes[$x, 0] = $db.Identity
	$arrDBSizes[$x, 1] = ($db.DatabaseSize - $db.AvailableNewMailboxSpace).ToMB()
	$x++
}


#printDBSizesArray


# The source is an Exchange 2007 server, so move all mailboxes found in the server
If ($Source2007Server) {
	ForEach ($mbx in (Get-Mailbox -ResultSize Unlimited -Server $Source2007Server)) {
		moveUser $mbx
	}
}


# The source is a specific DB, so move only mailboxes found in that DB
If ($Source2007Database) {
	ForEach ($mbx in (Get-Mailbox -ResultSize Unlimited -Database $Source2007Database))	{
		moveUser $mbx
	}
}


If ($SourceFile) {
	# Read the users to move from a CSV file containing the alias of the users in a column named 'User'
	Import-Csv $SourceFile | ForEach { moveUser $_.User }
}


# If the user didn't specify any source, then move all Exchange 2007 mailboxes
If (!$Source2007Server -and !$Source2007Database -and !$SourceFile) {
	ForEach ($mbx in (Get-MailboxServer | ? {$_.AdminDisplayVersion -match "version 8"} | Get-Mailbox -ResultSize Unlimited)) { moveUser $mbx }
}


finishMigration