<#
.Synopsis
   collect_report.ps1 is  a PowerShell Script to fetch basic Mailbox level details like DisplayName, Email Address, RecipientType,User Department etc.
   This script clubs multiple cmdlets to fecth different attributes.

.DESCRIPTION
   collect_report.ps1 is a PowerShell script to fetch mailbox level details. The script freches these details and writes them to a CSV file.
   The Script appends the date on which it was run on the CSV file. The Script is designed by combining multiple cmdlets and combines all data into the CSV file.
   The CSV filename will be like Mailbox_Report_yyyyMMdd.csv. The csv file will be written only after fetching all the mailboxes and processing them.

.EXAMPLE
   .\collect_report.ps1

.RAISE QUERIES TO
   Noble K Varghese [noblekalathoor@outlook.com]
#>

$report = @()
$reportfile = "Mailbox_Report_$((get-date).tostring(“yyyyMMdd”)).csv"

Write-Host "Fetching Mailboxes. Please Wait!"
$Mailboxes = Get-Mailbox -Resultsize Unlimited

$i = 1
$count = $Mailboxes.Count

foreach ($Mailbox in $Mailboxes) {
    
    $MailboxStat = Get-MailboxStatistics -Identity $($Mailbox.Alias)
    $Userdetails = Get-User -Identity $($Mailbox.Alias)

    $reportObj = New-Object PSObject
	$reportObj | Add-Member NoteProperty -Name "DisplayName" -Value $Mailbox.DisplayName
    $reportObj | Add-Member NoteProperty -Name "PrimaryEmailAddress" -Value "$($Mailbox.PrimarySmtpAddress.Local)@$($Mailbox.PrimarySmtpAddress.Domain)"
    $reportObj | Add-Member NoteProperty -Name "Mailbox Type" -Value $Mailbox.RecipientTypeDetails
    $reportObj | Add-Member NoteProperty -Name "Title" -Value $Userdetails.Title
	$reportObj | Add-Member NoteProperty -Name "Department" -Value $Userdetails.Department
    $reportObj | Add-Member NoteProperty -Name "Division" -Value $Mailbox.CustomAttribute3
    $reportObj | Add-Member NoteProperty -Name "Country" -Value $Userdetails.CountryOrRegion.DisplayName
    $reportObj | Add-Member NoteProperty -Name "City" -Value $Userdetails.City
    $reportObj | Add-Member NoteProperty -Name "Office" -Value $Userdetails.Office
    $reportObj | Add-Member NoteProperty -Name "Total Mailbox Size (MB)" -Value "$($MailboxStat.TotalItemSize.Value.ToMB()) MB"
    $reportObj | Add-Member NoteProperty -Name "Mailbox Recoverable Item Size (Mb)" -Value "$($MailboxStat.TotalDeletedItemSize.Value.ToMB()) MB"
    $reportObj | Add-Member NoteProperty -Name "Mailbox Items" -Value $MailboxStat.ItemCount
    $reportObj | Add-Member NoteProperty -Name "Primary Mailbox Database" -Value $Mailbox.Database.Name
    $reportObj | Add-Member NoteProperty -Name "Organizational Unit" -Value $Mailbox.OrganizationalUnit
    
    $report += $reportObj

    Write-Progress -activity "Processed User $($Mailbox.Alias)" -status "$i Out Of $count completed" -percentcomplete ($i / $count*100)

    $i++

}

$report | Export-Csv -NoTypeInformation $reportfile -Encoding UTF8
