<#PSScriptInfo

.VERSION 1.0.1

.GUID e479f80a-f3ce-4ff7-8cbc-c3807bc209f8

.AUTHOR Mike Hendrickson

#>

<#
.SYNOPSIS
 Gathers usage and consumption statistics on Exchange Database and Mailbox Usage.

.DESCRIPTION
 Gathers usage and consumption statistics on Exchange Database and Mailbox Usage.
 Statistics focus on categories like mailbox counts, mailbox quotas, mailbox sizes,
 and database sizes. Each time the script is run, the following information is written
 to disk:
    -Per Database statistics
    -Per Mailbox statistics
    -A database heat map, showing which databases are over or under provisioned
    -A mailbox move map, showing the required mailbox moves to reduce database heat

.PARAMETER Databases
 A list of the names of databases to get usage statistics for. If ommitted, Get-MailboxDatabase
 is used to discover all databases in the organization.

.PARAMETER MaxMailboxesPerDB
 The maximum number of standard or archive mailboxes allowed per database.
 Defaults to 62 (the default Exchange Server Role Requirements Calculator value).

.PARAMETER MaxQuotaPerDBinMB
 The maximum allowed defined quota, in Megabytes, of all standard and archive
 mailboxes per database.
 Defaults to 675180 (the default Exchange Server Role Requirements Calculator value; 62 * 10890).

.PARAMETER NormalMailboxCountPercentFloor
 Percentage is applied to MaxMailboxesPerDB, and used to calculate the bottom end of mailboxes
 per database that is considered normal. Defaults to 70.

.PARAMETER NormalMailboxCountPercentCeiling
 Percentage is applied to MaxMailboxesPerDB, and used to calculate the top end of mailboxes
 per database that is considered normal. Defaults to 90.

.PARAMETER NormalMailboxQuotaPercentFloor
 Percentage is applied to MaxQuotaPerDBinMB, and used to calculate the bottom end of defined quota
 per database that is considered normal. Defaults to 70.

.PARAMETER NormalMailboxQuotaPercentCeiling
 Percentage is applied to MaxQuotaPerDBinMB, and used to calculate the top end of defined quota
 per database that is considered normal. Defaults to 90.

.PARAMETER PrimaryQuotaProp
 The primary mailbox quota property used to determine defined quotas on standard mailboxes.
 Allowed values are IssueWarningQuota, ProhibitSendQuota, and ProhibitSendReceiveQuota.
 Defaults to ProhibitSendReceiveQutoa.

.PARAMETER SkipStandardMailboxes
 Whether the script should skip collecting and analyzing data from standard mailboxes. Defaults to $false.

.PARAMETER SkipMonitoringMailboxes
 Whether the script should skip collecting and analyzing data from monitoring mailboxes. Defaults to $true.

.PARAMETER SkipArbitrationMailboxes
 Whether the script should skip collecting and analyzing data from arbitration mailboxes. Defaults to $true.

.PARAMETER MoveMapMode
 Defines the behavior the script will follow when generating a mailbox move map. Options are:
 MoveHot: Moves mailboxes off databases that are considered hot.
 MoveWarm: Moves mailboxes off databases that are considered warm or worse.
 MoveNormal: Moves mailboxes off databases that are considered normal or worse.
 Defaults to MoveWarm.

.PARAMETER OutputPath
 The folder path to write usage reports to. Defaults to the Output subdirectory of the location where
 the script is being run.

.EXAMPLE
 > Gets Database Usage Stats for all Databases in the Organization
 PS> .\Get-ExchangeDatabaseUsage.ps1

.EXAMPLE
 > Gets Database Usage Stats for Databases DB1, DB2, and DB3. Outputs in Verbose mode.
 PS> .\Get-ExchangeDatabaseUsage.ps1 -Databases DB1,DB2,DB3 -Verbose

.INPUTS
 None. You cannot pipe objects to Get-ExchangeDatabaseUsage.ps1.

.OUTPUTS
 [PSObject]. Get-ExchangeDatabaseUsage.ps1 returns a custom PSObject containing various Database and Mailbox usage statistics.

#>

[CmdletBinding()]
param
(
    [String[]]
    $Databases,

	[Int]
	$MaxMailboxesPerDB = 62,

	[Int]
	$MaxQuotaPerDBinMB = 675180,

    [Int]
    $NormalMailboxCountPercentFloor = 70,

    [Int]
    $NormalMailboxCountPercentCeiling = 90,

    [Int]
    $NormalMailboxQuotaPercentFloor = 70,

    [Int]
    $NormalMailboxQuotaPercentCeiling = 90,

    [ValidateSet("IssueWarningQuota", "ProhibitSendQuota", "ProhibitSendReceiveQuota")]
    [String]
    $PrimaryQuotaProp = "ProhibitSendReceiveQuota",

    [Bool]
    $SkipStandardMailboxes = $false,

    [Bool]
    $SkipMonitoringMailboxes = $true,

    [Bool]
    $SkipArbitrationMailboxes = $true,

    [ValidateSet("MoveHot","MoveWarm","MoveNormal")]
    [String]
    $MoveMapMode = "MoveWarm",

    [String]
    $OutputPath = (Join-Path ((Get-Location).Path) "Output")
)

#Returns "Unknown" if the size object cannot be parsed
#Returns "Unlimited" if the size object is Unlimited
#Else returns the byte value corresponding to the size object
function Get-ByteValueFromSizeObject
{
    [CmdletBinding()]
    [OutputType([Object])]
    param
    (
		[parameter(Mandatory = $true)]
        [Microsoft.Exchange.Data.ByteQuantifiedSize]
        $SizeObject
    )

    [Object]$retValue = "Unknown"

    if ($null -ne $SizeObject)
    {
        if ($SizeObject.IsUnlimited -like "Unlimited*")
        {
            $retValue = "Unlimited"
        }
        elseif (($SizeObject | Get-Member | Where-Object {$_.Name -like "ToBytes"}).Count -gt 0)
        {
            $retValue = $SizeObject.ToBytes()
        }
    }

    return $retValue
}

function Get-ExchangeDatabaseAndMailboxData
{
    [CmdletBinding()]
    [OutputType([Object[]])]
    param
    (
		[parameter(Mandatory = $true)]
        [String[]]
        $Databases,

		[parameter(Mandatory = $true)]
        [String]
        $PrimaryQuotaProp,

		[parameter(Mandatory = $true)]
        [Bool]
        $SkipStandardMailboxes,

		[parameter(Mandatory = $true)]
        [Bool]
        $SkipMonitoringMailboxes,

		[parameter(Mandatory = $true)]
        [Bool]
        $SkipArbitrationMailboxes
    )

    [System.Collections.Generic.List[System.Object]]$allDBData = New-Object System.Collections.Generic.List[System.Object]

    for ($i = 1; $i -le $Databases.Count; $i++)
    {
        $database = $Databases[($i  - 1)]

        Write-Progress -Activity "Collecting Raw Database Data" -Status "Processing Database $($database) [$($i) / $($Databases.Count)]" -PercentComplete (($i * 100) / $Databases.Count)

        $dbProps = $null
        $stdMbxes = $null
        $arbMbxes = $null
        $monMbxes = $null

        Write-Verbose "$([DateTime]::Now): Getting properties for database: $($database)"

        $dbProps = Get-MailboxDatabase -Identity $database -Status -IncludePreExchange2013 -ErrorAction SilentlyContinue -Verbose:$false

        Write-Verbose "$([DateTime]::Now): Finished getting properties for database: $($database)"

        if ($null -eq $dbProps -or $null -eq $dbProps.DatabaseSize)
        {
            Write-Warning "$([DateTime]::Now): Failed to retrieve database $($database) with Get-MailboxDatabase. Skipping processing of this database."
            continue
        }

        $dbData = New-Object PSObject -Property @{Database                 = $database
                                                  DatabaseSize             = (Get-ByteValueFromSizeObject -SizeObject $dbProps.DatabaseSize)
                                                  IssueWarningQuota        = (Get-ByteValueFromSizeObject -SizeObject $dbProps.IssueWarningQuota)
                                                  ProhibitSendQuota        = (Get-ByteValueFromSizeObject -SizeObject $dbProps.ProhibitSendQuota)
                                                  ProhibitSendReceiveQuota = (Get-ByteValueFromSizeObject -SizeObject $dbProps.ProhibitSendReceiveQuota)
                                                  AllMailboxes             = $null
                                                  StandardMailboxes        = (New-Object System.Collections.Generic.List[System.Object])
                                                  ArchiveMailboxes        = (New-Object System.Collections.Generic.List[System.Object])
                                                  ArbitrationMailboxes     = (New-Object System.Collections.Generic.List[System.Object])
                                                  MonitoringMailboxes      = (New-Object System.Collections.Generic.List[System.Object])}    

        if (!$SkipStandardMailboxes)
        {
            Write-Verbose "$([DateTime]::Now): Getting standard mailboxes for database: $($dbData.Database)"
            $stdMbxes = Get-Mailbox -Database $dbData.Database -Verbose:$false
        }

        if (!$SkipArbitrationMailboxes)
        {
            Write-Verbose "$([DateTime]::Now): Getting arbitration mailboxes for database: $($dbData.Database)"
            $arbMbxes = Get-Mailbox -Database $dbData.Database -Arbitration -Verbose:$false
        }

        if (!$SkipMonitoringMailboxes)
        {
            Write-Verbose "$([DateTime]::Now): Getting monitoring mailboxes for database: $($dbData.Database)"
            $monMbxes = Get-Mailbox -Database $dbData.Database -Monitoring -Verbose:$false
        }

        Write-Verbose "$([DateTime]::Now): Finished getting mailboxes for database: $($dbData.Database)"

        foreach ($mbx in $stdMbxes)
        {
            $dbData.StandardMailboxes.Add((Get-MailboxData -DBData $dbData -MailboxObject $mbx -PrimaryQuotaProp $PrimaryQuotaProp -MailboxType "Standard"))

            if ($null -ne $mbx.ArchiveState -and $mbx.ArchiveState -like "Local") #This user has an on-prem archive. Let's count it.
            {
                $dbData.ArchiveMailboxes.Add((Get-MailboxData -DBData $dbData -MailboxObject $mbx -PrimaryQuotaProp $PrimaryQuotaProp -MailboxType "Archive"))
            }
        }

        foreach ($mbx in $arbMbxes)
        {
            $dbData.ArbitrationMailboxes.Add((Get-MailboxData -DBData $dbData -MailboxObject $mbx -PrimaryQuotaProp $PrimaryQuotaProp -MailboxType "Arbitration"))
        }

        foreach ($mbx in $monMbxes)
        {
            $dbData.MonitoringMailboxes.Add((Get-MailboxData -DBData $dbData -MailboxObject $mbx -PrimaryQuotaProp $PrimaryQuotaProp -MailboxType "Monitoring"))
        }

        $dbData.AllMailboxes = $dbData.StandardMailboxes + $dbData.ArchiveMailboxes + $dbData.ArbitrationMailboxes + $dbData.MonitoringMailboxes

        $allDBData.Add($dbData)
    }

    Write-Progress -Activity "Collecting Raw Database Data" -PercentComplete 100

    return $allDBData
}

function Get-MailboxData
{
    [CmdletBinding()]
    [OutputType([PSObject])]
    param
    (
		[parameter(Mandatory = $true)]
        [PSObject]
        $MailboxObject,

		[parameter(Mandatory = $true)]
        [PSObject]
        $DBData,

		[parameter(Mandatory = $true)]
        [String]
        $PrimaryQuotaProp,

        [ValidateSet("Standard", "Archive", "Arbitration", "Monitoring")]
        [String]
        $MailboxType
    )

    [PSObject]$mbInfo = New-Object PSObject -Property @{SmtpAddress                  = $MailboxObject.PrimarySmtpAddress
                                                        Database                     = $MailboxObject.Database.Name
                                                        UseDatabaseQuotaDefaults     = $MailboxObject.UseDatabaseQuotaDefaults
                                                        IssueWarningQuota            = $null
                                                        ProhibitSendQuota            = $null
                                                        ProhibitSendReceiveQuota     = $null
                                                        ItemCount                    = $null
                                                        DeletedItemCount             = $null
                                                        TotalItemSize                = $null
                                                        TotalDeletedItemSize         = $null
                                                        OverQuota                    = $null
                                                        LastLogonTime                = $null
                                                        MailboxType                  = $MailboxType}

    if ($MailboxType -notlike "Archive")
    {
        if ($MailboxObject.UseDatabaseQuotaDefaults -eq $true) #Get quotas from the DB
        {
            $mbInfo.IssueWarningQuota        = $DBData.IssueWarningQuota
            $mbInfo.ProhibitSendQuota        = $DBData.ProhibitSendQuota
            $mbInfo.ProhibitSendReceiveQuota = $DBData.ProhibitSendReceiveQuota             
        }
        else #Get quotas for this mailbox
        {
            $mbInfo.IssueWarningQuota        = (Get-ByteValueFromSizeObject -SizeObject $MailboxObject.IssueWarningQuota)
            $mbInfo.ProhibitSendQuota        = (Get-ByteValueFromSizeObject -SizeObject $MailboxObject.ProhibitSendQuota)
            $mbInfo.ProhibitSendReceiveQuota = (Get-ByteValueFromSizeObject -SizeObject $MailboxObject.ProhibitSendReceiveQuota)
        }
    }
    else
    {
        $mbInfo.IssueWarningQuota        = (Get-ByteValueFromSizeObject -SizeObject $MailboxObject.ArchiveWarningQuota)
        $mbInfo.ProhibitSendQuota        = (Get-ByteValueFromSizeObject -SizeObject $MailboxObject.ArchiveQuota)
        $mbInfo.ProhibitSendReceiveQuota = (Get-ByteValueFromSizeObject -SizeObject $MailboxObject.ArchiveQuota)
    }
    
    if ($MailboxType -notlike "Archive")
    {
        Write-Verbose "$([DateTime]::Now): Getting Mailbox Statistics for Mailbox: $($MailboxObject.DistinguishedName)"
        $mbxStats = $MailboxObject | Get-MailboxStatistics -WarningAction SilentlyContinue -Verbose:$false
    }
    else
    {
        Write-Verbose "$([DateTime]::Now): Getting Mailbox Statistics for Archive Mailbox: $($MailboxObject.DistinguishedName)"
        $mbxStats = $MailboxObject | Get-MailboxStatistics -Archive -WarningAction SilentlyContinue -Verbose:$false        
    }
   
    Write-Verbose "$([DateTime]::Now): Finished getting Mailbox Statistics for mailbox: $($MailboxObject.DistinguishedName)"

    if ($null -ne $mbxStats) #We got stats for the user
    {
        $mbInfo.ItemCount            = $mbxStats.ItemCount
        $mbInfo.DeletedItemCount     = $mbxStats.DeletedItemCount
        $mbInfo.TotalItemSize        = (Get-ByteValueFromSizeObject -SizeObject $mbxStats.TotalItemSize)
        $mbInfo.TotalDeletedItemSize = (Get-ByteValueFromSizeObject -SizeObject $mbxStats.TotalDeletedItemSize)

        #If the primary quota is not Unlimited or unknown, see if the mailbox size is greater than the quota
        if ($mbInfo."$($PrimaryQuotaProp)" -notlike "Unknown" -and $mbInfo."$($PrimaryQuotaProp)" -notlike "Unlimited" -and $mbInfo."$($PrimaryQuotaProp)" -lt $mbInfo.TotalItemSize)
        {
            $mbInfo.OverQuota = $true
        }
        else
        {
            $mbInfo.OverQuota = $false
        }

        $mbInfo.LastLogonTime = $mbxStats.LastLogonTime
    }
    else #This mailbox has probably never been logged into, so give them zero's for related stats
    {
        $mbInfo.ItemCount            = 0
        $mbInfo.DeletedItemCount     = 0
        $mbInfo.TotalItemSize        = 0
        $mbInfo.TotalDeletedItemSize = 0
        $mbInfo.OverQuota            = $false
        $mbInfo.LastLogonTime        = $null
    }

    return $mbInfo
}

function Get-ExchangeDatabaseUsageAnalysis
{
    [CmdletBinding()]
    [OutputType([PSObject])]
    param
    (
        [parameter(Mandatory = $true)]
        [PSObject]
        $ExchData,

        [parameter(Mandatory = $true)]
        [String]
        $PrimaryQuotaProp,

        [parameter(Mandatory = $true)]
        [PSObject]
        $UsageThresholds
    )

    Write-Verbose "$([DateTime]::Now): Analyzing data"

    [PSObject]$usageData = New-Object PSObject -Property @{RawData=$ExchData;Summaries=$null;HeatMap=$null}

    $usageData.Summaries = Get-DBSummaryCollection -UsageThresholds $UsageThresholds -DatabaseInfo $ExchData -PrimaryQuotaProp $PrimaryQuotaProp  
    $usageData.HeatMap   = Get-DatabaseHeatMap -DBSummaries $usageData.Summaries -UsageThresholds $UsageThresholds

    return $usageData
}

function Get-DBSummaryCollection
{
    [CmdletBinding()]
    [OutputType([Object[]])]
    param
    (
		[parameter(Mandatory = $true)]
        [PSObject[]]
        $DatabaseInfo,

		[parameter(Mandatory = $true)]
        [String]
        $PrimaryQuotaProp,

        [parameter(Mandatory = $true)]
        [PSObject]
        $UsageThresholds
    )

    [System.Collections.Generic.List[System.Object]]$dbSummaries = New-Object System.Collections.Generic.List[System.Object]

    for ($i = 1; $i -le $DatabaseInfo.Count; $i++)
    {
        $standardSummary = $null
        $archiveMailboxSummary = $null
        $arbitrationMailboxSummary = $null
        $monitoringMailboxSummary = $null

        $database = $DatabaseInfo[($i - 1)]

        Write-Progress -Activity "Summarizing Database Statistics" -Status "Processing Database $($database.Database) [$($i) / $($DatabaseInfo.Count)]"  -PercentComplete (($i * 100) / $DatabaseInfo.Count)

        #Get summaries for each mailbox type
        if ($database.StandardMailboxes.Count -gt 0)
        {
            $standardSummary           = Get-StatsForMailboxGroup -PropertyPrefix "Std" -MailboxInfo $database.StandardMailboxes -PrimaryQuotaProp $PrimaryQuotaProp
        }

        if ($database.ArchiveMailboxes.Count -gt 0)
        {
            $archiveMailboxSummary     = Get-StatsForMailboxGroup -PropertyPrefix "Arc" -MailboxInfo $database.ArchiveMailboxes -PrimaryQuotaProp $PrimaryQuotaProp
        }

        if ($database.ArbitrationMailboxes.Count -gt 0)
        {
            $arbitrationMailboxSummary = Get-StatsForMailboxGroup -PropertyPrefix "Arb" -MailboxInfo $database.ArbitrationMailboxes -PrimaryQuotaProp $PrimaryQuotaProp
        }

        if ($database.MonitoringMailboxes.Count -gt 0)
        {
            $monitoringMailboxSummary  = Get-StatsForMailboxGroup -PropertyPrefix "Mon" -MailboxInfo $database.MonitoringMailboxes -PrimaryQuotaProp $PrimaryQuotaProp
        }

        #Put together counts that will always show up
        $masterSummaryProps = @{
            Database                 = $database.Database
            Size                     = $database.DatabaseSize
            MaximumAllowedSize       = $UsageThresholds.MaxBytes
            FreeSpace                = ($UsageThresholds.MaxBytes - $database.DatabaseSize)
            PercentSpaceUtilized     = (($database.DatabaseSize * 100) / $UsageThresholds.MaxBytes)
            IssueWarningQuota        = $database.IssueWarningQuota
            ProhibitSendQuota        = $database.ProhibitSendQuota
            ProhibitSendReceiveQuota = $database.ProhibitSendReceiveQuota
        }        

        if ($null -ne $standardSummary)
        {
            foreach ($member in ($standardSummary | Get-Member | Where-Object {$_.MemberType -eq "NoteProperty"}))
            {
                $masterSummaryProps.Add($member.Name, $standardSummary."$($member.Name)")
            }

            $masterSummaryProps.Add("StdQuotaPercentAllocated", (($masterSummary.StdQuotaDefinedQuotaSum * 100) / $UsageThresholds.MaxBytes))
        }

        if ($null -ne $archiveMailboxSummary)
        {
            foreach ($member in ($archiveMailboxSummary | Get-Member | Where-Object {$_.MemberType -eq "NoteProperty"}))
            {
                $masterSummaryProps.Add($member.Name, $archiveMailboxSummary."$($member.Name)")
            }

            $masterSummaryProps.Add("ArcQuotaPercentAllocated", (($masterSummary.ArcQuotaDefinedQuotaSum * 100) / $UsageThresholds.MaxBytes))
        }

        if ($null -ne $arbitrationMailboxSummary)
        {
            foreach ($member in ($arbitrationMailboxSummary | Get-Member | Where-Object {$_.MemberType -eq "NoteProperty"}))
            {
                $masterSummaryProps.Add($member.Name, $arbitrationMailboxSummary."$($member.Name)")
            }

            $masterSummaryProps.Add("ArbQuotaPercentAllocated", (($masterSummary.ArbQuotaDefinedQuotaSum * 100) / $UsageThresholds.MaxBytes))
        }

        if ($null -ne $monitoringMailboxSummary)
        {
            foreach ($member in ($monitoringMailboxSummary | Get-Member | Where-Object {$_.MemberType -eq "NoteProperty"}))
            {
                $masterSummaryProps.Add($member.Name, $monitoringMailboxSummary."$($member.Name)")
            }

            $masterSummaryProps.Add("MonQuotaPercentAllocated", (($masterSummary.MonQuotaDefinedQuotaSum * 100) / $UsageThresholds.MaxBytes))
        }

        [PSObject]$masterSummary = New-Object PSObject -Property $masterSummaryProps

        $dbSummaries.Add($masterSummary)
    }

    Write-Progress -Activity "Summarizing Database Statistics" -PercentComplete 100

    return $dbSummaries
}

function Get-StatsForMailboxGroup
{
    [CmdletBinding()]
    [OutputType([PSObject])]
    param
    (
		[parameter(Mandatory = $true)]
        [PSObject[]]
        $MailboxInfo,

		[parameter(Mandatory = $true)]
        [String]
        $PrimaryQuotaProp,

		[parameter(Mandatory = $true)]
        [String]
        $PropertyPrefix
    )

    [UInt64]$defaultValue = 0

    $summary = New-Object PSObject -Property @{
                                        "$($PropertyPrefix)MailboxCount"              = $defaultValue
                                        "$($PropertyPrefix)ItemCountTotal"            = $defaultValue
                                        "$($PropertyPrefix)ItemCountAverage"          = $defaultValue
                                        "$($PropertyPrefix)DeletedItemCountTotal"     = $defaultValue
                                        "$($PropertyPrefix)DeletedItemCountAverage"   = $defaultValue
                                        "$($PropertyPrefix)ItemSizeTotal"             = $defaultValue
                                        "$($PropertyPrefix)ItemSizeAverage"           = $defaultValue
                                        "$($PropertyPrefix)DeletedItemSizeTotal"      = $defaultValue
                                        "$($PropertyPrefix)DeletedItemSizeAverage"    = $defaultValue
                                        "$($PropertyPrefix)LastLogonDaysOldest"       = $defaultValue
                                        "$($PropertyPrefix)LastLogonDaysNewest"       = $defaultValue     
                                        "$($PropertyPrefix)LastLogonDaysSum"          = $defaultValue 
                                        "$($PropertyPrefix)LastLogonDaysAverage"      = $defaultValue
                                        "$($PropertyPrefix)LastLogonDaysNotNullCount" = $defaultValue
                                        "$($PropertyPrefix)QuotaNonDefaultCount"      = $defaultValue
                                        "$($PropertyPrefix)QuotaOverQuotaCount"       = $defaultValue
                                        "$($PropertyPrefix)QuotaOverQuotaList"        = ""
                                        "$($PropertyPrefix)QuotaDefinedQuotaSum"      = $defaultValue
                                        "$($PropertyPrefix)QuotaUnlimitedCount"       = $defaultValue
                                        "$($PropertyPrefix)QuotaUnknownCount"         = $defaultValue
                                        }

    if ($MailboxInfo.Count -gt 0)
    {
        [Hashtable]$quotaHT = @{} 

        [TimeSpan]$totalTimeSinceLastLogon = 0
        $oldestLastLogon = 0
        $newestLastLogon = 0
        $setNewestLastLogon = $false

        [string[]]$overQuotaList = @()

        $now = [DateTime]::Now

        foreach ($mbx in $MailboxInfo)
        {
            ($summary."$($PropertyPrefix)MailboxCount")++

            $quotaValue = $mbx."$($PrimaryQuotaProp)"

            #Add to stats
            if (($quotaHT.ContainsKey($quotaValue)) -eq $true)
            {
                [int]$value = $quotaHT[$quotaValue]
                $value++
                $quotaHT[$mbx."$($PrimaryQuotaProp)"] = $value
            }
            else
            {
                $quotaHT.Add($mbx."$($PrimaryQuotaProp)", 1)
            }

            if ($quotaValue -like "Unknown")
            {
                ($summary."$($PropertyPrefix)QuotaUnknownCount")++
            }
            elseif ($quotaValue -like "Unlimited")
            {
                ($summary."$($PropertyPrefix)QuotaUnlimitedCount")++
            }
            else #This is a numerical value
            {
                ($summary."$($PropertyPrefix)QuotaDefinedQuotaSum") += $quotaValue
            }

            ($summary."$($PropertyPrefix)ItemCountTotal")        += $mbx.ItemCount
            ($summary."$($PropertyPrefix)DeletedItemCountTotal") += $mbx.DeletedItemCount
            ($summary."$($PropertyPrefix)ItemSizeTotal")         += $mbx.TotalItemSize
            ($summary."$($PropertyPrefix)DeletedItemSizeTotal")  += $mbx.TotalDeletedItemSize

            if ($mbx.UseDatabaseQuotaDefaults -eq $false)
            {
                ($summary."$($PropertyPrefix)QuotaNonDefaultCount")++
            }

            if ($mbx.OverQuota -eq $true)
            {
                ($summary."$($PropertyPrefix)QuotaOverQuotaCount")++

                $overQuotaList += $mbx.SmtpAddress
            }

            if ($null -ne $mbx.LastLogonTime -and $mbx.LastLogonTime.GetType().Name -eq "DateTime")
            {
                ($summary."$($PropertyPrefix)LastLogonDaysNotNullCount")++

                $timeSinceLastLogon       = $now - $mbx.LastLogonTime
                $totalTimeSinceLastLogon += $timeSinceLastLogon
                $daysSinceLastLogon       = $timeSinceLastLogon.Days

                if ($daysSinceLastLogon -gt $oldestLastLogon)
                {
                    $oldestLastLogon = $daysSinceLastLogon
                }

                if ($setNewestLastLogon -eq $false -or $daysSinceLastLogon -lt $newestLastLogon)
                {
                    $newestLastLogon    = $daysSinceLastLogon
                    $setNewestLastLogon = $true
                }
            }
        }

        $summary."$($PropertyPrefix)ItemCountAverage"        = $(if (($summary."$($PropertyPrefix)MailboxCount") -gt 0){($summary."$($PropertyPrefix)ItemCountTotal") / ($summary."$($PropertyPrefix)MailboxCount")} else{0})
        $summary."$($PropertyPrefix)DeletedItemCountAverage" = $(if (($summary."$($PropertyPrefix)MailboxCount") -gt 0){($summary."$($PropertyPrefix)DeletedItemCountTotal") / ($summary."$($PropertyPrefix)MailboxCount")} else{0})
        $summary."$($PropertyPrefix)ItemSizeAverage"         = $(if (($summary."$($PropertyPrefix)MailboxCount") -gt 0){($summary."$($PropertyPrefix)ItemSizeTotal") / ($summary."$($PropertyPrefix)MailboxCount")} else{0})
        $summary."$($PropertyPrefix)DeletedItemSizeAverage"  = $(if (($summary."$($PropertyPrefix)MailboxCount") -gt 0){($summary."$($PropertyPrefix)DeletedItemSizeTotal") / ($summary."$($PropertyPrefix)MailboxCount")} else{0})
        $summary."$($PropertyPrefix)LastLogonDaysOldest"     = $oldestLastLogon
        $summary."$($PropertyPrefix)LastLogonDaysNewest"     = $newestLastLogon   
        $summary."$($PropertyPrefix)LastLogonDaysSum"        = $totalTimeSinceLastLogon.Days
        $summary."$($PropertyPrefix)LastLogonDaysAverage"    = $(if (($summary."$($PropertyPrefix)LastLogonDaysNotNullCount") -gt 0){$totalTimeSinceLastLogon.Days / ($summary."$($PropertyPrefix)LastLogonDaysNotNullCount")} else{0}) 

        if ($overQuotaList.Count -gt 0)
        {
            $outStr = $overQuotaList[0]

            for ($i = 1; $i -lt $overQuotaList.Count; $i++)
            {
                $outStr += ("; " + $overQuotaList[$i])
            }

            $summary."$($PropertyPrefix)QuotaOverQuotaList" = $outStr
        }
    }

    return $summary
}

function Get-DatabaseHeatMap
{
    [CmdletBinding()]
    [OutputType([PSObject])]
    param
    (
        [parameter(Mandatory = $true)]
        $DBSummaries,

        [parameter(Mandatory = $true)]
        [PSObject]
        $UsageThresholds
    )

    [PSObject]$heatMap = New-Object PSObject -Property @{            
            HealthOverall                 = "Cold"
            HealthByCount                 = "Cold"
            HealthByDefinedQuota          = "Cold"
            HealthByItemSize              = "Cold"
            HealthByFileSize              = "Cold"

            AllDatabases                  = $DBSummaries

            HotDatabasesByCount           = ($DBSummaries | Where-Object {$_.StdMailboxCount -gt $UsageThresholds.MaxMailboxCount})
            WarmDatabasesByCount          = ($DBSummaries | Where-Object {$_.StdMailboxCount -ge $UsageThresholds.TooManyMailboxesCount -and $_.StdMailboxCount -lt $UsageThresholds.MaxMailboxCount})
            NormalDatabasesByCount        = ($DBSummaries | Where-Object {$_.StdMailboxCount -gt $UsageThresholds.TooFewMailboxesCount -and $_.StdMailboxCount -lt $UsageThresholds.TooManyMailboxesCount})
            CoolDatabasesByCount          = ($DBSummaries | Where-Object {$_.StdMailboxCount -le $UsageThresholds.TooFewMailboxesCount})

            HotDatabasesByDefinedQuota    = ($DBSummaries | Where-Object {$_.StdQuotaDefinedQuotaSum -gt $UsageThresholds.MaxBytes})
            WarmDatabasesByDefinedQuota   = ($DBSummaries | Where-Object {$_.StdQuotaDefinedQuotaSum -ge $UsageThresholds.TooManyBytes -and $_.StdQuotaDefinedQuotaSum -lt $UsageThresholds.MaxBytes})
            NormalDatabasesByDefinedQuota = ($DBSummaries | Where-Object {$_.StdQuotaDefinedQuotaSum -gt $UsageThresholds.TooFewBytes -and $_.StdQuotaDefinedQuotaSum -lt $UsageThresholds.TooManyBytes})
            CoolDatabasesByDefinedQuota   = ($DBSummaries | Where-Object {$_.StdQuotaDefinedQuotaSum -le $UsageThresholds.TooFewBytes})

            HotDatabasesByItemSize        = ($DBSummaries | Where-Object {$_.StdItemSizeTotal -gt $UsageThresholds.MaxBytes})
            WarmDatabasesByItemSize       = ($DBSummaries | Where-Object {$_.StdItemSizeTotal -ge $UsageThresholds.TooManyBytes -and $_.Size -lt $UsageThresholds.MaxBytes})
            NormalDatabasesByItemSize     = ($DBSummaries | Where-Object {$_.StdItemSizeTotal -gt $UsageThresholds.TooFewBytes -and $_.Size -lt $UsageThresholds.TooManyBytes})
            CoolDatabasesByItemSize       = ($DBSummaries | Where-Object {$_.StdItemSizeTotal -le $UsageThresholds.TooFewBytes})

            HotDatabasesByFileSize        = ($DBSummaries | Where-Object {$_.Size -gt $UsageThresholds.MaxBytes})
            WarmDatabasesByFileSize       = ($DBSummaries | Where-Object {$_.Size -ge $UsageThresholds.TooManyBytes -and $_.Size -lt $UsageThresholds.MaxBytes})
            NormalDatabasesByFileSize     = ($DBSummaries | Where-Object {$_.Size -gt $UsageThresholds.TooFewBytes -and $_.Size -lt $UsageThresholds.TooManyBytes})
            CoolDatabasesByFileSize       = ($DBSummaries | Where-Object {$_.Size -le $UsageThresholds.TooFewBytes})
            
            ReportHtml                    = $null}

    $heatMap.HealthOverall        = Get-HeatFromHeatCount -HotCount (([Object[]]$heatMap.HotDatabasesByCount).Count -gt 0 -or ([Object[]]$heatMap.HotDatabasesByDefinedQuota).Count -gt 0 -or ([Object[]]$heatMap.HotDatabasesByFileSize).Count -gt 0) -WarmCount (([Object[]]$heatMap.WarmDatabasesByCount).Count -gt 0 -or ([Object[]]$heatMap.WarmDatabasesByDefinedQuota).Count -gt 0 -or ([Object[]]$heatMap.WarmDatabasesByFileSize).Count -gt 0) -NormalCount (([Object[]]$heatMap.NormalDatabasesByCount).Count -gt 0 -or ([Object[]]$heatMap.NormalDatabasesByDefinedQuota).Count -gt 0 -or ([Object[]]$heatMap.NormalDatabasesByFileSize).Count -gt 0)
    $heatMap.HealthByCount        = Get-HeatFromHeatCount -HotCount ([Object[]]$heatMap.HotDatabasesByCount).Count -WarmCount ([Object[]]$heatMap.WarmDatabasesByCount).Count -NormalCount ([Object[]]$heatMap.NormalDatabasesByCount).Count
    $heatMap.HealthByDefinedQuota = Get-HeatFromHeatCount -HotCount ([Object[]]$heatMap.HotDatabasesByDefinedQuota).Count -WarmCount ([Object[]]$heatMap.WarmDatabasesByDefinedQuota).Count -NormalCount ([Object[]]$heatMap.NormalDatabasesByDefinedQuota).Count
    $heatMap.HealthByItemSize     = Get-HeatFromHeatCount -HotCount ([Object[]]$heatMap.HotDatabasesByItemSize).Count -WarmCount ([Object[]]$heatMap.WarmDatabasesByItemSize).Count -NormalCount ([Object[]]$heatMap.NormalDatabasesByItemSize).Count
    $heatMap.HealthByFileSize     = Get-HeatFromHeatCount -HotCount ([Object[]]$heatMap.HotDatabasesByFileSize).Count -WarmCount ([Object[]]$heatMap.WarmDatabasesByFileSize).Count -NormalCount ([Object[]]$heatMap.NormalDatabasesByFileSize).Count    
    $heatMap.ReportHtml           = Get-DatabaseHeatReportHtml -HeatMap $heatMap -UsageThresholds $UsageThresholds

    return $heatMap
}

function Get-ColorForHeat
{
    [CmdletBinding()]    
    [OutputType([String])]
    param
    (
		[parameter(Mandatory = $true)]
        [ValidateSet("Hot", "Warm", "Normal", "Cold")]
        [String]
        $Heat        
    )

    if ($Heat -like "Hot")
    {
        return "red"
    }
    elseif ($Heat -like "Warm")
    {
        return "orange"
    }
    elseif ($Heat -like "Normal")
    {
        return "lime"
    }
    else
    {
        return "cyan"
    }
}

function Get-HeatFromHeatCount
{
    [CmdletBinding()]
    [OutputType([String])]
    param
    (		
        [Int]
        $HotCount = 0,

        [Int]
        $WarmCount = 0,

        [Int]
        $NormalCount = 0
    )

    if ($HotCount -gt 0)
    {
        return "Hot"
    }
    elseif ($WarmCount -gt 0)
    {
        return "Warm"
    }
    elseif ($NormalCount -gt 0)
    {
        return "Normal"
    }
    else
    {
        return "Cold"
    }
}

function Get-DatabaseHeatReportHtml
{
    [CmdletBinding()]
    [OutputType([String])]
    param
    (
        [parameter(Mandatory = $true)]
        [PSObject]
        $HeatMap,

        [parameter(Mandatory = $true)]
        [PSObject]
        $UsageThresholds
    )

    $overallHeatColor = Get-ColorForHeat -Heat $HeatMap.HealthOverall

    $reportBuilder = New-Object -TypeName System.Text.StringBuilder

    [Void]$reportBuilder.Append("<!DOCTYPE html><html lang=`"en`" xmlns=`"http://www.w3.org/1999/xhtml`"><head><meta charset=`"utf-8`" /><title>Exchange Database Heat Report</title></head><body>")

    [Void]$reportBuilder.Append("<span style=`"color:$overallHeatColor`"><strong>Environment Heat: $($HeatMap.HealthOverall)</strong></span><br/><br/>")

    #Start the heat map table
    [Void]$reportBuilder.Append("<table style=`"color:black; border-collapse: collapse; border: 1px solid black; padding: 5px`">")

    #Do the table headers
    [Void]$reportBuilder.Append("<tr>")
    [Void]$reportBuilder.Append("<th style=`"color:black; border-collapse: collapse; border: 1px solid black; padding: 5px`">Database</th>")
    [Void]$reportBuilder.Append("<th style=`"background-color:$overallHeatColor; border-collapse: collapse; border: 1px solid black; padding: 5px`">Overall Heat</th>")
    [Void]$reportBuilder.Append("<th style=`"background-color:$(Get-ColorForHeat -Heat $HeatMap.HealthByCount); border-collapse: collapse; border: 1px solid black; padding: 5px`">Mailbox Count</th>")
    [Void]$reportBuilder.Append("<th style=`"background-color:$(Get-ColorForHeat -Heat $HeatMap.HealthByDefinedQuota); border-collapse: collapse; border: 1px solid black; padding: 5px`">Defined Quota Sum (GB)</th>")
    [Void]$reportBuilder.Append("<th style=`"background-color:$(Get-ColorForHeat -Heat $HeatMap.HealthByItemSize); border-collapse: collapse; border: 1px solid black; padding: 5px`">Item Size (GB)</th>")    
    [Void]$reportBuilder.Append("<th style=`"background-color:$(Get-ColorForHeat -Heat $HeatMap.HealthByFileSize); border-collapse: collapse; border: 1px solid black; padding: 5px`">File Size (GB)</th>")    
    [Void]$reportBuilder.Append("</tr>")

    #Do the table rows
    foreach ($database in $HeatMap.AllDatabases)
    {
        
        $dbHeat = Get-HeatFromHeatCount -HotCount (([Object[]]($HeatMap.HotDatabasesByCount | Where-Object {$_.Database -like $database.Database})).Count + ([Object[]]($HeatMap.HotDatabasesByDefinedQuota | Where-Object {$_.Database -like $database.Database})).Count + ([Object[]]($HeatMap.HotDatabasesByFileSize | Where-Object {$_.Database -like $database.Database})).Count) -WarmCount (([Object[]]($HeatMap.WarmDatabasesByCount | Where-Object {$_.Database -like $database.Database})).Count + ([Object[]]($HeatMap.WarmDatabasesByDefinedQuota | Where-Object {$_.Database -like $database.Database})).Count + ([Object[]]($HeatMap.WarmDatabasesByFileSize | Where-Object {$_.Database -like $database.Database})).Count) -NormalCount (([Object[]]($HeatMap.NormalDatabasesByCount | Where-Object {$_.Database -like $database.Database})).Count + ([Object[]]($HeatMap.NormalDatabasesByDefinedQuota | Where-Object {$_.Database -like $database.Database})).Count + ([Object[]]($HeatMap.NormalDatabasesByFileSize | Where-Object {$_.Database -like $database.Database})).Count)
        $dbHeatColor = Get-ColorForHeat -Heat $dbHeat

        [Void]$reportBuilder.Append("<tr>")
        [Void]$reportBuilder.Append("<td style=`"background-color:$dbHeatColor; border-collapse: collapse; border: 1px solid black; padding: 5px`">$($database.Database)</td>")
        [Void]$reportBuilder.Append("<td style=`"background-color:$dbHeatColor; border-collapse: collapse; border: 1px solid black; padding: 5px`">$dbHeat</td>")
        [Void]$reportBuilder.Append("<td style=`"background-color:$(Get-ColorForHeat -Heat (Get-HeatFromHeatCount -HotCount ([Object[]]($HeatMap.HotDatabasesByCount | Where-Object {$_.Database -like $database.Database})).Count -WarmCount ([Object[]]($HeatMap.WarmDatabasesByCount | Where-Object {$_.Database -like $database.Database})).Count -NormalCount ([Object[]]($HeatMap.NormalDatabasesByCount | Where-Object {$_.Database -like $database.Database})).Count)); border-collapse: collapse; border: 1px solid black; padding: 5px`">$($database.StdMailboxCount)</td>")
        [Void]$reportBuilder.Append("<td style=`"background-color:$(Get-ColorForHeat -Heat (Get-HeatFromHeatCount -HotCount ([Object[]]($HeatMap.HotDatabasesByDefinedQuota | Where-Object {$_.Database -like $database.Database})).Count -WarmCount ([Object[]]($HeatMap.WarmDatabasesByDefinedQuota | Where-Object {$_.Database -like $database.Database})).Count -NormalCount ([Object[]]($HeatMap.NormalDatabasesByDefinedQuota | Where-Object {$_.Database -like $database.Database})).Count)); border-collapse: collapse; border: 1px solid black; padding: 5px`">$(([Math]::Round(($database.StdQuotaDefinedQuotaSum / 1024 / 1024 / 1024), 2)))</td>")
        [Void]$reportBuilder.Append("<td style=`"background-color:$(Get-ColorForHeat -Heat (Get-HeatFromHeatCount -HotCount ([Object[]]($HeatMap.HotDatabasesByItemSize | Where-Object {$_.Database -like $database.Database})).Count -WarmCount ([Object[]]($HeatMap.WarmDatabasesByItemSize | Where-Object {$_.Database -like $database.Database})).Count -NormalCount ([Object[]]($HeatMap.NormalDatabasesByItemSize | Where-Object {$_.Database -like $database.Database})).Count)); border-collapse: collapse; border: 1px solid black; padding: 5px`">$(([Math]::Round(($database.Size / 1024 / 1024 / 1024), 2)))</td>")        
        [Void]$reportBuilder.Append("<td style=`"background-color:$(Get-ColorForHeat -Heat (Get-HeatFromHeatCount -HotCount ([Object[]]($HeatMap.HotDatabasesByFileSize | Where-Object {$_.Database -like $database.Database})).Count -WarmCount ([Object[]]($HeatMap.WarmDatabasesByFileSize | Where-Object {$_.Database -like $database.Database})).Count -NormalCount ([Object[]]($HeatMap.NormalDatabasesByFileSize | Where-Object {$_.Database -like $database.Database})).Count)); border-collapse: collapse; border: 1px solid black; padding: 5px`">$(([Math]::Round(($database.Size / 1024 / 1024 / 1024), 2)))</td>")        
        [Void]$reportBuilder.Append("</tr>")
    }

    [Void]$reportBuilder.Append("</table><br/><br/>")

    #Create heat map legend
    [Void]$reportBuilder.Append("<span style=`"color:black`"><strong>Legend<strong></span><br/>")
    [Void]$reportBuilder.Append("<table style=`"color:black; border-collapse: collapse; border: 1px solid black; padding: 5px`">")
    [Void]$reportBuilder.Append("<tr><th style=`"color:black; border-collapse: collapse; border: 1px solid black; padding: 5px`">Limit</th><th style=`"color:black; border-collapse: collapse; border: 1px solid black; padding: 5px`">Cold</th><th style=`"color:black; border-collapse: collapse; border: 1px solid black; padding: 5px`">Normal</th><th style=`"color:black; border-collapse: collapse; border: 1px solid black; padding: 5px`">Warm</th><th style=`"color:black; border-collapse: collapse; border: 1px solid black; padding: 5px`">Hot</th></tr>")
    [Void]$reportBuilder.Append("<tr><td style=`"color:black; border-collapse: collapse; border: 1px solid black; padding: 5px`">Mailbox Count</td><td style=`"color:black; border-collapse: collapse; border: 1px solid black; padding: 5px`">[0, $($UsageThresholds.TooFewMailboxesCount)]</td><td style=`"color:black; border-collapse: collapse; border: 1px solid black; padding: 5px`">($($UsageThresholds.TooFewMailboxesCount),$($UsageThresholds.TooManyMailboxesCount))</td><td style=`"color:black; border-collapse: collapse; border: 1px solid black; padding: 5px`">[$($UsageThresholds.TooManyMailboxesCount),$($UsageThresholds.MaxMailboxCount)]</td><td style=`"color:black; border-collapse: collapse; border: 1px solid black; padding: 5px`">($($UsageThresholds.MaxMailboxCount),*)</td></tr>")
    [Void]$reportBuilder.Append("<tr><td style=`"color:black; border-collapse: collapse; border: 1px solid black; padding: 5px`">Size (GB)</td><td style=`"color:black; border-collapse: collapse; border: 1px solid black; padding: 5px`">[0, $($UsageThresholds.TooFewGB)]</td><td style=`"color:black; border-collapse: collapse; border: 1px solid black; padding: 5px`">($($UsageThresholds.TooFewGB),$($UsageThresholds.TooManyGB))</td><td style=`"color:black; border-collapse: collapse; border: 1px solid black; padding: 5px`">[$($UsageThresholds.TooManyGB),$($UsageThresholds.MaxGB)]</td><td style=`"color:black; border-collapse: collapse; border: 1px solid black; padding: 5px`">($($UsageThresholds.MaxGB),*)</td></tr>")
    [Void]$reportBuilder.Append("</table><br/><br/>")

    #Finish the body
    [Void]$reportBuilder.Append("</body></html>")

    return $reportBuilder.ToString()
}

function Test-IsDBOverQuota
{
    [CmdletBinding()]
    [OutputType([Bool])]
    param
    (
        [parameter(Mandatory = $true)]
        [PSObject]
        $DatabaseSummary,

        [parameter(Mandatory = $true)]
        [Int]
        $MaxMailboxCount,
        
        [Int]
        $ExtraMailboxCount = 0,
        
        [parameter(Mandatory = $true)]
        [UInt64]
        $MaxQuotaBytes,

        [UInt64]
        $ExtraQuotaBytes = 0,

        [parameter(Mandatory = $true)]
        [UInt64]
        $MaxSizeBytes,

        [UInt64]
        $ExtraSizeBytes = 0
    )

    return (($DatabaseSummary.StdMailboxCount + $DatabaseSummary.ArcMailboxCount + $ExtraMailboxCount) -ge $MaxMailboxCount -or ($DatabaseSummary.StdQuotaDefinedQuotaSum + $DatabaseSummary.ArcQuotaDefinedQuotaSum + $ExtraQuotaBytes) -ge $MaxQuotaBytes -or ($DatabaseSummary.StdItemSizeTotal + $DatabaseSummary.ArcItemSizeTotal + $ExtraSizeBytes) -ge $MaxSizeBytes)
}

function Get-MoveMap
{
    [CmdletBinding()]
    [OutputType([PSObject])]
    param
    (
        [parameter(Mandatory = $true)]
        [PSObject]
        $HeatMap,

        [parameter(Mandatory = $true)]
        [PSObject]
        $MailboxData,

        [parameter(Mandatory = $true)]
        [PSObject]
        $UsageThresholds,

		[parameter(Mandatory = $true)]
        [ValidateSet("IssueWarningQuota", "ProhibitSendQuota", "ProhibitSendReceiveQuota")]
        [String]
        $PrimaryQuotaProp,

		[parameter(Mandatory = $true)]
        [ValidateSet("MoveHot","MoveWarm","MoveNormal")]
        [String]
        $MoveMapMode
    )

    [System.Collections.Generic.List[PSObject]]$moveMap = New-Object System.Collections.Generic.List[PSObject]

    if (([Object[]]$HeatMap.AllDatabases).Count -le 1)
    {
        Write-Verbose "$([DateTime]::Now): Not enough databases to calculate move map"
        return $moveMap
    }

    Write-Verbose "$([DateTime]::Now): Calculating move map"

    if ($MoveMapMode -like "MoveHot")
    {
		[Int]$maxMailboxCount = $UsageThresholds.MaxMailboxCount + 1
		[UInt64]$maxDefinedQuota = $UsageThresholds.MaxBytes + 1
		[UInt64]$maxItemSize = $UsageThresholds.MaxBytes + 1
    }
    elseif ($MoveMapMode -like "MoveWarm")
    {
		[Int]$maxMailboxCount = $UsageThresholds.TooManyMailboxesCount
		[UInt64]$maxDefinedQuota = $UsageThresholds.TooManyBytes
		[UInt64]$maxItemSize = $UsageThresholds.TooManyBytes
    }
    elseif ($MoveMapMode -like "MoveNormal")
    {
		[Int]$maxMailboxCount = $UsageThresholds.TooFewMailboxesCount + 1
		[UInt64]$maxDefinedQuota = $UsageThresholds.TooManyBytes + 1
		[UInt64]$maxItemSize = $UsageThresholds.TooFewBytes + 1
    }

    [System.Collections.Generic.List[PSObject]]$dbBeforeStats = New-Object System.Collections.Generic.List[PSObject]
    $addedBeforeStats = $false
    $passNumber = 0

    do
    {
        $passNumber++
        Write-Verbose "$([DateTime]::Now): Beginning move map pass #$($passNumber)"

        $madeMoves = $false

        #Get the source databases
        [PSObject[]]$dbSources = $HeatMap.AllDatabases | Where-Object {(Test-IsDBOverQuota -DatabaseSummary $_ -MaxMailboxCount $maxMailboxCount -MaxQuotaBytes $maxDefinedQuota -MaxSizeBytes $maxItemSize)}
 
        if ($dbSources.Count -eq 0)
        {
            Write-Warning "$([DateTime]::Now): Unable to determine source databases for move map."
            return
        }

        #Get a modifiable list of source mailboxes
        [System.Collections.Generic.List[PSObject]]$mbxSources = New-Object System.Collections.Generic.List[PSObject]

        foreach ($sourceDB in $dbSources)
        {
            if (!$addedBeforeStats)
            {
                $dbBeforeStats.Add((New-Object PSObject -Property @{Database=$sourceDB.Database;MailboxCount=($sourceDB.StdMailboxCount + $sourceDB.ArcMailboxCount);DefinedQuotaSum=($sourceDB.StdQuotaDefinedQuotaSum + $sourceDB.ArcQuotaDefinedQuotaSum);ItemSizeTotal=($sourceDB.StdItemSizeTotal + $sourceDB.ArcItemSizeTotal)}))
            }

            foreach ($mbx in $MailboxData | Where-Object {$_.Database -like $sourceDB.Database} | Sort-Object -Property @{Expression={$_.$PrimaryQuotaProp};Descending=$true}, @{Expression="TotalItemSize";Descending=$true})
            {
                if (([String[]]$moveMap.SmtpAddress).Count -eq 0 -or (([String[]]$moveMap.SmtpAddress).Contains($mbx.SmtpAddress)) -eq $false)
                {
                    $mbxSources.Add($mbx)
                }
            }
        }

        #Get a modifiable list of target databases
        [System.Collections.Generic.List[PSObject]]$dbTargets = New-Object System.Collections.Generic.List[PSObject]

        foreach ($db in $HeatMap.AllDatabases | Where-Object {!(Test-IsDBOverQuota -DatabaseSummary $_ -MaxMailboxCount $maxMailboxCount -MaxQuotaBytes $maxDefinedQuota -MaxSizeBytes $maxItemSize)} | Sort-Object -Property @{Expression={($_.StdQuotaDefinedQuotaSum + $_.ArcQuotaDefinedQuotaSum)};Descending=$false}, @{Expression={($_.StdItemSizeTotal + $_.ArcItemSizeTotal)};Descending=$false})
        {
            if (!$addedBeforeStats)
            {
                $dbBeforeStats.Add((New-Object PSObject -Property @{Database=$db.Database;MailboxCount=($db.StdMailboxCount + $db.ArcMailboxCount);DefinedQuotaSum=($db.StdQuotaDefinedQuotaSum + $db.ArcQuotaDefinedQuotaSum);ItemSizeTotal=($db.StdItemSizeTotal + $db.ArcItemSizeTotal)}))
            }

            $dbTargets.Add($db)
        }

        $addedBeforeStats = $true

        Write-Verbose "$([DateTime]::Now): Attempting to add up to $($mbxSources.Count) mailboxes to move map."

        while ($mbxSources.Count -gt 0 -and $dbTargets.Count -gt 0)
        {
		    #Get the next mailbox and remove from list
            $currentMailbox = $mbxSources[0]
            $mbxSources.RemoveAt(0)

		    #Get the source DB
            $sourceDB = $HeatMap.AllDatabases | Where-Object {$_.Database -like $currentMailbox.Database}
      
            #Get the target DB
            $targetDB = $null

            #See if this will put the database over the quota
            while($null -eq $targetDB -and $dbTargets.Count -gt 0)
            {
                $targetDB = $dbTargets[0]

                if ((Test-IsDBOverQuota -DatabaseSummary $targetDB -MaxMailboxCount $maxMailboxCount -MaxQuotaBytes $maxDefinedQuota -MaxSizeBytes $maxItemSize -ExtraMailboxCount 1 -ExtraQuotaBytes $currentMailbox.$PrimaryQuotaProp -ExtraSizeBytes $currentMailbox.TotalItemSize))
                {
                    Write-Verbose "$([DateTime]::Now): Target database would go over quota with more mailboxes. Database: $($targetDB.Database)."
                    $targetDB = $null
                    $dbTargets.RemoveAt(0)
                }
            }

            #Couldn't find a DB target, so break out of loop.
            if ($null -eq $targetDB)
            {
                break;
            }

            #Add this mailbox to target DB totals
            if ($currentMailbox.MailboxType -notlike "Archive")
            {
                $sourceDB.StdMailboxCount--
                $sourceDB.StdQuotaDefinedQuotaSum -= $currentMailbox.$PrimaryQuotaProp
                $sourceDB.StdItemSizeTotal -= $currentMailbox.TotalItemSize
                $targetDB.StdMailboxCount++
                $targetDB.StdQuotaDefinedQuotaSum += $currentMailbox.$PrimaryQuotaProp
                $targetDB.StdItemSizeTotal += $currentMailbox.TotalItemSize
            }
            else
            {
                $sourceDB.ArcMailboxCount--
                $sourceDB.ArcQuotaDefinedQuotaSum -= $currentMailbox.$PrimaryQuotaProp
                $sourceDB.ArcItemSizeTotal -= $currentMailbox.TotalItemSize
                $targetDB.ArcMailboxCount++
                $targetDB.ArcQuotaDefinedQuotaSum += $currentMailbox.$PrimaryQuotaProp
                $targetDB.ArcItemSizeTotal += $currentMailbox.TotalItemSize
            }

            #Add to move map
            $moveMap.Add((New-Object PSObject -Property @{SmtpAddress=$currentMailbox.SmtpAddress;SourceDB=$sourceDB.Database;TargetDB=$targetDB.Database;IsArchive=($currentMailbox.MailboxType -like "Archive")}))
            $madeMoves = $true

            Write-Verbose "$([DateTime]::Now): Move $($currentMailbox.SmtpAddress) from $($sourceDB.Database) to $($targetDB.Database). Mailbox Quota: $($currentMailbox.$PrimaryQuotaProp). Mailbox Size: $($currentMailbox.TotalItemSize). IsArchive: $(($currentMailbox.MailboxType -like "Archive"))"

            #Check if this DB is still over provisioned
            if (!(Test-IsDBOverQuota -DatabaseSummary $sourceDB -MaxMailboxCount $maxMailboxCount -MaxQuotaBytes $maxDefinedQuota -MaxSizeBytes $maxItemSize))
            {
                Write-Verbose "$([DateTime]::Now): Database $($sourceDB.Database) is no longer over provisioned."
                $mbxSources = $mbxSources | Where-Object {$_.Database -notlike $sourceDB.Database}
            }

            #If our target has gone over quota or mailbox count, go to the next database
            if ((Test-IsDBOverQuota -DatabaseSummary $targetDB -MaxMailboxCount $maxMailboxCount -MaxQuotaBytes $maxDefinedQuota -MaxSizeBytes $maxItemSize))
            {
                Write-Verbose "$([DateTime]::Now): Target database has reached limit. Database: $($targetDB.Database)."
                $dbTargets.RemoveAt(0)
            }
        }
    } while ($madeMoves -and $mbxSources.Count -gt 0)
    
    if ($mbxSources.Count -gt 0 -and $dbTargets.Count -eq 0)
    {
        Write-Warning "$([DateTime]::Now): Not enough target databases to finish moves"
    }

    Write-Verbose "$([DateTime]::Now): Databases Before and After Moves:"

    foreach ($dbBefore in $dbBeforeStats | Sort-Object -Property Database)
    {
        $dbAfter = $HeatMap.AllDatabases | Where-Object {$_.Database -like $dbBefore.Database}

        Write-Verbose "Database: $($dbAfter.Database). Count Before: $($dbBefore.MailboxCount). Count After:  $(($dbAfter.StdMailboxCount + $dbAfter.ArcMailboxCount)). Quota Before: $($dbBefore.DefinedQuotaSum). Quota After: $(($dbAfter.StdQuotaDefinedQuotaSum + $dbAfter.ArcQuotaDefinedQuotaSum)). Item Size Before: $($dbBefore.ItemSizeTotal). Item Size After: $(($dbAfter.StdItemSizeTotal + $dbAfter.ArcItemSizeTotal))."
    }
    

    return $moveMap
}

function Write-ExchangeDatabaseUsage
{
    [CmdletBinding()]
    [OutputType([Object])]
    param
    (
        [parameter(Mandatory = $true)]
        [PSObject]
        $UsageData,

        [PSObject]
        $MoveMap,

        [parameter(Mandatory = $true)]
        [String]
        $OutputPath
    )

    Write-Verbose "$([DateTime]::Now): Saving results to disk"

    $now = [DateTime]::Now
    $timeString = "$($now.Year)$($now.Month)$($now.Day)$($now.Hour)$($now.Minute)$($now.Second)"

    $usageData.RawData.AllMailboxes | Export-Csv (Join-Path $OutputPath "MailboxStats-$($timeString).csv") -NoTypeInformation

    $usageData.Summaries | Export-Csv (Join-Path $OutputPath "DatabaseStats-$($timeString).csv") -NoTypeInformation

    $usageData.HeatMap.ReportHtml | Out-File (Join-Path $OutputPath "DatabaseHeatReport-$($timeString).html")

    if ($null -ne $MoveMap)
    {
        $moveMap | Export-Csv (Join-Path $OutputPath "MoveMap-$($timeString).csv") -NoTypeInformation
    }
}

####################################################################################################
# Script execution starts here
####################################################################################################

if ($SkipStandardMailboxes -and $SkipMonitoringMailboxes -and $SkipArbitrationMailboxes)
{
    return
}

#Make the log output directory if it doesnt' exist
if (!(Test-Path -Path $OutputPath))
{
    mkdir -Path $OutputPath
}

#Calculate the max bytes and max mailboxes per DB
[UInt64]$maxBytesPerDB = $MaxQuotaPerDBinMB * 1024 * 1024

[PSObject]$usageThresholds = New-Object PSObject -Property @{MaxMailboxCount       = $MaxMailboxesPerDB
                                                             MaxBytes              = $maxBytesPerDB
                                                             MaxGB                 = ([Math]::Round(($maxBytesPerDB / 1024 / 1024 / 1024), 2))
                                                             TooFewMailboxesCount  = ([Int]($maxMailboxesPerDB * $NormalMailboxCountPercentFloor * .01))
                                                             TooManyMailboxesCount = ([Int]($maxMailboxesPerDB * $NormalMailboxCountPercentCeiling * .01))
                                                             TooFewBytes           = ([UInt64]($maxBytesPerDB * $NormalMailboxQuotaPercentFloor * .01))
                                                             TooFewGB              = ([Math]::Round(($maxBytesPerDB * $NormalMailboxQuotaPercentFloor * .01 / 1024 / 1024 / 1024), 2))
                                                             TooManyBytes          = ($maxBytesPerDB * $NormalMailboxQuotaPercentCeiling * .01)
                                                             TooManyGB             = ([Math]::Round(($maxBytesPerDB * $NormalMailboxQuotaPercentCeiling * .01 / 1024 / 1024 / 1024), 2))}

if ($Databases.Count -eq 0)
{
    Write-Verbose "$([DateTime]::Now): No value was passed for Databases. Discovering all Databases in the Organization."

    $Databases = (Get-MailboxDatabase).Name
}

$exchData = Get-ExchangeDatabaseAndMailboxData -Databases $Databases -PrimaryQuotaProp $PrimaryQuotaProp -SkipStandardMailboxes $SkipStandardMailboxes -SkipMonitoringMailboxes $SkipMonitoringMailboxes -SkipArbitrationMailboxes $SkipArbitrationMailboxes
$usageData = Get-ExchangeDatabaseUsageAnalysis -ExchData $exchData -PrimaryQuotaProp $PrimaryQuotaProp -UsageThresholds $usageThresholds -Verbose
$moveMap = Get-MoveMap -HeatMap $usageData.HeatMap -MailboxData ($exchData.StandardMailboxes + $exchData.ArchiveMailboxes) -UsageThresholds $usageThresholds -PrimaryQuotaProp $PrimaryQuotaProp -MoveMapMode $MoveMapMode -Verbose
Write-ExchangeDatabaseUsage -UsageData $usageData -MoveMap $moveMap -OutputPath $OutputPath -Verbose

return $usageData