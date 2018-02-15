$MaxFragmentation = "15"
$DriveDefragJobs = 0

# Defragmentation Processes
function DefragDisk ($MaxFragmentation)
{
	try
	{
		# Get list of Pysical Disks for determining health status and drive type
		$Disks = New-Object -TypeName System.Collections.ArrayList
		
		# Determine if the machine is virtual or physical then query disks
		If ((Get-WmiObject -Class Win32_ComputerSystem | Select-Object -Property Model) -like "*Virtual*")
		{
			$VirtualDisks = (Get-Partition | Get-Disk | Get-PhysicalDisk).DeviceID | Select-Object -Unique | Sort-Object
			
			# Add disk to disk array if not present
			if ($VirtualDisks -ne $null)
			{
				foreach ($Disk in $VirtualDisks)
				{
					if ($Disk -notin $Disks)
					{
						$Disks.Add($Disk) | Out-Null
					}
				}
			}
		}
		else
		{
			$DiskSerialNumbers = Get-Partition | Get-Disk | Select-Object SerialNumber -Unique
			$PhyiscalDisks = foreach ($Serial in $DiskSerialNumbers.SerialNumber) { Get-PhysicalDisk | Where-Object { $_.SerialNumber -eq $Serial.Trim() } | Select-Object FriendlyName, DeviceID | Sort-Object }
			
			# Add disk to disk array if not present
			if ($PhyiscalDisks -ne $null)
			{
				foreach ($Disk in $PhyiscalDisks)
				{
					if ($Disk.DeviceID -notin $Disks)
					{
						$Disks.Add($Disk.DeviceID) | Out-Null
					}
				}
			}
		}
		# Add disk to disk array if not present
		if ($PhyiscalDisks -ne $null)
		{
			foreach ($Disk in $PhyiscalDisks)
			{
				if ($Disk.DeviceID -notin $Disks)
				{
					$Disks.Add($Disk.DeviceID) | Out-Null
				}
			}
		}
		# Start Defrag Process for each disk
		foreach ($Disk in $Disks)
		{
			$CurrentDisk = Get-PhysicalDisk | Where-Object { $_.DeviceID -eq $Disk }
			If ($CurrentDisk.HealthStatus -eq "Healthy")
			{
				$Partitions = Get-Disk -DeviceId $CurrentDisk.DeviceID | Get-Partition | Where-Object { $_.NoDefaultDriveLetter -ne $true } | Get-Volume | Where-Object { $_.DriveType -eq "Fixed" }
				If ($Partitions -ne $null)
				{
					foreach ($Partition in $Partitions)
					{
						# Check Volume Fragmentation Level
						$DriveLetter = ($Partition.DriveLetter + ":")
						$DriveLetter = [string]$DriveLetter.Trim()
						if ($DriveLetter -ne ":")
						{
							$PartitionToDefrag = Get-WmiObject -Class Win32_Volume -Filter "DriveLetter = '$DriveLetter'"
							$PartitionReport = $PartitionToDefrag.DefragAnalysis()
							$FragmentationLevel = $PartitionReport.DefragAnalysis.FilePercentFragmentation
							$FreeSpacePercentage = ($PartitionToDefrag.FreeSpace)/(($PartitionToDefrag.Capacity)/100)
							if ($FragmentationLevel -gt $MaxFragmentation)
							{
								$DriveDefragJobs++
								If ($FreeSpacePercentage -gt "10")
								{
									If ($CurrentDisk.MediaType -eq "SSD")
									{
										Optimize-Volume -DriveLetter $Partition.DriveLetter -ReTrim -AsJob
									}
									else
									{
										Optimize-Volume -DriveLetter $Partition.DriveLetter -Defrag -AsJob
									}
								}
								else
								{
									$Output = "Insufficient Free Space To Defragment Disk"
									Break
								}
							}
						}
					}
				}
			}
			else
			{
				$Output = "Disk Health Status Issue Detected"
				Break
			}
		}
		if ($DriveDefragJobs -gt 0)
		{
			$Output = "Defragmentation Jobs Initialised For $DriveDefragJobs Partitions"
		}
		else
		{
			$Output = "Drive Fragmentation Levels Are Healthy. No Defragment Jobs Required"
		}
		Return $Output
	}	
	catch { }
}

DefragDisk ($MaxFragmentation)
