$MaxFragmentation = "15"

# Defragmentation Processes
function DefragCheck ($MaxFragmentation)
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
				Write-Host "Working on $CurrentDisk with partitions $Paritions"
				If ($Partitions -ne $null)
				{
					# Create array for multiple disk drive fragmentation levels
					$VolumeFragmentLevels = New-Object -TypeName System.Collections.ArrayList
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
							if ($FragmentationLevel -gt $MaxFragmentation)
							{
								$VolumeFragmentLevels.Add($FragmentationLevel) | Out-Null
							}
						}
					}
				}
			}
		}
		# Output the drive with the highest fragmentation level
		$CurrentFragmentation = $VolumeFragmentLevels | Sort-Object -Descending | Select-Object -First 1
		
		# Prepare output for SCCM CI. If the highest level of fragmentation is lower
		# than the value specified in the $MaxFragmentation variable, the drive
		# fragmentation level will be considered to be "low".
		
		if ($CurrentFragmentation -gt $MaxFragmentation)
		{
			$Fragmentation = "High"
		}
		else
		{
			$Fragmentation = "Low"
		}
		Return $Fragmentation
	}
	catch { }
}

DefragCheck ($MaxFragmentation)