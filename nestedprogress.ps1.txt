$ActivityId = Get-Random -Maximum ([int]::MaxValue)
$ItemThreshold = 2000
foreach($n in 1..10){
    Write-Progress -Activity "Doing all the things..." -Status "Working on item $n" -Id $ActivityId -PercentComplete $(100*$n/10)
    $start = Get-Date
    $end   = $start.AddMilliseconds($ItemThreshold)
    while($end -ge ($now = Get-Date)){
        $PercentComplete = 100 * ($now - $start).TotalMilliseconds / $ItemThreshold
        Write-Progress -Activity "Doing one of the things..." -Status "Still working on item $n" -ParentId $ActivityId -PercentComplete $PercentComplete
    }
    Write-Progress -Activity "Doing one of the things..." -Status "Done working on item $n" -ParentId $ActivityId -Completed
}
Write-Progress -Activity "Doing all the things..." -Status "Done!" -Id $ActivityId -Completed