# Define input collection
$LastOctets = 1..100

# Define code to run
$CodeToRun = {
    param($LastOctet)

    $Address = "10.156.22.$LastOctet"
    New-Object -TypeName psobject -Property @{
        Address = $Address
        Status = (Test-Connection -ComputerName $Address -Quiet -Count 1)
        HostName = try
        {
            [System.Net.Dns]::GetHostEntry($Address).HostName
        }
        catch 
        {
        }
    }
}

# Create and open the runspacepool
$RunspacePool = [runspacefactory]::CreateRunspacePool()
$RunspacePool.Open()

# Create a new PowerShell instance per "job", collect these along with the IAsyncResult handle (we'll need it later)
$Jobs = foreach($LastOctet in $LastOctets)
{
    $PSInstance = [powershell]::Create()
    [void]$PSInstance.AddScript($CodeToRun).AddParameter('LastOctet',$LastOctet)

    New-Object psobject -Property @{
        Instance = $PSInstance
        IAResult = $PSInstance.BeginInvoke()
    }
}

# Wait for runspaces to complete
while($InProgress = @($Jobs |Where-Object {-not $_.IAResult.IsCompleted})){
    # Here you could also use Write-Progress
    Write-Host "$($InProgress.Count) jobs still in progress..."
    Start-Sleep -Milliseconds 500
}

# Collect the output
$Results = foreach($Job in $Jobs)
{
    $Job.Instance.EndInvoke($Job.IAResult)
}

# Dispose of the runspacepool
$RunspacePool.Dispose()

$Results
0