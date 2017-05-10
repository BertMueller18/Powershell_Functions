# Define list of remote hosts (can be host names or IPs)
$RemoteHosts = @('www.google.com')

# Initiate a Ping asynchronously per remote host, pick up the result task objects
$Tasks = foreach($ComputerName in $RemoteHosts) {
    (New-Object System.Net.NetworkInformation.Ping).SendPingAsync($ComputerName)
}

# Wait for all tasks to finish
[System.Threading.Tasks.Task]::WaitAll($Task)

# Output results
$Tasks |Select-Object Result