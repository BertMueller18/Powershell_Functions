# Install WSUS and execute post install
Write-Output "Execute: Install WSUS and execute post install"
$WSUSContentFolder = "C:\WSUSContent"
if (-not(Test-Path -Path $WSUSContentFolder)) {
    New-Item -Path $WSUSContentFolder -ItemType Directory | Out-Null
}
Add-WindowsFeature -Name "UpdateServices","UpdateServices-WidDB","UpdateServices-Services","UpdateServices-RSAT","UpdateServices-API","UpdateServices-UI"
$WSUSUtil = "$($Env:ProgramFiles)\Update Services\Tools\WsusUtil.exe"
$WSUSUtilArgs = "POSTINSTALL CONTENT_DIR=$($WSUSContentFolder)"
Start-Process -FilePath $WSUSUtil -ArgumentList $WSUSUtilArgs -Wait
 
# Load WSUS assembly
Write-Output "Execute: Load WSUS assembly"
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration")
 
# Configure WSUS to synchronize from Microsoft Update
Write-Output "Execute: Configure WSUS to synchronize from Microsoft Update"
$WSUSServer = Get-WsusServer -Name localhost -Port 8530
Set-WsusServerSynchronization -UpdateServer $WSUSServer -SyncFromMU
 
# Disable all Update Classifications
Write-Output "Execute: Disable all Update Classifications"
$WSUSUpdateClassificationList = New-Object System.Collections.ArrayList
$WSUSUpdateClassificationList.AddRange(@("Critical Updates","Definition Updates","Drivers","Feature Packs","Security Updates","Service Packs","Tools","Update Rollups","Updates"))
foreach ($WSUSUpdateClassification in $WSUSUpdateClassificationList) {
    Get-WsusClassification | Where-Object { $_.Classification.Title -like "$($WSUSUpdateClassification)" } | Set-WsusClassification -Disable
}
 
# Enable required Update Classifications for Windows 10
Write-Output "Execute: Enable required Update Classifications for Windows 10"
$WSUSUpdateClassificationEnabledList = New-Object System.Collections.ArrayList
$WSUSUpdateClassificationEnabledList.AddRange(@("Critical Updates","Feature Packs","Security Updates"))
foreach ($WSUSUpdateClassification in $WSUSUpdateClassificationEnabledList) {
    Get-WsusClassification | Where-Object { $_.Classification.Title -eq $WSUSUpdateClassification } | Set-WsusClassification
}
 
# Configure Default Automatic Approval rule
Write-Output "Execute: Configure Default Automatic Approval rule"
$WSUSConnection = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer("localhost", $false, "8530")
$WSUSRuleName = "Default Automatic Approval Rule"
$WSUSRule = $WSUSConnection.GetInstallApprovalRules() | Where-Object { $_.Name -like $WSUSRuleName }
$WSUSClassificationCollection = New-Object Microsoft.UpdateServices.Administration.UpdateClassificationCollection
foreach ($UpdateTitle in $WSUSUpdateClassificationEnabledList) {
    $RuleUpdateClassification = $WSUSConnection.GetUpdateClassifications() | Where-Object { $_.Title -like $UpdateTitle }
    $WSUSClassificationCollection.Add($RuleUpdateClassification) | Out-Null
}
$WSUSRule.SetUpdateClassifications($WSUSClassificationCollection)
$WSUSRule.Enabled = $true
$WSUSRule.Save()
 
# Create a WSUS automatic synchronization schedule
Write-Output "Execute: Create a WSUS automatic synchronization schedule"
$WSUSSubscription = $WSUSConnection.GetSubscription()
$WSUSSubscription.SynchronizeAutomatically = $true
$WSUSSubscription.SynchronizeAutomaticallyTimeOfDay = (New-TimeSpan -Hours 6)
$WSUSSubscription.Save()
 
# Configure WSUS update languages
Write-Output "Execute: Configure WSUS update languages"
$WSUSLanguages = @("en","sv")
$WSUSLanguageCollection = New-Object System.Collections.ArrayList
$WSUSConfiguration = $WSUSConnection.GetConfiguration()
$WSUSConfiguration.AllUpdateLanguagesEnabled = $false
foreach ($WSUSLanguage in $WSUSLanguages) {
    $WSUSLanguageCollection.Add($WSUSLanguage) | Out-Null
}
$WSUSConfiguration.SetEnabledUpdateLanguages($WSUSLanguageCollection)
$WSUSConfiguration.Save()
 
# Start WSUS initial synchronization
Write-Output "Execute: Start WSUS initial synchronization"
$WSUSSubscription.StartSynchronizationForCategoryOnly()
while ($WSUSSubscription.GetSynchronizationStatus() -ne "NotProcessing") {
    Write-Host "." -NoNewline
    Start-Sleep -Seconds 3
}
 
# Configure WSUS products
Write-Output "Execute: Configure WSUS products"
$WSUSSubscription = $WSUSConnection.GetSubscription()
$WSUSCategoryCollection = New-Object Microsoft.UpdateServices.Administration.UpdateCategoryCollection
$WSUSProduct = $WSUSConnection.GetUpdateCategories() | Where-Object { $_.Id -eq "a3c2375d-0c8a-42f9-bce0-28333e198407" }
$WSUSCategoryCollection.Add($WSUSProduct) | Out-Null
$WSUSSubscription.SetUpdateCategories($WSUSCategoryCollection)
$WSUSSubscription.Save()
 
# Start WSUS synchronization
Write-Output "Execute: Start WSUS synchronization"
$WSUSSubscription.StartSynchronization()
while ($WSUSSubscription.GetSynchronizationStatus() -ne "NotProcessing") {
    Write-Host "." -NoNewline
    Start-Sleep -Seconds 3
}
 
# Run Default Automatic Approval rule
Write-Output "Execute: Run Default Automatic Approval rule"
$WSUSConnection = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer("localhost", $false, "8530")
$WSUSAutomaticRule = $WSUSConnection.GetInstallApprovalRules() | Where-Object { $_.Name -like "Default Automatic Approval Rule" }
$WSUSAutomaticRule.ApplyRule()
