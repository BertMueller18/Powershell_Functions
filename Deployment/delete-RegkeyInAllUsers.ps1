# ==============================================================================================
# Author : clearByte
# Date : 18.09.2015
# Purpose : go through all user profiles in HKEY_USERS and delete desired key(s)
# ==============================================================================================
 
#GLOBALS =======================================================================================
$strKey = "Software\clearbyte"

# Get current location to return to at the end of the script
$CurLoc = Get-Location
 
# check if HKU branch is already mounted as a PSDrive. If so, remove it first
$HKU = Get-PSDrive HKU -ea silentlycontinue

#END GLOBALS ===================================================================================
 
#MAIN ==========================================================================================
#check HKU branch mount status
if (!$HKU ) {
 # recreate a HKU as a PSDrive and navigate to it
 New-PSDrive -Name HKU -PsProvider Registry HKEY_USERS | out-null
 Set-Location HKU:
}
 
# select all desired user profiles, exlude *_classes & .DEFAULT
$regProfiles = Get-ChildItem -Path HKU: | ? { ($_.PSChildName.Length -gt 8) -and ($_.PSChildName -notlike "*.DEFAULT") }

# loop through all selected profiles & delete registry
ForEach ($profile in $regProfiles ) {
 If(Test-Path -Path $profile\$strKey){
 Remove-Item -Path $profile\$strKey -recurse
 }
}
 
# return to initial location at the end of the execution
Set-Location $CurLoc
Remove-PSDrive -Name HKU
 
#END MAIN ======================================================================================