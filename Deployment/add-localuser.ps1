$MDTUserName = "MDTUser"
$MDTPassword = "Passw0rd!"
$ComputerObject = [ADSI]"WinNT://localhost"
$UserObject = $ComputerObject.Create("User",$MDTUserName)
$UserObject.SetPassword($MDTPassword)
$UserObject.SetInfo() | Out-Null
[ADSI]"WinNT://localhost/$($MDTUserName)"
