﻿<#
The sample scripts are not supported under any Microsoft standard support 
program or service. The sample scripts are provided AS IS without warranty  
of any kind. Microsoft further disclaims all implied warranties including,  
without limitation, any implied warranties of merchantability or of fitness for 
a particular purpose. The entire risk arising out of the use or performance of  
the sample scripts and documentation remains with you. In no event shall 
Microsoft, its authors, or anyone else involved in the creation, production, or 
delivery of the scripts be liable for any damages whatsoever (including, 
without limitation, damages for loss of business profits, business interruption, 
loss of business information, or other pecuniary loss) arising out of the use 
of or inability to use the sample scripts or documentation, even if Microsoft 
has been advised of the possibility of such damages.
#> 

#requires -Version 2
Param
(
	[String]$Identity
)
If ($Identity -eq "")
{
	$DistributionList=Get-DistributionGroup -ResultSize unlimited
}
Else
{
	$DistributionList=Get-DistributionGroup -Identity $Identity
}
Foreach($DistributionGroup in $DistributionList)
{
	$DisabledMembers=Get-DistributionGroupMember -Identity $DistributionGroup -ResultSize unlimited|`
	? {$_.RecipientType -eq "User"}
	$DisabledMembers|%{Remove-DistributionGroupMember -Identity $DistributionGroup -Member $_ -Confirm:$false}
}
