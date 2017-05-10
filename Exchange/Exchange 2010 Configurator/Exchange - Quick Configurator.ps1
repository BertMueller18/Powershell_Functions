#----------------------------------------------
#region Import Assemblies
#----------------------------------------------
[void][Reflection.Assembly]::Load("System, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
[void][Reflection.Assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
[void][Reflection.Assembly]::Load("System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
[void][Reflection.Assembly]::Load("mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
[void][Reflection.Assembly]::Load("System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
[void][Reflection.Assembly]::Load("System.Xml, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
[void][Reflection.Assembly]::Load("System.DirectoryServices, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
#endregion Import Assemblies

function Main {
	Param ([String]$Commandline)
	#Note: This function starts the application
	#Note: $Commandline contains the complete argument string passed to the packager
	#Note: $Args contains the parsed arguments passed to the packager (Type: System.Array)
	#Note: To get the script directory in the Packager use: Split-Path $hostinvocation.MyCommand.path
	#Note: To get the console output in the Packager (Windows Mode) use: $ConsoleOutput (Type: System.Collections.ArrayList)
	#TODO: Initialize and add Function calls to forms
	
	if((Call-MainForm_pff) -eq "OK")
	{
		
	}
	
	$global:ExitCode = 0 #Set the exit code for the Packager
}
#endregion Source: Startup.pfs

#region Source: MainForm.pff
function Call-MainForm_pff
{
	#----------------------------------------------
	#region Import the Assemblies
	#----------------------------------------------
	[void][reflection.assembly]::Load("System, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
	[void][reflection.assembly]::Load("mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Xml, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.DirectoryServices, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
	#endregion Import Assemblies

	#----------------------------------------------
	#region Generated Form Objects
	#----------------------------------------------
	[System.Windows.Forms.Application]::EnableVisualStyles()
	$frmMain = New-Object 'System.Windows.Forms.Form'
	$btnLoad = New-Object 'System.Windows.Forms.Button'
	$labelLoadConfig = New-Object 'System.Windows.Forms.Label'
	$buttonBrowse2 = New-Object 'System.Windows.Forms.Button'
	$textboxLoad = New-Object 'System.Windows.Forms.TextBox'
	$btnSave = New-Object 'System.Windows.Forms.Button'
	$labelSaveConfig = New-Object 'System.Windows.Forms.Label'
	$buttonBrowse = New-Object 'System.Windows.Forms.Button'
	$textboxSave = New-Object 'System.Windows.Forms.TextBox'
	$buttonBrowseFolder = New-Object 'System.Windows.Forms.Button'
	$txtOutputDir = New-Object 'System.Windows.Forms.TextBox'
	$tabcontrol1 = New-Object 'System.Windows.Forms.TabControl'
	$tabpage1 = New-Object 'System.Windows.Forms.TabPage'
	$btnAdd = New-Object 'System.Windows.Forms.Button'
	$groupboxGeneral = New-Object 'System.Windows.Forms.GroupBox'
	$labelDomain = New-Object 'System.Windows.Forms.Label'
	$txtInternalDomain = New-Object 'System.Windows.Forms.TextBox'
	$txtSiteName = New-Object 'System.Windows.Forms.TextBox'
	$comboOS = New-Object 'System.Windows.Forms.ComboBox'
	$txtServerName = New-Object 'System.Windows.Forms.TextBox'
	$txtRegion = New-Object 'System.Windows.Forms.TextBox'
	$comboExchRole = New-Object 'System.Windows.Forms.ComboBox'
	$labelServerName = New-Object 'System.Windows.Forms.Label'
	$comboExchVer = New-Object 'System.Windows.Forms.ComboBox'
	$labelOS = New-Object 'System.Windows.Forms.Label'
	$labelVersion = New-Object 'System.Windows.Forms.Label'
	$labelRegion = New-Object 'System.Windows.Forms.Label'
	$labelSite = New-Object 'System.Windows.Forms.Label'
	$labelRole = New-Object 'System.Windows.Forms.Label'
	$groupboxURL = New-Object 'System.Windows.Forms.GroupBox'
	$txtArrayDAGURL = New-Object 'System.Windows.Forms.TextBox'
	$txtAutodiscoverURL = New-Object 'System.Windows.Forms.TextBox'
	$txtCASURL = New-Object 'System.Windows.Forms.TextBox'
	$labelAccessURL = New-Object 'System.Windows.Forms.Label'
	$labelAutodiscoverURL = New-Object 'System.Windows.Forms.Label'
	$labelArrayDAGURL = New-Object 'System.Windows.Forms.Label'
	$groupboxIPInfo = New-Object 'System.Windows.Forms.GroupBox'
	$comboNLBClusterMode = New-Object 'System.Windows.Forms.ComboBox'
	$txtNLBDAGIP = New-Object 'System.Windows.Forms.TextBox'
	$labelNLBDAGIP = New-Object 'System.Windows.Forms.Label'
	$chkSSLOffload = New-Object 'System.Windows.Forms.CheckBox'
	$txtIPGateway = New-Object 'System.Windows.Forms.TextBox'
	$txtIPSubnet = New-Object 'System.Windows.Forms.TextBox'
	$labelSSLOffload = New-Object 'System.Windows.Forms.Label'
	$txtIPAddress = New-Object 'System.Windows.Forms.TextBox'
	$labelInterfaceName = New-Object 'System.Windows.Forms.Label'
	$chkInternetFacing = New-Object 'System.Windows.Forms.CheckBox'
	$labelGateway = New-Object 'System.Windows.Forms.Label'
	$labelInternetFacing = New-Object 'System.Windows.Forms.Label'
	$labelIPAddress = New-Object 'System.Windows.Forms.Label'
	$chkNLB = New-Object 'System.Windows.Forms.CheckBox'
	$labelSubnet = New-Object 'System.Windows.Forms.Label'
	$labelWindowsNLB = New-Object 'System.Windows.Forms.Label'
	$labelReplicationNetwork = New-Object 'System.Windows.Forms.Label'
	$chkDAGReplication = New-Object 'System.Windows.Forms.CheckBox'
	$txtIntName = New-Object 'System.Windows.Forms.TextBox'
	$groupbox1 = New-Object 'System.Windows.Forms.GroupBox'
	$label1 = New-Object 'System.Windows.Forms.Label'
	$label2 = New-Object 'System.Windows.Forms.Label'
	$label3 = New-Object 'System.Windows.Forms.Label'
	$label4 = New-Object 'System.Windows.Forms.Label'
	$label5 = New-Object 'System.Windows.Forms.Label'
	$label6 = New-Object 'System.Windows.Forms.Label'
	$label7 = New-Object 'System.Windows.Forms.Label'
	$tabpage2 = New-Object 'System.Windows.Forms.TabPage'
	$labelGeneratePrerequisite = New-Object 'System.Windows.Forms.Label'
	$chkGeneratePrerequisites = New-Object 'System.Windows.Forms.CheckBox'
	$labelGenerateMigrationCoe = New-Object 'System.Windows.Forms.Label'
	$chkGenerateMigrationTips = New-Object 'System.Windows.Forms.CheckBox'
	$labelExternalDNSNamespace = New-Object 'System.Windows.Forms.Label'
	$txtExternalDomainNameSpace = New-Object 'System.Windows.Forms.TextBox'
	$labelUseInternalCASArrayD = New-Object 'System.Windows.Forms.Label'
	$chkUseInternalCASDomain = New-Object 'System.Windows.Forms.CheckBox'
	$labelGenerateCASArrays = New-Object 'System.Windows.Forms.Label'
	$chkGenerateCASArrays = New-Object 'System.Windows.Forms.CheckBox'
	$labelAssumeReverseProxy = New-Object 'System.Windows.Forms.Label'
	$chkReverseProxy = New-Object 'System.Windows.Forms.CheckBox'
	$labelRedirectURLs = New-Object 'System.Windows.Forms.Label'
	$chkRedirection = New-Object 'System.Windows.Forms.CheckBox'
	$labelInternalExchangeVirt = New-Object 'System.Windows.Forms.Label'
	$txtvdirlocation = New-Object 'System.Windows.Forms.TextBox'
	$labelOutputFolder = New-Object 'System.Windows.Forms.Label'
	$chkGunSlinger = New-Object 'System.Windows.Forms.CheckBox'
	$btnRun = New-Object 'System.Windows.Forms.Button'
	$btnGenerate = New-Object 'System.Windows.Forms.Button'
	$btnExit = New-Object 'System.Windows.Forms.Button'
	$groupboxServerList = New-Object 'System.Windows.Forms.GroupBox'
	$datagridServers = New-Object 'System.Windows.Forms.DataGridView'
	$btnDelete = New-Object 'System.Windows.Forms.Button'
	$errorprovider1 = New-Object 'System.Windows.Forms.ErrorProvider'
	$folderbrowserdialog1 = New-Object 'System.Windows.Forms.FolderBrowserDialog'
	$openfiledialogSave = New-Object 'System.Windows.Forms.OpenFileDialog'
	$openfiledialogLoad = New-Object 'System.Windows.Forms.OpenFileDialog'
	$errorprovider2 = New-Object 'System.Windows.Forms.ErrorProvider'
	$errorprovider3 = New-Object 'System.Windows.Forms.ErrorProvider'
	$errorprovider4 = New-Object 'System.Windows.Forms.ErrorProvider'
	$ttipGeneral = New-Object 'System.Windows.Forms.ToolTip'
	$InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'
	#endregion Generated Form Objects

	#region Control Helper Functions
	function Load-DataGridView
	{
		<#
		.SYNOPSIS
			This functions helps you load items into a DataGridView.
	
		.DESCRIPTION
			Use this function to dynamically load items into the DataGridView control.
	
		.PARAMETER  DataGridView
			The ComboBox control you want to add items to.
	
		.PARAMETER  Item
			The object or objects you wish to load into the ComboBox's items collection.
		
		.PARAMETER  DataMember
			Sets the name of the list or table in the data source for which the DataGridView is displaying data.
	
		#>
		Param (
			[ValidateNotNull()]
			[Parameter(Mandatory=$true)]
			[System.Windows.Forms.DataGridView]$DataGridView,
			[ValidateNotNull()]
			[Parameter(Mandatory=$true)]
			$Item,
		    [Parameter(Mandatory=$false)]
			[string]$DataMember
		)
		$DataGridView.SuspendLayout()
		$DataGridView.DataMember = $DataMember
		
		if ($Item -is [System.ComponentModel.IListSource]`
		-or $Item -is [System.ComponentModel.IBindingList] -or $Item -is [System.ComponentModel.IBindingListView] )
		{
			$DataGridView.DataSource = $Item
		}
		else
		{
			$array = New-Object System.Collections.ArrayList
			
			if ($Item -is [System.Collections.IList])
			{
				$array.AddRange($Item)
			}
			else
			{	
				$array.Add($Item)	
			}
			$DataGridView.DataSource = $array
		}
		
		$DataGridView.ResumeLayout()
	}
	
	function Load-ComboBox 
	{
	<#
		.SYNOPSIS
			This functions helps you load items into a ComboBox.
	
		.DESCRIPTION
			Use this function to dynamically load items into the ComboBox control.
	
		.PARAMETER  ComboBox
			The ComboBox control you want to add items to.
	
		.PARAMETER  Items
			The object or objects you wish to load into the ComboBox's Items collection.
	
		.PARAMETER  DisplayMember
			Indicates the property to display for the items in this control.
		
		.PARAMETER  Append
			Adds the item(s) to the ComboBox without clearing the Items collection.
		
		.EXAMPLE
			Load-ComboBox $combobox1 "Red", "White", "Blue"
		
		.EXAMPLE
			Load-ComboBox $combobox1 "Red" -Append
			Load-ComboBox $combobox1 "White" -Append
			Load-ComboBox $combobox1 "Blue" -Append
		
		.EXAMPLE
			Load-ComboBox $combobox1 (Get-Process) "ProcessName"
	#>
		Param (
			[ValidateNotNull()]
			[Parameter(Mandatory=$true)]
			[System.Windows.Forms.ComboBox]$ComboBox,
			[ValidateNotNull()]
			[Parameter(Mandatory=$true)]
			$Items,
		    [Parameter(Mandatory=$false)]
			[string]$DisplayMember,
			[switch]$Append
		)
		
		if(-not $Append)
		{
			$ComboBox.Items.Clear()	
		}
		
		if($Items -is [Object[]])
		{
			$ComboBox.Items.AddRange($Items)
		}
		elseif ($Items -is [Array])
		{
			$ComboBox.BeginUpdate()
			foreach($obj in $Items)
			{
				$ComboBox.Items.Add($obj)	
			}
			$ComboBox.EndUpdate()
		}
		else
		{
			$ComboBox.Items.Add($Items)	
		}
	
		$ComboBox.DisplayMember = $DisplayMember	
	}
	
	function Load-ListBox 
	{
	<#
		.SYNOPSIS
			This functions helps you load items into a ListBox.
	
		.DESCRIPTION
			Use this function to dynamically load items into the ListBox control.
	
		.PARAMETER  ListBox
			The ListBox control you want to add items to.
	
		.PARAMETER  Items
			The object or objects you wish to load into the ListBox's Items collection.
	
		.PARAMETER  DisplayMember
			Indicates the property to display for the items in this control.
		
		.PARAMETER  Append
			Adds the item(s) to the ListBox without clearing the Items collection.
		
		.EXAMPLE
			Load-ListBox $ListBox1 "Red", "White", "Blue"
		
		.EXAMPLE
			Load-ListBox $listBox1 "Red" -Append
			Load-ListBox $listBox1 "White" -Append
			Load-ListBox $listBox1 "Blue" -Append
		
		.EXAMPLE
			Load-ListBox $listBox1 (Get-Process) "ProcessName"
	#>
		Param (
			[ValidateNotNull()]
			[Parameter(Mandatory=$true)]
			[System.Windows.Forms.ListBox]$listBox,
			[ValidateNotNull()]
			[Parameter(Mandatory=$true)]
			$Items,
		    [Parameter(Mandatory=$false)]
			[string]$DisplayMember,
			[switch]$Append
		)
		
		if(-not $Append)
		{
			$listBox.Items.Clear()	
		}
		
		if($Items -is [System.Windows.Forms.ListBox+ObjectCollection])
		{
			$listBox.Items.AddRange($Items)
		}
		elseif ($Items -is [Array])
		{
			$listBox.BeginUpdate()
			foreach($obj in $Items)
			{
				$listBox.Items.Add($obj)
			}
			$listBox.EndUpdate()
		}
		else
		{
			$listBox.Items.Add($Items)	
		}
	
		$listBox.DisplayMember = $DisplayMember	
	}
	#endregion
	
	$buttonBrowse_Click={
	
		if($openfiledialogSave.ShowDialog() -eq 'OK')
		{
			$txtOutputFile.Text = $openfiledialogSave.FileName
		}
	}
	
	
	$datagridServers_Navigate=[System.Windows.Forms.NavigateEventHandler]{
	#Event Argument: $_ = [System.Windows.Forms.NavigateEventArgs]
	
		$datagridServers.Row
	}
	
	$btnAdd_Click={
		if ($comboExchRole.Text -and $comboExchVer.Text -and $comboOS.Text)
		{
			$ServerToAdd = New-Object PSObject -Property @{
			    "Server Name" = $txtServerName.Text
			    "Windows Version" = $comboOS.Text
			    "Role" = $comboExchRole.Text
	            "Internal Domain" = $txtInternalDomain.Text
			    "Exchange Version" = $comboExchVer.Text
			    "Site" = $txtSiteName.Text
			    "Region" = $txtRegion.Text
			    "IP" = $txtIPAddress.Text
			    "IP Subnet" = $txtIPSubnet.Text
			    "IP Gateway" = $txtIPGateway.Text
			    "Interface Name" = $txtIntName.Text
			    "NLB or DAG IP" = $txtNLBDAGIP.Text
			    "DAG Replication" = $chkDAGReplication.Checked		
			    "Internet Facing" = $chkInternetFacing.Checked
			    "NLB" = $chkNLB.Checked
			    "NLB Mode" = $comboNLBClusterMode.Text
			    "Autodiscover URL" = $txtAutodiscoverURL.Text
			    "CAS URL" = $txtCASURL.Text
	            "ArrayDAG URL" = $txtArrayDAGURL.Text
			    "SSL Offload" = $chkSSLOffload.Checked
			}
			$Script:Servers+=$ServerToAdd
			$ServerList = $Script:Servers | Out-DataTable
			$datagridServers.DataSource = $ServerList
			$datagridServers.Refresh
		}
		Else
		{
			#[void][reflection.assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
			[void][System.Windows.Forms.MessageBox]::Show("Missing some required info bub, try again.","Sorry")
		}
	}
	
	$btnDelete_Click={
		$selectedRow = $dataGridServers.CurrentRow.Index
		if ($Script:Servers.Count -le 1)
		{
			$Script:Servers = @()
		}
		else
		{
			$Script:Servers = $Script:Servers | Where-Object {$_ -ne $Script:Servers[$selectedRow]}
		
		}
		$ServerList = $Script:Servers | Out-DataTable
		$datagridServers.DataSource = $ServerList
		$datagridServers.Refresh
		
	}
	
	$btnGenerate_Click={
	    $Redirection = $chkRedirection.Checked
	    $vDirLocation = $txtvdirlocation.Text
	    $ReverseProxy = $chkReverseProxy.Checked
	    $GenerateCASArrays = $chkGenerateCASArrays.Checked
	    $UseCASInternalDomainNameSpace = $chkUseInternalCASDomain.Checked
	    $ExternalNamespace = $txtExternalDomainNameSpace.Text
		$GenerateMigrationTips = $chkGenerateMigrationTips.Checked
		$GeneratePrerequisties = $chkGeneratePrerequisites.Checked
		If ($txtOutputDir)
		{
			Sort-ExchangeServers
			
			[string]$path = $txtOutputDir.Text
			if ($GeneratePrerequisties)
			{
				$prereqs = . Generate-Preqs
				Out-File -FilePath "$path\generate-prereqs.ps1" -Force -InputObject $prereqs -Width 400 -Encoding ascii
			}
			if ($GenerateCASArrays)
			{
				$hubcasconfig = . Generate-CASConfiguration
				Out-File -FilePath "$path\generate-hubcasroles.ps1" -Force -InputObject $hubcasconfig -Width 400 -Encoding ascii
			}
			if ($GenerateMigrationTips)
			{
				$MigrationTips = . Generate-Coexistence
				Out-File -FilePath "$path\MigrationTips.ps1" -Force -InputObject $MigrationTips -Width 400 -Encoding ascii
			}
		}
		
		Else
		{
			#[void][reflection.assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
			[void][System.Windows.Forms.MessageBox]::Show("Sorry, no file was specified for output","Where's the output man?")
		}
	}
	
	$chkGunSlinger_CheckedChanged={
		if ($chkGunSlinger.Checked)
		{
			$btnRun.Enabled = $true
		}
		else
		{
			$btnRun.Enabled = $false
		}
	}
	
	$comboExchRole_SelectedValueChanged={
		switch ($comboExchRole.Text) {
			"HUB/CAS" {
				#Enable/Disable appropriate form elements
				$chkSSLOffload.Enabled = $true
				$chkInternetFacing.Enabled = $true
				$chkNLB.Enabled = $true
				$txtCASURL.Enabled = $true
				$txtArrayDAGURL.Enabled = $true
				$txtAutodiscoverURL.Enabled = $true
				$chkDAGReplication.Enabled = $false
				
				#Set default values
				$chkDAGReplication.Checked = $false
			}
			"Database" {
				#Enable/Disable appropriate form elements
				$chkDAGReplication.Enabled = $true
				$chkSSLOffload.Enabled = $false
				$chkInternetFacing.Enabled = $false
				$chkNLB.Enabled = $false
				$txtCASURL.Enabled = $false
				$txtArrayDAGURL.Enabled = $true
				$txtAutodiscoverURL.Enabled = $false
				$txtNLBDAGIP.Enabled = $true
	            
				#Set default values
				$chkInternetFacing.Checked = $false
				$chkNLB.Checked = $false
				$chkSSLOffload.Checked = $false			
				$txtCASURL.Text = '(ie. owa.contoso.com)'
				$txtAutodiscoverURL.Text = '(ie. autodiscover.contoso.com)'
				$txtArrayDAGURL.Text = '(ie. dag1.contoso.com)'
			}
			"Edge" {
				#Enable/Disable appropriate form elements
				$chkSSLOffload.Enabled = $false
				$chkInternetFacing.Enabled = $false
				$chkNLB.Enabled = $false
				$txtCASURL.Enabled = $false
				$txtArrayDAGURL.Enabled = $false
				$txtAutodiscoverURL.Enabled = $false
				$chkDAGReplication.Enabled = $false
				
				#Set default values
				$chkInternetFacing.Checked = $false
				$chkNLB.Checked = $false
				$chkSSLOffload.Checked = $false
				$chkDAGReplication.Checked = $false
				$txtCASURL.Text = '(ie. owa.contoso.com)'
				$txtAutodiscoverURL.Text = '(ie. autodiscover.contoso.com)'
				$txtArrayDAGURL.Text = '(ie. owa.contoso.com)'
			}
			"Front-End" {
				#Enable/Disable appropriate form elements
				$chkDAGReplication.Enabled = $false
	
				#Set default values
				$chkInternetFacing.Checked = $false
				$chkNLB.Checked = $false
				$chkSSLOffload.Checked = $false
				$chkDAGReplication.Checked = $false
				$txtCASURL.Text = '(ie. owa.contoso.com)'
				$txtAutodiscoverURL.Text = '(ie. autodiscover.contoso.com)'
				$txtArrayDAGURL.Text = '(ie. owa.contoso.com)'
	
			}
			"Back-End" {
				#Set default values
				$chkDAGReplication.Checked = $false
				$chkSSLOffload.Enabled = $false
				$chkInternetFacing.Enabled = $false
				$chkNLB.Enabled = $false
				$txtCASURL.Enabled = $false
				$txtArrayDAGURL.Enabled = $false
				$txtAutodiscoverURL.Enabled = $false
				$chkDAGReplication.Enabled = $false
	            $txtNLBDAGIP.Enabled = $false
				
				#Set default values
				$chkInternetFacing.Checked = $false
				$chkNLB.Checked = $false
				$chkSSLOffload.Checked = $false
				$chkDAGReplication.Checked = $false
				$txtCASURL.Text = '(ie. owa.contoso.com)'
				$txtAutodiscoverURL.Text = '(ie. autodiscover.contoso.com)'
				$txtArrayDAGURL.Text = '(ie. owa.contoso.com)'
			}
			"CAS" {
				#Set default values
				$chkDAGReplication.Checked = $false
	            $txtNLBDAGIP.Enabled = $true
				$chkNLB.Enabled = $true
			}
			"Hub/Transport/Database" {
				$chkDAGReplication.Enabled = $true
				$chkSSLOffload.Enabled = $false
				$chkInternetFacing.Enabled = $false
				$chkNLB.Enabled = $false
				$txtCASURL.Enabled = $false
				$txtArrayDAGURL.Enabled = $true
				$txtAutodiscoverURL.Enabled = $false
	            $txtNLBDAGIP.Enabled = $true
			}
	        "All-In-One" {
				$chkDAGReplication.Enabled = $true
				$chkSSLOffload.Enabled = $true
				$chkInternetFacing.Enabled = $true
				$chkNLB.Enabled = $false
				$txtCASURL.Enabled = $true
				$txtArrayDAGURL.Enabled = $true
				$txtAutodiscoverURL.Enabled = $true
	            $txtNLBDAGIP.Enabled = $true
			}
			default {
				#Enable/Disable appropriate form elements
				$chkSSLOffload.Enabled = $false
				$chkInternetFacing.Enabled = $false
				$chkNLB.Enabled = $false
				$txtCASURL.Enabled = $false
				$txtArrayDAGURL.Enabled = $false
				$txtAutodiscoverURL.Enabled = $false
				$chkDAGReplication.Enabled = $false
			}
		}
	}
	
	$textboxIP_Validating=[System.ComponentModel.CancelEventHandler]{
	#Event Argument: $_ = [System.ComponentModel.CancelEventArgs]
	
		$result = -not (Validate-IsIP $txtIPAddress.Text)
	
		if($result -eq $true)
		{
			$_.Cancel = $true
			$errorprovider1.SetError($this, "Invalid IP address");
		}
	}
	
	$control_Validated={
		#Pass the calling control and clear error message
		$errorprovider1.SetError($this, "");	
	}
	
	function Validate-IsIP ([string] $IP)
	{
		<#
			.SYNOPSIS
				Validates if input is an IP Address
		
			.DESCRIPTION
				Validates if input is an IP Address
		
			.PARAMETER  IP
				A string containing an IP address
		
			.INPUTS
				System.String
		
			.OUTPUTS
				System.Boolean
		#>	
		#Regular Express
		#return $IP -match "\b(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b"
		#Parse using a IPAddress static method
		try
		{
			return ([System.Net.IPAddress]::Parse($IP) -ne $null)
		}
		catch
		{ }
		
		return $false	
	}
	
	$txtIPAddress_Validating2=[System.ComponentModel.CancelEventHandler]{
	#Event Argument: $_ = [System.ComponentModel.CancelEventArgs]
	
		$result = -not (Validate-IsIP $txtIPAddress.Text)
	
		if($result -eq $true)
		{
			$_.Cancel = $true
			$errorprovider1.SetError($this, "Invalid IP address");
		}
	}
	
	$control_Validated2={
		#Pass the calling control and clear error message
		$errorprovider1.SetError($this, "");	
	}
	$buttonBrowse2_Click={
	
		if($openfiledialogLoad.ShowDialog() -eq 'OK')
		{
			$textboxSave.Text = $openfiledialogLoad.FileName
		}
	}
	
	$buttonBrowseFolder_Click={
		if($folderbrowserdialog1.ShowDialog() -eq 'OK')
		{
			$txtOutputDir.Text = $folderbrowserdialog1.SelectedPath
		}
	}
	
	$buttonBrowse_Click2={
	
		if($openfiledialogSave.ShowDialog() -eq 'OK')
		{
			$textboxSave.Text = $openfiledialogSave.FileName
		}
	}
	
	$buttonBrowse2_Click2={
	
		if($openfiledialogLoad.ShowDialog() -eq 'OK')
		{
			$textboxLoad.Text = $openfiledialogLoad.FileName
		}
	}
	
	$btnSave_Click={
		if ($textboxSave.Text)
		{
			$Script:Servers | Export-Csv -NoTypeInformation -Path $textboxSave.Text
		}
		Else
		{
			#[void][reflection.assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
			[void][System.Windows.Forms.MessageBox]::Show("Need a filename","What the heck?")
		}
	}
	
	$btnLoad_Click={
		if ($textboxLoad.Text)
		{
			$Script:Servers = Import-Csv -Path $textboxLoad.Text
			$ServerList = $Script:Servers | Out-DataTable
			$datagridServers.DataSource = $ServerList
			$datagridServers.Refresh
		}
		Else
		{
			#[void][reflection.assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
			[void][System.Windows.Forms.MessageBox]::Show("Need a file to load","Huh?")
		}
	}
	
	$btnExit_Click={
		$frmMain.Close()	
	}
	
	
	$comboExchVer_SelectedIndexChanged={
			$comboExchRole.Enabled = $true		
			$chkSSLOffload.Enabled = $false
			$chkInternetFacing.Enabled = $false
			$chkNLB.Enabled = $false
			$txtCASURL.Enabled = $false
			$txtArrayDAGURL.Enabled = $false
			$txtAutodiscoverURL.Enabled = $false
			$chkDAGReplication.Enabled = $false
			switch ($comboExchVer.Text) 
	        {
	    		"Exchange 2003" 
	            {
	    			Load-ComboBox $comboExchRole "Front-End"
	    			Load-ComboBox $comboExchRole "Back-End" -Append
				 	Load-ComboBox $comboExchRole "All-In-One" -Append
					Load-ComboBox $comboOS "Windows 2003"
	    		}
	    		"Exchange 2007" 
	            {
	    			Load-ComboBox $comboExchRole "Front-End"
	    			Load-ComboBox $comboExchRole "Back-End" -Append
				 	Load-ComboBox $comboExchRole "All-In-One" -Append
					Load-ComboBox $comboOS "Windows 2003"
					Load-ComboBox $comboOS "Windows 2008" -Append
					Load-ComboBox $comboOS "Windows 2008 R2" -Append
	    		}
	    		"Exchange 2010" 
	            {
	    			Load-ComboBox $comboExchRole "HUB/CAS"
	    			Load-ComboBox $comboExchRole "Database" -Append
	    			Load-ComboBox $comboExchRole "Edge" -Append
	                Load-ComboBox $comboExchRole "All-In-One" -Append
					Load-ComboBox $comboOS "Windows 2008"
					Load-ComboBox $comboOS "Windows 2008 R2" -Append
					Load-ComboBox $comboOS "Windows 2012" -Append
	    		}
	    		"Exchange 2012" 
	            {
	    			Load-ComboBox $comboExchRole "CAS"
	    			Load-ComboBox $comboExchRole "HUB/Transport/Database" -Append
	                Load-ComboBox $comboExchRole "All-In-One" -Append
					Load-ComboBox $comboOS "Windows 2008 R2"
					Load-ComboBox $comboOS "Windows 2012" -Append
	    		}
	    		default 
	    		{
	    		
	    		}
		}	
	}
	
	
	$textboxIP_Validating2=[System.ComponentModel.CancelEventHandler]{
	#Event Argument: $_ = [System.ComponentModel.CancelEventArgs]
	
		$result = -not (Validate-IsIP $txtIPSubnet.Text)
	
		if($result -eq $true)
		{
			$_.Cancel = $true
			$errorprovider2.SetError($this, "Invalid IP address");
		}
	}
	
	$control_Validated3={
		#Pass the calling control and clear error message
		$errorprovider2.SetError($this, "");	
	}
	
	$textboxIP_Validating3=[System.ComponentModel.CancelEventHandler]{
	#Event Argument: $_ = [System.ComponentModel.CancelEventArgs]
	
		$result = -not (Validate-IsIP $txtIPSubnet.Text)
	
		if($result -eq $true)
		{
			$_.Cancel = $true
			$errorprovider2.SetError($this, "Invalid IP address");
		}
	}
	
	$control_Validated4={
		#Pass the calling control and clear error message
		$errorprovider2.SetError($this, "");	
	}
	
	$txtIPSubnet_Validating4=[System.ComponentModel.CancelEventHandler]{
	#Event Argument: $_ = [System.ComponentModel.CancelEventArgs]
	
		$result = -not (Validate-IsIP $txtIPSubnet.Text)
	
		if($result -eq $true)
		{
			$_.Cancel = $true
			$errorprovider2.SetError($this, "Invalid IP address");
		}
	}
	
	$control_Validated5={
		#Pass the calling control and clear error message
		$errorprovider2.SetError($this, "");	
	}
	
	
	$txtIPGateway_Validating4=[System.ComponentModel.CancelEventHandler]{
	#Event Argument: $_ = [System.ComponentModel.CancelEventArgs]
	
		$result = -not (Validate-IsIP $txtIPGateway.Text)
	
		if($result -eq $true)
		{
			$_.Cancel = $true
			$errorprovider3.SetError($this, "Invalid IP address");
		}
	}
	
	$control_Validated6={
		#Pass the calling control and clear error message
		$errorprovider3.SetError($this, "");	
	}
	
	$textboxIP_Validating4=[System.ComponentModel.CancelEventHandler]{
	    #Event Argument: $_ = [System.ComponentModel.CancelEventArgs]
		$result = -not (Validate-IsIP $txtNLBDAGIP.Text)
	
		if($result -eq $true)
		{
			$_.Cancel = $true
			$errorprovider4.SetError($this, "Invalid IP address");
		}
	}
	
	$control_Validated7={
		#Pass the calling control and clear error message
		$errorprovider4.SetError($this, "");	
	}
	
	
	$txtNLBDAGIP_Validating5=[System.ComponentModel.CancelEventHandler]{
	    #Event Argument: $_ = [System.ComponentModel.CancelEventArgs]
		$result = -not (Validate-IsIP $txtNLBDAGIP.Text)
	
		if($result -eq $true)
		{
			$_.Cancel = $true
			$errorprovider4.SetError($this, "Invalid IP address");
		}
	}
	
	$control_Validated8={
		#Pass the calling control and clear error message
		$errorprovider4.SetError($this, "");	
	}
	
	$chkDAGReplication_CheckedChanged=
	{
		If ($chkDAGReplication.Checked)
		{
			$txtNLBDAGIP.Enabled = $false
		}
		Else
		{
			$txtNLBDAGIP.Enabled = $true
		}
	}
	
	$chkNLB_CheckedChanged={
		If ($chkNLB.Checked)
		{
			$txtNLBDAGIP.Enabled = $true
	        $txtArrayDAGURL.Enabled = $true
			$comboNLBClusterMode.Enabled = $true
			$txtNLBDAGIP.Text = "192.168.1.200"
			$comboNLBClusterMode.Text = "Multicast"
		}
		Else
		{
			$txtNLBDAGIP.Enabled = $false
			$txtNLBDAGIP.Text = ""
	        $txtArrayDAGURL.Enabled = $false
			$comboNLBClusterMode.Text = ""
			$comboNLBClusterMode.Enabled = $false
			
		}	
	}
	
	$chkUseInternalCASDomain_CheckedChanged={
		If ($chkUseInternalCASDomain.Checked)
		{
			$txtExternalDomainNameSpace.Enabled = $false        
		}
		Else
		{
			$txtExternalDomainNameSpace.Enabled = $true
		}
		
	}	# --End User Generated Script--
	#----------------------------------------------
	#region Generated Events
	#----------------------------------------------
	
	$Form_StateCorrection_Load=
	{
		#Correct the initial state of the form to prevent the .Net maximized form issue
		$frmMain.WindowState = $InitialFormWindowState
	}
	
	$Form_StoreValues_Closing=
	{
		#Store the control values
		$script:MainForm_textboxLoad = $textboxLoad.Text
		$script:MainForm_textboxSave = $textboxSave.Text
		$script:MainForm_txtOutputDir = $txtOutputDir.Text
		$script:MainForm_txtInternalDomain = $txtInternalDomain.Text
		$script:MainForm_txtSiteName = $txtSiteName.Text
		$script:MainForm_comboOS = $comboOS.Text
		$script:MainForm_txtServerName = $txtServerName.Text
		$script:MainForm_txtRegion = $txtRegion.Text
		$script:MainForm_comboExchRole = $comboExchRole.Text
		$script:MainForm_comboExchVer = $comboExchVer.Text
		$script:MainForm_txtArrayDAGURL = $txtArrayDAGURL.Text
		$script:MainForm_txtAutodiscoverURL = $txtAutodiscoverURL.Text
		$script:MainForm_txtCASURL = $txtCASURL.Text
		$script:MainForm_comboNLBClusterMode = $comboNLBClusterMode.Text
		$script:MainForm_txtNLBDAGIP = $txtNLBDAGIP.Text
		$script:MainForm_chkSSLOffload = $chkSSLOffload.Checked
		$script:MainForm_txtIPGateway = $txtIPGateway.Text
		$script:MainForm_txtIPSubnet = $txtIPSubnet.Text
		$script:MainForm_txtIPAddress = $txtIPAddress.Text
		$script:MainForm_chkInternetFacing = $chkInternetFacing.Checked
		$script:MainForm_chkNLB = $chkNLB.Checked
		$script:MainForm_chkDAGReplication = $chkDAGReplication.Checked
		$script:MainForm_txtIntName = $txtIntName.Text
		$script:MainForm_chkGeneratePrerequisites = $chkGeneratePrerequisites.Checked
		$script:MainForm_chkGenerateMigrationTips = $chkGenerateMigrationTips.Checked
		$script:MainForm_txtExternalDomainNameSpace = $txtExternalDomainNameSpace.Text
		$script:MainForm_chkUseInternalCASDomain = $chkUseInternalCASDomain.Checked
		$script:MainForm_chkGenerateCASArrays = $chkGenerateCASArrays.Checked
		$script:MainForm_chkReverseProxy = $chkReverseProxy.Checked
		$script:MainForm_chkRedirection = $chkRedirection.Checked
		$script:MainForm_txtvdirlocation = $txtvdirlocation.Text
		$script:MainForm_chkGunSlinger = $chkGunSlinger.Checked
	}

	
	$Form_Cleanup_FormClosed=
	{
		#Remove all event handlers from the controls
		try
		{
			$btnLoad.remove_Click($btnLoad_Click)
			$buttonBrowse2.remove_Click($buttonBrowse2_Click2)
			$btnSave.remove_Click($btnSave_Click)
			$buttonBrowse.remove_Click($buttonBrowse_Click2)
			$buttonBrowseFolder.remove_Click($buttonBrowseFolder_Click)
			$btnAdd.remove_Click($btnAdd_Click)
			$comboExchRole.remove_SelectedValueChanged($comboExchRole_SelectedValueChanged)
			$comboExchVer.remove_SelectedIndexChanged($comboExchVer_SelectedIndexChanged)
			$txtNLBDAGIP.remove_Validating($txtNLBDAGIP_Validating5)
			$txtNLBDAGIP.remove_Validated($control_Validated8)
			$txtIPGateway.remove_Validating($txtIPGateway_Validating4)
			$txtIPGateway.remove_Validated($control_Validated6)
			$txtIPSubnet.remove_Validating($txtIPSubnet_Validating4)
			$txtIPSubnet.remove_Validated($control_Validated5)
			$txtIPAddress.remove_Validating($txtIPAddress_Validating2)
			$txtIPAddress.remove_Validated($control_Validated2)
			$chkNLB.remove_CheckedChanged($chkNLB_CheckedChanged)
			$chkDAGReplication.remove_CheckedChanged($chkDAGReplication_CheckedChanged)
			$chkUseInternalCASDomain.remove_CheckedChanged($chkUseInternalCASDomain_CheckedChanged)
			$chkGunSlinger.remove_CheckedChanged($chkGunSlinger_CheckedChanged)
			$btnGenerate.remove_Click($btnGenerate_Click)
			$btnExit.remove_Click($btnExit_Click)
			$btnDelete.remove_Click($btnDelete_Click)
			$frmMain.remove_Load($OnLoadFormEvent)
			$frmMain.remove_Load($Form_StateCorrection_Load)
			$frmMain.remove_Closing($Form_StoreValues_Closing)
			$frmMain.remove_FormClosed($Form_Cleanup_FormClosed)
		}
		catch [Exception]
		{ }
	}
	#endregion Generated Events

	#----------------------------------------------
	#region Generated Form Code
	#----------------------------------------------
	#
	# frmMain
	#
	$frmMain.Controls.Add($btnLoad)
	$frmMain.Controls.Add($labelLoadConfig)
	$frmMain.Controls.Add($buttonBrowse2)
	$frmMain.Controls.Add($textboxLoad)
	$frmMain.Controls.Add($btnSave)
	$frmMain.Controls.Add($labelSaveConfig)
	$frmMain.Controls.Add($buttonBrowse)
	$frmMain.Controls.Add($textboxSave)
	$frmMain.Controls.Add($buttonBrowseFolder)
	$frmMain.Controls.Add($txtOutputDir)
	$frmMain.Controls.Add($tabcontrol1)
	$frmMain.Controls.Add($labelOutputFolder)
	$frmMain.Controls.Add($chkGunSlinger)
	$frmMain.Controls.Add($btnRun)
	$frmMain.Controls.Add($btnGenerate)
	$frmMain.Controls.Add($btnExit)
	$frmMain.Controls.Add($groupboxServerList)
	$frmMain.ClientSize = '1040, 459'
	$frmMain.Name = "frmMain"
	$frmMain.StartPosition = 'CenterScreen'
	$frmMain.Text = "Exchange Quick Configurator (beta)"
	$frmMain.add_Load($OnLoadFormEvent)
	#
	# btnLoad
	#
	$btnLoad.Location = '929, 331'
	$btnLoad.Name = "btnLoad"
	$btnLoad.Size = '93, 23'
	$btnLoad.TabIndex = 36
	$btnLoad.TabStop = $False
	$btnLoad.Text = "Load"
	$btnLoad.UseVisualStyleBackColor = $True
	$btnLoad.add_Click($btnLoad_Click)
	#
	# labelLoadConfig
	#
	$labelLoadConfig.Font = "Microsoft Sans Serif, 7.8pt"
	$labelLoadConfig.Location = '537, 336'
	$labelLoadConfig.Name = "labelLoadConfig"
	$labelLoadConfig.Size = '89, 19'
	$labelLoadConfig.TabIndex = 35
	$labelLoadConfig.Text = "Load Config:"
	#
	# buttonBrowse2
	#
	$buttonBrowse2.Location = '893, 331'
	$buttonBrowse2.Name = "buttonBrowse2"
	$buttonBrowse2.Size = '30, 23'
	$buttonBrowse2.TabIndex = 1
	$buttonBrowse2.TabStop = $False
	$buttonBrowse2.Text = "..."
	$buttonBrowse2.UseVisualStyleBackColor = $True
	$buttonBrowse2.add_Click($buttonBrowse2_Click2)
	#
	# textboxLoad
	#
	$textboxLoad.AutoCompleteMode = 'SuggestAppend'
	$textboxLoad.AutoCompleteSource = 'FileSystem'
	$textboxLoad.Location = '632, 333'
	$textboxLoad.Name = "textboxLoad"
	$textboxLoad.Size = '255, 22'
	$textboxLoad.TabIndex = 0
	$textboxLoad.TabStop = $False
	#
	# btnSave
	#
	$btnSave.Location = '929, 305'
	$btnSave.Name = "btnSave"
	$btnSave.Size = '93, 23'
	$btnSave.TabIndex = 34
	$btnSave.TabStop = $False
	$btnSave.Text = "Save"
	$btnSave.UseVisualStyleBackColor = $True
	$btnSave.add_Click($btnSave_Click)
	#
	# labelSaveConfig
	#
	$labelSaveConfig.Font = "Microsoft Sans Serif, 7.8pt"
	$labelSaveConfig.Location = '537, 310'
	$labelSaveConfig.Name = "labelSaveConfig"
	$labelSaveConfig.Size = '89, 15'
	$labelSaveConfig.TabIndex = 33
	$labelSaveConfig.Text = "Save Config:"
	#
	# buttonBrowse
	#
	$buttonBrowse.Location = '893, 305'
	$buttonBrowse.Name = "buttonBrowse"
	$buttonBrowse.Size = '30, 23'
	$buttonBrowse.TabIndex = 19
	$buttonBrowse.TabStop = $False
	$buttonBrowse.Text = "..."
	$buttonBrowse.UseVisualStyleBackColor = $True
	$buttonBrowse.add_Click($buttonBrowse_Click2)
	#
	# textboxSave
	#
	$textboxSave.AutoCompleteMode = 'SuggestAppend'
	$textboxSave.AutoCompleteSource = 'FileSystem'
	$textboxSave.Location = '632, 307'
	$textboxSave.Name = "textboxSave"
	$textboxSave.Size = '255, 22'
	$textboxSave.TabIndex = 0
	$textboxSave.TabStop = $False
	#
	# buttonBrowseFolder
	#
	$buttonBrowseFolder.Location = '893, 279'
	$buttonBrowseFolder.Name = "buttonBrowseFolder"
	$buttonBrowseFolder.Size = '30, 23'
	$buttonBrowseFolder.TabIndex = 18
	$buttonBrowseFolder.TabStop = $False
	$buttonBrowseFolder.Text = "..."
	$buttonBrowseFolder.UseVisualStyleBackColor = $True
	$buttonBrowseFolder.add_Click($buttonBrowseFolder_Click)
	#
	# txtOutputDir
	#
	$txtOutputDir.AutoCompleteMode = 'SuggestAppend'
	$txtOutputDir.AutoCompleteSource = 'FileSystemDirectories'
	$txtOutputDir.Location = '632, 281'
	$txtOutputDir.Name = "txtOutputDir"
	$txtOutputDir.Size = '255, 22'
	$txtOutputDir.TabIndex = 3
	$txtOutputDir.TabStop = $False
	#
	# tabcontrol1
	#
	$tabcontrol1.Controls.Add($tabpage1)
	$tabcontrol1.Controls.Add($tabpage2)
	$tabcontrol1.Location = '12, 12'
	$tabcontrol1.Name = "tabcontrol1"
	$tabcontrol1.SelectedIndex = 0
	$tabcontrol1.Size = '501, 438'
	$tabcontrol1.TabIndex = 0
	$tabcontrol1.TabStop = $False
	#
	# tabpage1
	#
	$tabpage1.Controls.Add($btnAdd)
	$tabpage1.Controls.Add($groupboxGeneral)
	$tabpage1.Controls.Add($groupboxURL)
	$tabpage1.Controls.Add($groupboxIPInfo)
	$tabpage1.Controls.Add($groupbox1)
	$tabpage1.BackColor = 'WhiteSmoke'
	$tabpage1.BorderStyle = 'FixedSingle'
	$tabpage1.Location = '4, 25'
	$tabpage1.Name = "tabpage1"
	$tabpage1.Padding = '3, 3, 3, 3'
	$tabpage1.Size = '493, 409'
	$tabpage1.TabIndex = 0
	$tabpage1.Text = "Add Servers"
	#
	# btnAdd
	#
	$btnAdd.Font = "Microsoft Sans Serif, 8.25pt"
	$btnAdd.Location = '396, 385'
	$btnAdd.Name = "btnAdd"
	$btnAdd.Size = '75, 23'
	$btnAdd.TabIndex = 3
	$btnAdd.Text = "Add"
	$btnAdd.UseVisualStyleBackColor = $True
	$btnAdd.add_Click($btnAdd_Click)
	#
	# groupboxGeneral
	#
	$groupboxGeneral.Controls.Add($labelDomain)
	$groupboxGeneral.Controls.Add($txtInternalDomain)
	$groupboxGeneral.Controls.Add($txtSiteName)
	$groupboxGeneral.Controls.Add($comboOS)
	$groupboxGeneral.Controls.Add($txtServerName)
	$groupboxGeneral.Controls.Add($txtRegion)
	$groupboxGeneral.Controls.Add($comboExchRole)
	$groupboxGeneral.Controls.Add($labelServerName)
	$groupboxGeneral.Controls.Add($comboExchVer)
	$groupboxGeneral.Controls.Add($labelOS)
	$groupboxGeneral.Controls.Add($labelVersion)
	$groupboxGeneral.Controls.Add($labelRegion)
	$groupboxGeneral.Controls.Add($labelSite)
	$groupboxGeneral.Controls.Add($labelRole)
	$groupboxGeneral.FlatStyle = 'Popup'
	$groupboxGeneral.Font = "Microsoft Sans Serif, 8.25pt, style=Underline"
	$groupboxGeneral.Location = '6, 10'
	$groupboxGeneral.Name = "groupboxGeneral"
	$groupboxGeneral.Size = '479, 118'
	$groupboxGeneral.TabIndex = 0
	$groupboxGeneral.TabStop = $False
	$groupboxGeneral.Text = "General"
	#
	# labelDomain
	#
	$labelDomain.Font = "Microsoft Sans Serif, 7.8pt"
	$labelDomain.Location = '5, 93'
	$labelDomain.Name = "labelDomain"
	$labelDomain.Size = '66, 17'
	$labelDomain.TabIndex = 26
	$labelDomain.Text = "Domain"
	#
	# txtInternalDomain
	#
	$txtInternalDomain.Font = "Microsoft Sans Serif, 8.25pt"
	$txtInternalDomain.Location = '100, 89'
	$txtInternalDomain.Name = "txtInternalDomain"
	$txtInternalDomain.Size = '149, 23'
	$txtInternalDomain.TabIndex = 3
	$txtInternalDomain.Text = "(ie. contoso.local)"
	$ttipGeneral.SetToolTip($txtInternalDomain, "Primarily used for EWS NLB bypass url configuration.")
	#
	# txtSiteName
	#
	$txtSiteName.Font = "Microsoft Sans Serif, 8.25pt"
	$txtSiteName.Location = '100, 36'
	$txtSiteName.Name = "txtSiteName"
	$txtSiteName.Size = '150, 23'
	$txtSiteName.TabIndex = 1
	$txtSiteName.Text = "(ie. Default-First-Site-Name)"
	$ttipGeneral.SetToolTip($txtSiteName, "Used to determine naming and deployment of CAS arrays 
(one per site is best practice in 2010, not really as important in 2012). 
This may be used later so you may as well fill this out accurately.")
	#
	# comboOS
	#
	$comboOS.DropDownStyle = 'DropDownList'
	$comboOS.FormattingEnabled = $True
	[void]$comboOS.Items.Add("Windows 2008 R2")
	[void]$comboOS.Items.Add("Windows 2012")
	$comboOS.Location = '328, 53'
	$comboOS.Name = "comboOS"
	$comboOS.Size = '145, 25'
	$comboOS.TabIndex = 5
	#
	# txtServerName
	#
	$txtServerName.Font = "Microsoft Sans Serif, 8.25pt"
	$txtServerName.Location = '100, 10'
	$txtServerName.Name = "txtServerName"
	$txtServerName.Size = '149, 23'
	$txtServerName.TabIndex = 0
	$txtServerName.Text = "(ie. CAS1)"
	$ttipGeneral.SetToolTip($txtServerName, "NetBIOS (non-FQDN) name of the server.")
	#
	# txtRegion
	#
	$txtRegion.Font = "Microsoft Sans Serif, 8.25pt"
	$txtRegion.Location = '100, 61'
	$txtRegion.Name = "txtRegion"
	$txtRegion.Size = '149, 23'
	$txtRegion.TabIndex = 2
	$txtRegion.Text = "(ie. US)"
	$ttipGeneral.SetToolTip($txtRegion, "Used for documentation purposes only currently.")
	#
	# comboExchRole
	#
	$comboExchRole.DropDownStyle = 'DropDownList'
	$comboExchRole.Enabled = $False
	$comboExchRole.FormattingEnabled = $True
	[void]$comboExchRole.Items.Add("HUB/CAS")
	$comboExchRole.Location = '328, 83'
	$comboExchRole.Name = "comboExchRole"
	$comboExchRole.Size = '145, 25'
	$comboExchRole.TabIndex = 6
	$comboExchRole.add_SelectedValueChanged($comboExchRole_SelectedValueChanged)
	#
	# labelServerName
	#
	$labelServerName.Font = "Microsoft Sans Serif, 7.8pt"
	$labelServerName.Location = '3, 17'
	$labelServerName.Name = "labelServerName"
	$labelServerName.Size = '100, 15'
	$labelServerName.TabIndex = 3
	$labelServerName.Text = "Server Name:"
	#
	# comboExchVer
	#
	$comboExchVer.DropDownStyle = 'DropDownList'
	$comboExchVer.FormattingEnabled = $True
	[void]$comboExchVer.Items.Add("Exchange 2003")
	[void]$comboExchVer.Items.Add("Exchange 2007")
	[void]$comboExchVer.Items.Add("Exchange 2010")
	[void]$comboExchVer.Items.Add("Exchange 2012")
	$comboExchVer.Location = '328, 22'
	$comboExchVer.Name = "comboExchVer"
	$comboExchVer.Size = '145, 25'
	$comboExchVer.TabIndex = 4
	$comboExchVer.add_SelectedIndexChanged($comboExchVer_SelectedIndexChanged)
	#
	# labelOS
	#
	$labelOS.Font = "Microsoft Sans Serif, 7.8pt"
	$labelOS.Location = '281, 58'
	$labelOS.Name = "labelOS"
	$labelOS.Size = '41, 17'
	$labelOS.TabIndex = 24
	$labelOS.Text = "OS:"
	#
	# labelVersion
	#
	$labelVersion.Font = "Microsoft Sans Serif, 7.8pt"
	$labelVersion.Location = '258, 22'
	$labelVersion.Name = "labelVersion"
	$labelVersion.Size = '64, 25'
	$labelVersion.TabIndex = 22
	$labelVersion.Text = "Version:"
	#
	# labelRegion
	#
	$labelRegion.Font = "Microsoft Sans Serif, 7.8pt"
	$labelRegion.Location = '4, 61'
	$labelRegion.Name = "labelRegion"
	$labelRegion.Size = '75, 17'
	$labelRegion.TabIndex = 22
	$labelRegion.Text = "Region:"
	#
	# labelSite
	#
	$labelSite.Font = "Microsoft Sans Serif, 7.8pt"
	$labelSite.Location = '3, 39'
	$labelSite.Name = "labelSite"
	$labelSite.Size = '68, 17'
	$labelSite.TabIndex = 9
	$labelSite.Text = "Site:"
	#
	# labelRole
	#
	$labelRole.Font = "Microsoft Sans Serif, 7.8pt"
	$labelRole.Location = '281, 83'
	$labelRole.Name = "labelRole"
	$labelRole.Size = '44, 25'
	$labelRole.TabIndex = 0
	$labelRole.Text = "Role:"
	#
	# groupboxURL
	#
	$groupboxURL.Controls.Add($txtArrayDAGURL)
	$groupboxURL.Controls.Add($txtAutodiscoverURL)
	$groupboxURL.Controls.Add($txtCASURL)
	$groupboxURL.Controls.Add($labelAccessURL)
	$groupboxURL.Controls.Add($labelAutodiscoverURL)
	$groupboxURL.Controls.Add($labelArrayDAGURL)
	$groupboxURL.Font = "Microsoft Sans Serif, 8.25pt, style=Underline"
	$groupboxURL.Location = '6, 280'
	$groupboxURL.Name = "groupboxURL"
	$groupboxURL.Size = '479, 103'
	$groupboxURL.TabIndex = 2
	$groupboxURL.TabStop = $False
	$groupboxURL.Text = "URL Information"
	#
	# txtArrayDAGURL
	#
	$txtArrayDAGURL.Enabled = $False
	$txtArrayDAGURL.Font = "Microsoft Sans Serif, 8.25pt"
	$txtArrayDAGURL.Location = '124, 74'
	$txtArrayDAGURL.Name = "txtArrayDAGURL"
	$txtArrayDAGURL.Size = '150, 23'
	$txtArrayDAGURL.TabIndex = 17
	$txtArrayDAGURL.Text = "(ie. owa.contoso.com)"
	#
	# txtAutodiscoverURL
	#
	$txtAutodiscoverURL.Enabled = $False
	$txtAutodiscoverURL.Font = "Microsoft Sans Serif, 8.25pt"
	$txtAutodiscoverURL.Location = '125, 46'
	$txtAutodiscoverURL.Name = "txtAutodiscoverURL"
	$txtAutodiscoverURL.Size = '150, 23'
	$txtAutodiscoverURL.TabIndex = 16
	$txtAutodiscoverURL.Text = "(ie. autodiscover.contoso.com)"
	#
	# txtCASURL
	#
	$txtCASURL.Enabled = $False
	$txtCASURL.Font = "Microsoft Sans Serif, 8.25pt"
	$txtCASURL.Location = '125, 22'
	$txtCASURL.Name = "txtCASURL"
	$txtCASURL.Size = '150, 23'
	$txtCASURL.TabIndex = 15
	$txtCASURL.Text = "(ie. owa.contoso.com)"
	$ttipGeneral.SetToolTip($txtCASURL, "CAS Access URL (for 2007/2003 servers this is the legacy url)")
	#
	# labelAccessURL
	#
	$labelAccessURL.Font = "Microsoft Sans Serif, 7.8pt"
	$labelAccessURL.Location = '5, 30'
	$labelAccessURL.Name = "labelAccessURL"
	$labelAccessURL.Size = '89, 15'
	$labelAccessURL.TabIndex = 6
	$labelAccessURL.Text = "Access URL:"
	#
	# labelAutodiscoverURL
	#
	$labelAutodiscoverURL.Font = "Microsoft Sans Serif, 7.8pt"
	$labelAutodiscoverURL.Location = '5, 52'
	$labelAutodiscoverURL.Name = "labelAutodiscoverURL"
	$labelAutodiscoverURL.Size = '118, 15'
	$labelAutodiscoverURL.TabIndex = 7
	$labelAutodiscoverURL.Text = "Autodiscover URL:"
	#
	# labelArrayDAGURL
	#
	$labelArrayDAGURL.Font = "Microsoft Sans Serif, 7.8pt"
	$labelArrayDAGURL.Location = '7, 77'
	$labelArrayDAGURL.Name = "labelArrayDAGURL"
	$labelArrayDAGURL.Size = '107, 16'
	$labelArrayDAGURL.TabIndex = 8
	$labelArrayDAGURL.Text = "Array/DAG URL:"
	#
	# groupboxIPInfo
	#
	$groupboxIPInfo.Controls.Add($comboNLBClusterMode)
	$groupboxIPInfo.Controls.Add($txtNLBDAGIP)
	$groupboxIPInfo.Controls.Add($labelNLBDAGIP)
	$groupboxIPInfo.Controls.Add($chkSSLOffload)
	$groupboxIPInfo.Controls.Add($txtIPGateway)
	$groupboxIPInfo.Controls.Add($txtIPSubnet)
	$groupboxIPInfo.Controls.Add($labelSSLOffload)
	$groupboxIPInfo.Controls.Add($txtIPAddress)
	$groupboxIPInfo.Controls.Add($labelInterfaceName)
	$groupboxIPInfo.Controls.Add($chkInternetFacing)
	$groupboxIPInfo.Controls.Add($labelGateway)
	$groupboxIPInfo.Controls.Add($labelInternetFacing)
	$groupboxIPInfo.Controls.Add($labelIPAddress)
	$groupboxIPInfo.Controls.Add($chkNLB)
	$groupboxIPInfo.Controls.Add($labelSubnet)
	$groupboxIPInfo.Controls.Add($labelWindowsNLB)
	$groupboxIPInfo.Controls.Add($labelReplicationNetwork)
	$groupboxIPInfo.Controls.Add($chkDAGReplication)
	$groupboxIPInfo.Controls.Add($txtIntName)
	$groupboxIPInfo.Font = "Microsoft Sans Serif, 8.25pt, style=Underline"
	$groupboxIPInfo.Location = '6, 134'
	$groupboxIPInfo.Name = "groupboxIPInfo"
	$groupboxIPInfo.Size = '479, 145'
	$groupboxIPInfo.TabIndex = 1
	$groupboxIPInfo.TabStop = $False
	$groupboxIPInfo.Text = "IP Address Info"
	#
	# comboNLBClusterMode
	#
	$comboNLBClusterMode.Enabled = $False
	$comboNLBClusterMode.FormattingEnabled = $True
	[void]$comboNLBClusterMode.Items.Add("Multicast")
	[void]$comboNLBClusterMode.Items.Add("Unicast")
	$comboNLBClusterMode.Location = '301, 110'
	$comboNLBClusterMode.Name = "comboNLBClusterMode"
	$comboNLBClusterMode.Size = '121, 25'
	$comboNLBClusterMode.TabIndex = 38
	#
	# txtNLBDAGIP
	#
	$txtNLBDAGIP.Anchor = 'Top, Left, Right'
	$txtNLBDAGIP.Enabled = $False
	$txtNLBDAGIP.Font = "Microsoft Sans Serif, 8.25pt"
	$errorprovider4.SetIconAlignment($txtNLBDAGIP, 'MiddleLeft')
	$errorprovider4.SetIconPadding($txtNLBDAGIP, 2)
	$txtNLBDAGIP.Location = '123, 109'
	$txtNLBDAGIP.MaxLength = 200
	$txtNLBDAGIP.Name = "txtNLBDAGIP"
	$txtNLBDAGIP.Size = '161, 23'
	$txtNLBDAGIP.TabIndex = 10
	$ttipGeneral.SetToolTip($txtNLBDAGIP, "For NLB: This is just the shared NLB IP you are going to use.
This (obviously) is only applicable per site.

For DAG: This is the DAG IP (on the MAPI network) you will
be utilizing. Note that you need one of these per site.

If you are setting up an all-in-one role server this is the DAG IP
(as you cannot simultaneously have NLB and MSCS running)")
	$txtNLBDAGIP.add_Validating($txtNLBDAGIP_Validating5)
	$txtNLBDAGIP.add_Validated($control_Validated8)
	#
	# labelNLBDAGIP
	#
	$labelNLBDAGIP.Font = "Microsoft Sans Serif, 7.8pt"
	$labelNLBDAGIP.Location = '6, 112'
	$labelNLBDAGIP.Name = "labelNLBDAGIP"
	$labelNLBDAGIP.Size = '106, 20'
	$labelNLBDAGIP.TabIndex = 37
	$labelNLBDAGIP.Text = "NLB/DAG IP:"
	#
	# chkSSLOffload
	#
	$chkSSLOffload.Enabled = $False
	$chkSSLOffload.Location = '444, 38'
	$chkSSLOffload.Name = "chkSSLOffload"
	$chkSSLOffload.Size = '18, 24'
	$chkSSLOffload.TabIndex = 12
	$ttipGeneral.SetToolTip($chkSSLOffload, "Not implemented (yet!)")
	$chkSSLOffload.UseVisualStyleBackColor = $True
	#
	# txtIPGateway
	#
	$txtIPGateway.Anchor = 'Top, Left, Right'
	$txtIPGateway.Font = "Microsoft Sans Serif, 8.25pt"
	$errorprovider3.SetIconAlignment($txtIPGateway, 'MiddleLeft')
	$errorprovider3.SetIconPadding($txtIPGateway, 2)
	$txtIPGateway.Location = '123, 64'
	$txtIPGateway.MaxLength = 200
	$txtIPGateway.Name = "txtIPGateway"
	$txtIPGateway.Size = '161, 23'
	$txtIPGateway.TabIndex = 8
	$txtIPGateway.Text = "192.168.1.1"
	$ttipGeneral.SetToolTip($txtIPGateway, "The gateway for the network interface being defined.
If this is a DAG replication network you still want to define
a gateway here, but leave the actual gateway on the NIC
empty.")
	$txtIPGateway.add_Validating($txtIPGateway_Validating4)
	$txtIPGateway.add_Validated($control_Validated6)
	#
	# txtIPSubnet
	#
	$txtIPSubnet.Anchor = 'Top, Left, Right'
	$txtIPSubnet.Font = "Microsoft Sans Serif, 8.25pt"
	$errorprovider2.SetIconAlignment($txtIPSubnet, 'MiddleLeft')
	$errorprovider2.SetIconPadding($txtIPSubnet, 2)
	$txtIPSubnet.Location = '123, 42'
	$txtIPSubnet.MaxLength = 200
	$txtIPSubnet.Name = "txtIPSubnet"
	$txtIPSubnet.Size = '161, 23'
	$txtIPSubnet.TabIndex = 7
	$txtIPSubnet.Text = "255.255.255.0"
	$txtIPSubnet.add_Validating($txtIPSubnet_Validating4)
	$txtIPSubnet.add_Validated($control_Validated5)
	#
	# labelSSLOffload
	#
	$labelSSLOffload.Font = "Microsoft Sans Serif, 7.8pt"
	$labelSSLOffload.Location = '301, 42'
	$labelSSLOffload.Name = "labelSSLOffload"
	$labelSSLOffload.Size = '108, 14'
	$labelSSLOffload.TabIndex = 32
	$labelSSLOffload.Text = "SSL Offload?"
	#
	# txtIPAddress
	#
	$txtIPAddress.Anchor = 'Top, Left, Right'
	$txtIPAddress.Font = "Microsoft Sans Serif, 8.25pt"
	$errorprovider1.SetIconAlignment($txtIPAddress, 'MiddleLeft')
	$errorprovider1.SetIconPadding($txtIPAddress, 2)
	$txtIPAddress.Location = '123, 19'
	$txtIPAddress.MaxLength = 200
	$txtIPAddress.Name = "txtIPAddress"
	$txtIPAddress.Size = '161, 23'
	$txtIPAddress.TabIndex = 6
	$txtIPAddress.Text = "192.168.1.100"
	$txtIPAddress.add_Validating($txtIPAddress_Validating2)
	$txtIPAddress.add_Validated($control_Validated2)
	#
	# labelInterfaceName
	#
	$labelInterfaceName.Font = "Microsoft Sans Serif, 7.8pt"
	$labelInterfaceName.Location = '6, 91'
	$labelInterfaceName.Name = "labelInterfaceName"
	$labelInterfaceName.Size = '106, 21'
	$labelInterfaceName.TabIndex = 2
	$labelInterfaceName.Text = "Interface Name:"
	#
	# chkInternetFacing
	#
	$chkInternetFacing.Enabled = $False
	$chkInternetFacing.Location = '444, 60'
	$chkInternetFacing.Name = "chkInternetFacing"
	$chkInternetFacing.Size = '18, 24'
	$chkInternetFacing.TabIndex = 13
	$ttipGeneral.SetToolTip($chkInternetFacing, "If this is externally facing and you are using a reverse proxy
(as you should be doing) then the authentication for owa/ecp
might need to be changed and this should be selected. Otherwise
if this is used for internal access only then leave this unselected.

External Example:
Internet <-> Reverse Proxy <-> External CAS NIC <->
CAS Server. 

Internal Example:
Internal Network <-> Internal CAS NIC <-> CAS Server. 
")
	$chkInternetFacing.UseVisualStyleBackColor = $True
	#
	# labelGateway
	#
	$labelGateway.Font = "Microsoft Sans Serif, 7.8pt"
	$labelGateway.Location = '6, 67'
	$labelGateway.Name = "labelGateway"
	$labelGateway.Size = '76, 17'
	$labelGateway.TabIndex = 36
	$labelGateway.Text = "Gateway:"
	#
	# labelInternetFacing
	#
	$labelInternetFacing.Font = "Microsoft Sans Serif, 7.8pt"
	$labelInternetFacing.Location = '301, 64'
	$labelInternetFacing.Name = "labelInternetFacing"
	$labelInternetFacing.Size = '132, 14'
	$labelInternetFacing.TabIndex = 4
	$labelInternetFacing.Text = "Internet Facing?"
	#
	# labelIPAddress
	#
	$labelIPAddress.Font = "Microsoft Sans Serif, 7.8pt"
	$labelIPAddress.Location = '6, 25'
	$labelIPAddress.Name = "labelIPAddress"
	$labelIPAddress.Size = '76, 17'
	$labelIPAddress.TabIndex = 1
	$labelIPAddress.Text = "IP Address:"
	#
	# chkNLB
	#
	$chkNLB.Enabled = $False
	$chkNLB.Location = '444, 84'
	$chkNLB.Name = "chkNLB"
	$chkNLB.Size = '18, 24'
	$chkNLB.TabIndex = 14
	$ttipGeneral.SetToolTip($chkNLB, "This is ""usually"" your internal CAS getting Windows Network Load Balanced.
If using a hardware load balancer don't select this")
	$chkNLB.UseVisualStyleBackColor = $True
	$chkNLB.add_CheckedChanged($chkNLB_CheckedChanged)
	#
	# labelSubnet
	#
	$labelSubnet.Font = "Microsoft Sans Serif, 7.8pt"
	$labelSubnet.Location = '6, 42'
	$labelSubnet.Name = "labelSubnet"
	$labelSubnet.Size = '76, 20'
	$labelSubnet.TabIndex = 24
	$labelSubnet.Text = "Subnet:"
	#
	# labelWindowsNLB
	#
	$labelWindowsNLB.Font = "Microsoft Sans Serif, 7.8pt"
	$labelWindowsNLB.Location = '301, 88'
	$labelWindowsNLB.Name = "labelWindowsNLB"
	$labelWindowsNLB.Size = '132, 14'
	$labelWindowsNLB.TabIndex = 5
	$labelWindowsNLB.Text = "Windows NLB?"
	#
	# labelReplicationNetwork
	#
	$labelReplicationNetwork.Font = "Microsoft Sans Serif, 7.8pt"
	$labelReplicationNetwork.Location = '301, 16'
	$labelReplicationNetwork.Name = "labelReplicationNetwork"
	$labelReplicationNetwork.Size = '132, 14'
	$labelReplicationNetwork.TabIndex = 33
	$labelReplicationNetwork.Text = "Replication Network?"
	#
	# chkDAGReplication
	#
	$chkDAGReplication.Enabled = $False
	$chkDAGReplication.Location = '444, 14'
	$chkDAGReplication.Name = "chkDAGReplication"
	$chkDAGReplication.Size = '18, 24'
	$chkDAGReplication.TabIndex = 11
	$ttipGeneral.SetToolTip($chkDAGReplication, "If you are defining a DAG replication interface then select
this. Note that in any DAG a second dedicated replicastion
VLAN is best practice. In a multiple site DAG the same best
practice should be followed and each of these dedicated
VLANs should be routable to one another.")
	$chkDAGReplication.UseVisualStyleBackColor = $True
	$chkDAGReplication.add_CheckedChanged($chkDAGReplication_CheckedChanged)
	#
	# txtIntName
	#
	$txtIntName.Font = "Microsoft Sans Serif, 8.25pt"
	$txtIntName.Location = '123, 87'
	$txtIntName.Name = "txtIntName"
	$txtIntName.Size = '161, 23'
	$txtIntName.TabIndex = 9
	$txtIntName.Text = "(ie. INTERNAL)"
	$ttipGeneral.SetToolTip($txtIntName, "The literal NIC interface name.")
	#
	# groupbox1
	#
	$groupbox1.Controls.Add($label1)
	$groupbox1.Controls.Add($label2)
	$groupbox1.Controls.Add($label3)
	$groupbox1.Controls.Add($label4)
	$groupbox1.Controls.Add($label5)
	$groupbox1.Controls.Add($label6)
	$groupbox1.Controls.Add($label7)
	$groupbox1.FlatStyle = 'Popup'
	$groupbox1.Font = "Microsoft Sans Serif, 8.25pt, style=Underline"
	$groupbox1.Location = '6, 10'
	$groupbox1.Name = "groupbox1"
	$groupbox1.Size = '468, 118'
	$groupbox1.TabIndex = 0
	$groupbox1.TabStop = $False
	$groupbox1.Text = "General"
	#
	# label1
	#
	$label1.Font = "Microsoft Sans Serif, 7.8pt"
	$label1.Location = '5, 93'
	$label1.Name = "label1"
	$label1.Size = '118, 17'
	$label1.TabIndex = 26
	$label1.Text = "Internal Domain:"
	#
	# label2
	#
	$label2.Font = "Microsoft Sans Serif, 7.8pt"
	$label2.Location = '3, 17'
	$label2.Name = "label2"
	$label2.Size = '89, 15'
	$label2.TabIndex = 3
	$label2.Text = "Server Name:"
	#
	# label3
	#
	$label3.Font = "Microsoft Sans Serif, 7.8pt"
	$label3.Location = '280, 17'
	$label3.Name = "label3"
	$label3.Size = '53, 17'
	$label3.TabIndex = 24
	$label3.Text = "OS:"
	#
	# label4
	#
	$label4.Font = "Microsoft Sans Serif, 7.8pt"
	$label4.Location = '4, 65'
	$label4.Name = "label4"
	$label4.Size = '118, 17'
	$label4.TabIndex = 22
	$label4.Text = "Exchange Version:"
	#
	# label5
	#
	$label5.Font = "Microsoft Sans Serif, 7.8pt"
	$label5.Location = '279, 39'
	$label5.Name = "label5"
	$label5.Size = '53, 17'
	$label5.TabIndex = 22
	$label5.Text = "Region:"
	#
	# label6
	#
	$label6.Font = "Microsoft Sans Serif, 7.8pt"
	$label6.Location = '3, 39'
	$label6.Name = "label6"
	$label6.Size = '46, 17'
	$label6.TabIndex = 9
	$label6.Text = "Site:"
	#
	# label7
	#
	$label7.Font = "Microsoft Sans Serif, 7.8pt"
	$label7.Location = '280, 75'
	$label7.Name = "label7"
	$label7.Size = '49, 15'
	$label7.TabIndex = 0
	$label7.Text = "Role:"
	#
	# tabpage2
	#
	$tabpage2.Controls.Add($labelGeneratePrerequisite)
	$tabpage2.Controls.Add($chkGeneratePrerequisites)
	$tabpage2.Controls.Add($labelGenerateMigrationCoe)
	$tabpage2.Controls.Add($chkGenerateMigrationTips)
	$tabpage2.Controls.Add($labelExternalDNSNamespace)
	$tabpage2.Controls.Add($txtExternalDomainNameSpace)
	$tabpage2.Controls.Add($labelUseInternalCASArrayD)
	$tabpage2.Controls.Add($chkUseInternalCASDomain)
	$tabpage2.Controls.Add($labelGenerateCASArrays)
	$tabpage2.Controls.Add($chkGenerateCASArrays)
	$tabpage2.Controls.Add($labelAssumeReverseProxy)
	$tabpage2.Controls.Add($chkReverseProxy)
	$tabpage2.Controls.Add($labelRedirectURLs)
	$tabpage2.Controls.Add($chkRedirection)
	$tabpage2.Controls.Add($labelInternalExchangeVirt)
	$tabpage2.Controls.Add($txtvdirlocation)
	$tabpage2.BackColor = 'WhiteSmoke'
	$tabpage2.BorderStyle = 'FixedSingle'
	$tabpage2.Location = '4, 25'
	$tabpage2.Name = "tabpage2"
	$tabpage2.Padding = '3, 3, 3, 3'
	$tabpage2.Size = '493, 409'
	$tabpage2.TabIndex = 1
	$tabpage2.Text = "Options"
	#
	# labelGeneratePrerequisite
	#
	$labelGeneratePrerequisite.Font = "Microsoft Sans Serif, 7.8pt"
	$labelGeneratePrerequisite.Location = '19, 172'
	$labelGeneratePrerequisite.Name = "labelGeneratePrerequisite"
	$labelGeneratePrerequisite.Size = '262, 23'
	$labelGeneratePrerequisite.TabIndex = 15
	$labelGeneratePrerequisite.Text = "Generate Prerequisites:"
	#
	# chkGeneratePrerequisites
	#
	$chkGeneratePrerequisites.Checked = $True
	$chkGeneratePrerequisites.CheckState = 'Checked'
	$chkGeneratePrerequisites.Location = '287, 167'
	$chkGeneratePrerequisites.Name = "chkGeneratePrerequisites"
	$chkGeneratePrerequisites.Size = '17, 24'
	$chkGeneratePrerequisites.TabIndex = 14
	$chkGeneratePrerequisites.UseVisualStyleBackColor = $True
	#
	# labelGenerateMigrationCoe
	#
	$labelGenerateMigrationCoe.Font = "Microsoft Sans Serif, 7.8pt"
	$labelGenerateMigrationCoe.Location = '19, 140'
	$labelGenerateMigrationCoe.Name = "labelGenerateMigrationCoe"
	$labelGenerateMigrationCoe.Size = '262, 23'
	$labelGenerateMigrationCoe.TabIndex = 13
	$labelGenerateMigrationCoe.Text = "Generate Migration/Co-existence Tips:"
	#
	# chkGenerateMigrationTips
	#
	$chkGenerateMigrationTips.Checked = $True
	$chkGenerateMigrationTips.CheckState = 'Checked'
	$chkGenerateMigrationTips.Location = '287, 135'
	$chkGenerateMigrationTips.Name = "chkGenerateMigrationTips"
	$chkGenerateMigrationTips.Size = '17, 24'
	$chkGenerateMigrationTips.TabIndex = 12
	$chkGenerateMigrationTips.UseVisualStyleBackColor = $True
	#
	# labelExternalDNSNamespace
	#
	$labelExternalDNSNamespace.Font = "Microsoft Sans Serif, 7.8pt"
	$labelExternalDNSNamespace.Location = '19, 229'
	$labelExternalDNSNamespace.Name = "labelExternalDNSNamespace"
	$labelExternalDNSNamespace.Size = '262, 23'
	$labelExternalDNSNamespace.TabIndex = 11
	$labelExternalDNSNamespace.Text = "External DNS Namespace:"
	#
	# txtExternalDomainNameSpace
	#
	$txtExternalDomainNameSpace.Enabled = $False
	$txtExternalDomainNameSpace.Location = '287, 229'
	$txtExternalDomainNameSpace.Name = "txtExternalDomainNameSpace"
	$txtExternalDomainNameSpace.Size = '196, 22'
	$txtExternalDomainNameSpace.TabIndex = 10
	#
	# labelUseInternalCASArrayD
	#
	$labelUseInternalCASArrayD.Font = "Microsoft Sans Serif, 7.8pt"
	$labelUseInternalCASArrayD.Location = '19, 202'
	$labelUseInternalCASArrayD.Name = "labelUseInternalCASArrayD"
	$labelUseInternalCASArrayD.Size = '262, 23'
	$labelUseInternalCASArrayD.TabIndex = 9
	$labelUseInternalCASArrayD.Text = "Use Internal CAS Array Domain:"
	#
	# chkUseInternalCASDomain
	#
	$chkUseInternalCASDomain.Checked = $True
	$chkUseInternalCASDomain.CheckState = 'Checked'
	$chkUseInternalCASDomain.Location = '287, 197'
	$chkUseInternalCASDomain.Name = "chkUseInternalCASDomain"
	$chkUseInternalCASDomain.Size = '17, 24'
	$chkUseInternalCASDomain.TabIndex = 8
	$chkUseInternalCASDomain.UseVisualStyleBackColor = $True
	$chkUseInternalCASDomain.add_CheckedChanged($chkUseInternalCASDomain_CheckedChanged)
	#
	# labelGenerateCASArrays
	#
	$labelGenerateCASArrays.Font = "Microsoft Sans Serif, 7.8pt"
	$labelGenerateCASArrays.Location = '19, 111'
	$labelGenerateCASArrays.Name = "labelGenerateCASArrays"
	$labelGenerateCASArrays.Size = '262, 23'
	$labelGenerateCASArrays.TabIndex = 7
	$labelGenerateCASArrays.Text = "Generate CAS Arrays:"
	#
	# chkGenerateCASArrays
	#
	$chkGenerateCASArrays.Checked = $True
	$chkGenerateCASArrays.CheckState = 'Checked'
	$chkGenerateCASArrays.Location = '287, 106'
	$chkGenerateCASArrays.Name = "chkGenerateCASArrays"
	$chkGenerateCASArrays.Size = '17, 24'
	$chkGenerateCASArrays.TabIndex = 6
	$chkGenerateCASArrays.UseVisualStyleBackColor = $True
	#
	# labelAssumeReverseProxy
	#
	$labelAssumeReverseProxy.Font = "Microsoft Sans Serif, 7.8pt"
	$labelAssumeReverseProxy.Location = '19, 82'
	$labelAssumeReverseProxy.Name = "labelAssumeReverseProxy"
	$labelAssumeReverseProxy.Size = '268, 23'
	$labelAssumeReverseProxy.TabIndex = 5
	$labelAssumeReverseProxy.Text = "Assume Reverse Proxy:"
	#
	# chkReverseProxy
	#
	$chkReverseProxy.Checked = $True
	$chkReverseProxy.CheckState = 'Checked'
	$chkReverseProxy.Location = '287, 75'
	$chkReverseProxy.Name = "chkReverseProxy"
	$chkReverseProxy.Size = '17, 24'
	$chkReverseProxy.TabIndex = 4
	$chkReverseProxy.UseVisualStyleBackColor = $True
	#
	# labelRedirectURLs
	#
	$labelRedirectURLs.Font = "Microsoft Sans Serif, 7.8pt"
	$labelRedirectURLs.Location = '19, 50'
	$labelRedirectURLs.Name = "labelRedirectURLs"
	$labelRedirectURLs.Size = '268, 23'
	$labelRedirectURLs.TabIndex = 3
	$labelRedirectURLs.Text = "Redirect URLs:"
	#
	# chkRedirection
	#
	$chkRedirection.Checked = $True
	$chkRedirection.CheckState = 'Checked'
	$chkRedirection.Location = '287, 45'
	$chkRedirection.Name = "chkRedirection"
	$chkRedirection.Size = '17, 24'
	$chkRedirection.TabIndex = 2
	$chkRedirection.UseVisualStyleBackColor = $True
	#
	# labelInternalExchangeVirt
	#
	$labelInternalExchangeVirt.Font = "Microsoft Sans Serif, 7.8pt"
	$labelInternalExchangeVirt.Location = '19, 22'
	$labelInternalExchangeVirt.Name = "labelInternalExchangeVirt"
	$labelInternalExchangeVirt.Size = '268, 23'
	$labelInternalExchangeVirt.TabIndex = 1
	$labelInternalExchangeVirt.Text = "Internal Exchange Virtual Directory Location:"
	#
	# txtvdirlocation
	#
	$txtvdirlocation.Location = '287, 19'
	$txtvdirlocation.Name = "txtvdirlocation"
	$txtvdirlocation.Size = '196, 22'
	$txtvdirlocation.TabIndex = 0
	$txtvdirlocation.Text = "c:\inetpub\exchange-internal"
	#
	# labelOutputFolder
	#
	$labelOutputFolder.Font = "Microsoft Sans Serif, 7.8pt"
	$labelOutputFolder.Location = '537, 284'
	$labelOutputFolder.Name = "labelOutputFolder"
	$labelOutputFolder.Size = '89, 15'
	$labelOutputFolder.TabIndex = 31
	$labelOutputFolder.Text = "Output Folder:"
	#
	# chkGunSlinger
	#
	$chkGunSlinger.Location = '537, 364'
	$chkGunSlinger.Name = "chkGunSlinger"
	$chkGunSlinger.Size = '404, 77'
	$chkGunSlinger.TabIndex = 20
	$chkGunSlinger.TabStop = $False
	$chkGunSlinger.Text = "By checking this box you are agreeing that you alone are responsible for your Exchange environment and are patently ignorning the fact that this script is BETA and may screw things up."
	$chkGunSlinger.UseVisualStyleBackColor = $True
	$chkGunSlinger.Visible = $False
	$chkGunSlinger.add_CheckedChanged($chkGunSlinger_CheckedChanged)
	#
	# btnRun
	#
	$btnRun.Enabled = $False
	$btnRun.Location = '947, 371'
	$btnRun.Name = "btnRun"
	$btnRun.Size = '75, 23'
	$btnRun.TabIndex = 21
	$btnRun.TabStop = $False
	$btnRun.Text = "Run Script"
	$btnRun.UseVisualStyleBackColor = $True
	$btnRun.Visible = $False
	#
	# btnGenerate
	#
	$btnGenerate.Location = '929, 279'
	$btnGenerate.Name = "btnGenerate"
	$btnGenerate.Size = '93, 23'
	$btnGenerate.TabIndex = 19
	$btnGenerate.TabStop = $False
	$btnGenerate.Text = "Generate!"
	$btnGenerate.UseVisualStyleBackColor = $True
	$btnGenerate.add_Click($btnGenerate_Click)
	#
	# btnExit
	#
	$btnExit.Location = '947, 398'
	$btnExit.Name = "btnExit"
	$btnExit.Size = '75, 23'
	$btnExit.TabIndex = 22
	$btnExit.TabStop = $False
	$btnExit.Text = "Exit"
	$btnExit.UseVisualStyleBackColor = $True
	$btnExit.add_Click($btnExit_Click)
	#
	# groupboxServerList
	#
	$groupboxServerList.Controls.Add($datagridServers)
	$groupboxServerList.Controls.Add($btnDelete)
	$groupboxServerList.Font = "Microsoft Sans Serif, 8.25pt, style=Underline"
	$groupboxServerList.Location = '527, 27'
	$groupboxServerList.Name = "groupboxServerList"
	$groupboxServerList.Size = '501, 247'
	$groupboxServerList.TabIndex = 24
	$groupboxServerList.TabStop = $False
	$groupboxServerList.Text = "Server List"
	#
	# datagridServers
	#
	$datagridServers.AllowUserToAddRows = $False
	$datagridServers.ColumnHeadersHeightSizeMode = 'AutoSize'
	$datagridServers.Location = '10, 19'
	$datagridServers.MultiSelect = $False
	$datagridServers.Name = "datagridServers"
	$datagridServers.Size = '485, 191'
	$datagridServers.TabIndex = 2
	$datagridServers.TabStop = $False
	#
	# btnDelete
	#
	$btnDelete.Font = "Microsoft Sans Serif, 8.25pt"
	$btnDelete.Location = '402, 216'
	$btnDelete.Name = "btnDelete"
	$btnDelete.Size = '93, 23'
	$btnDelete.TabIndex = 16
	$btnDelete.TabStop = $False
	$btnDelete.Text = "Delete"
	$btnDelete.UseVisualStyleBackColor = $True
	$btnDelete.add_Click($btnDelete_Click)
	#
	# errorprovider1
	#
	$errorprovider1.ContainerControl = $frmMain
	#
	# folderbrowserdialog1
	#
	#
	# openfiledialogSave
	#
	$openfiledialogSave.CheckFileExists = $False
	$openfiledialogSave.DefaultExt = "csv"
	$openfiledialogSave.Filter = "Comma Separated Value File (.csv)|*.csv"
	$openfiledialogSave.ShowHelp = $True
	#
	# openfiledialogLoad
	#
	$openfiledialogLoad.DefaultExt = "csv"
	$openfiledialogLoad.Filter = "Comma Separated Value File (.csv)|*.csv"
	$openfiledialogLoad.ShowHelp = $True
	#
	# errorprovider2
	#
	$errorprovider2.ContainerControl = $frmMain
	#
	# errorprovider3
	#
	$errorprovider3.ContainerControl = $frmMain
	#
	# errorprovider4
	#
	$errorprovider4.ContainerControl = $frmMain
	#
	# ttipGeneral
	#
	$ttipGeneral.AutomaticDelay = 0
	$ttipGeneral.AutoPopDelay = 100000
	$ttipGeneral.InitialDelay = 500
	$ttipGeneral.ReshowDelay = 100
	$ttipGeneral.ShowAlways = $True
	#endregion Generated Form Code

	#----------------------------------------------

	#Save the initial state of the form
	$InitialFormWindowState = $frmMain.WindowState
	#Init the OnLoad event to correct the initial state of the form
	$frmMain.add_Load($Form_StateCorrection_Load)
	#Clean up the control events
	$frmMain.add_FormClosed($Form_Cleanup_FormClosed)
	#Store the control values when form is closing
	$frmMain.add_Closing($Form_StoreValues_Closing)
	#Show the Form
	return $frmMain.ShowDialog()

}
#endregion Source: MainForm.pff

#region Source: Globals.ps1
	#region Global Variables
	$ServerList = New-Object System.Data.DataTable
	$Servers = @()
	
	# Exchange 2010 Variables
	$2010HubCASServers = @()
	$2010HubCASServersInternal = @()
	$2010HubCASServersExternal = @()
	$2010HubCASSets = @()
	$2010HubCASServersInternalNLB = @()
	$2010HubCASServersExternalNLB = @()
	$2010DatabaseServers = @()
	$2010AllInOneServers = @()
	$2010DatabaseDAGServers = @()
	$2010MailboxMAPIServers = @()
	$2010MailboxDAGServers = @()
	$2010EdgeServers = @()
	$2010NLBServers = @()
	$2012DatabaseDAGServers = @()
	
	# General Exchange versions in environment
	$2003Servers = @()
	$2007Servers = @()
	$2010Servers = @()
	$2012Servers = @()
	
	# Upgrade paths
	$2003to2010 = $false
	$2003to2012 = $false
	$2007to2010 = $false
	$2007to2012 = $false
	$2010to2012 = $false
	
	# OSes
	$2012Servers = @()
	$2008R2Servers = @()
	
	# Interfaces
	$ServersWithMultipleIPsOnOneInterface = @()
	
	# Options
	$ReverseProxy = $true
	$vDirLocation = "C:\inetpub\exchange-internal\"
	$Redirection = $true
	$GenerateCASArrays = $true
	$GenerateMigrationTips = $true
	$GeneratePrerequisties = $true
	$UseCASInternalDomainNameSpace = $true
	$ExternalNamespace = $txtExternalDomainNameSpace.Text
	#endregion Global Variables
	
	#region Script Output Statics
	# Some static script output
	$LoadScriptModules = @"
#region Module/Snapin/Dot Sourcing
if (! (Get-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction:SilentlyContinue) )
{ 
	Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 
}
Import-Module WebAdministration
#endregion Module/Snapin/Dot Sourcing
"@
	#endregion Script Output Statics
	
	#region General Functions
	function ConvertTo-ScriptBlock 
	{
		param
		(
			[parameter(ValueFromPipeline=$true,Position=0)]
			[string]
			$string
		)
		$sb = [scriptblock]::Create($string)
		return $sb
	}
	
	function Get-Type 
	{ 
	    param($type) 
	 
	$types = @( 
	'System.Boolean', 
	'System.Byte[]', 
	'System.Byte', 
	'System.Char', 
	'System.Datetime', 
	'System.Decimal', 
	'System.Double', 
	'System.Guid', 
	'System.Int16', 
	'System.Int32', 
	'System.Int64', 
	'System.Single', 
	'System.UInt16', 
	'System.UInt32', 
	'System.UInt64') 
	 
	    if ( $types -contains $type ) { 
	        Write-Output "$type" 
	    } 
	    else { 
	        Write-Output 'System.String' 
	         
	    } 
	} #Get-Type 
	 
	####################### 
	<# 
	.SYNOPSIS 
	Creates a DataTable for an object 
	.DESCRIPTION 
	Creates a DataTable based on an objects properties. 
	.INPUTS 
	Object 
	    Any object can be piped to Out-DataTable 
	.OUTPUTS 
	   System.Data.DataTable 
	.EXAMPLE 
	$dt = Get-psdrive| Out-DataTable 
	This example creates a DataTable from the properties of Get-psdrive and assigns output to $dt variable 
	.NOTES 
	Adapted from script by Marc van Orsouw see link 
	Version History 
	v1.0  - Chad Miller - Initial Release 
	v1.1  - Chad Miller - Fixed Issue with Properties 
	v1.2  - Chad Miller - Added setting column datatype by property as suggested by emp0 
	v1.3  - Chad Miller - Corrected issue with setting datatype on empty properties 
	v1.4  - Chad Miller - Corrected issue with DBNull 
	v1.5  - Chad Miller - Updated example 
	v1.6  - Chad Miller - Added column datatype logic with default to string 
	.LINK 
	http://thepowershellguy.com/blogs/posh/archive/2007/01/21/powershell-gui-scripblock-monitor-script.aspx 
	#> 
	function Out-DataTable 
	{ 
	    [CmdletBinding()] 
	    param([Parameter(Position=0, Mandatory=$true, ValueFromPipeline = $true)] [PSObject[]]$InputObject) 
	 
	    Begin 
	    { 
	        $dt = new-object Data.datatable   
	        $First = $true  
	    } 
	    Process 
	    { 
	        foreach ($object in $InputObject) 
	        { 
	            $DR = $DT.NewRow()   
	            foreach($property in $object.PsObject.get_properties()) 
	            {   
	                if ($first) 
	                {   
	                    $Col =  new-object Data.DataColumn   
	                    $Col.ColumnName = $property.Name.ToString()   
	                    if ($property.value) 
	                    { 
	                        if ($property.value -isnot [System.DBNull]) { 
	                            $Col.DataType = [System.Type]::GetType("$(Get-Type $property.TypeNameOfValue)") 
	                         } 
	                    } 
	                    $DT.Columns.Add($Col) 
	                }   
	                if ($property.IsArray) { 
	                    $DR.Item($property.Name) =$property.value | ConvertTo-XML -AS String -NoTypeInformation -Depth 1 
	                }   
	               else { 
	                    $DR.Item($property.Name) = $property.value 
	                } 
	            }   
	            $DT.Rows.Add($DR)   
	            $First = $false 
	        } 
	    }  
	      
	    End 
	    { 
	        Write-Output @(,($dt)) 
	    } 
	 
	} #Out-DataTable
	
	#Sample function that provides the location of the script
	function Get-ScriptDirectory
	{ 
		if($hostinvocation -ne $null)
		{
			Split-Path $hostinvocation.MyCommand.path
		}
		else
		{
			Split-Path $script:MyInvocation.MyCommand.Path
		}
	}
	
	#Sample variable that provides the location of the script
	[string]$ScriptDirectory = Get-ScriptDirectory
	#endregion general functions
	
	#region Exchange Validation
	# Validate the exchange environment is within the bounds of reality
	function Validate-Environment
	{
		# DAG Checks
		
		#### More than one server per DAG
		###$DAGServers = ($ExchangeEnv | Select -unique Name,Role,DAG_Name,Replication,Site | Where {($_.Role -eq "Mailbox") -and ($_.Name) -and ($_.DAG_Name)})
		###If ( $DAGServers.count -ge 2 ) 
		###{
		###	# Each DAG server has both a mapi and a replication network
		###	$MAPIInterfaces = @($DAGServers | Where {$_.Replication -eq "FALSE"})
		###	$REPLInterfaces = @($DAGServers | Where {$_.Replication -eq "TRUE"})
		###	If ($MAPIInterfaces.Count -eq $REPLInterfaces.Count) 
		###	{
		###		$IntCheckFailure = $false
		###		ForEach ($MAPIInt in $MAPIInterfaces)
		###		{
		###			$Match = $false
		###			ForEach ($REPLInt in $REPLInterfaces) 
		###			{
		###				If ($MAPIInt.ServerName -eq $REPLInt.ServerName)
		###				{
		###					$Match = $true
		###				}
		###			}
		###			If (!$Match)
		###			{
		###				$IntCheckFailure = $True
		###			}
		###		}
		###		If (!$IntCheckFailure)
		###		{
		###			# Each Site for each DAG has a DAG IP on the same interface as the MAPI interface
		###			ForEach ($
		###		}
		###		Else
		###		{
		###			$ValidDAGEnvironment = $false
		###			Write-Host -ForegroundColor Yellow "Not all DAG servers contain both a MAPI and Replication interface."
		###		}
		###	}
		###	Else
		###	{
		###		$ValidDAGEnvironment = $false
		###		Write-Host -ForegroundColor Yellow "Your MAPI and Replication networks do not match up."
		###	}
		###
		###}
		###Else
		###{
		###	$ValidDAGEnvironment = $false
		###	Write-Host -ForegroundColor Yellow "You need multiple servers for to define a DAG"
		###}
	
		# CAS Checks
		$ValidCASEnvironment = $true
			# All CAS servers have an array defined
				# There is more than one cas server in the same site with NLB defined
					# When NLB defined, an NLB shared IP is defined on the same interface
	}
	#endregion Exchange Validation
	
	#region Exchange Server Classification
	# Essentially get a list of all servers by role, version, sub-role, et cetera
	# Also get each hubcas server collection by their internet facing status. Joy!
	# I use the $script: construct to modify the global variables... and it makes
	# me feel a bit dirty on the inside....
	function Sort-ExchangeServers
	{
		# Exchange 2010 servers
		ForEach ($Server in $Servers)
		{
			switch ($Server."Exchange Version")
			{
			'Exchange 2012'
				{
					switch ($Server."Role")
					{
					'Mailbox' 
						{
							#<code>
						}
					'Hub/Transport/Database' 
						{
							#<code>
						}
					}
				}
			'Exchange 2010'
				{
	            $script:2010Servers += $Server
				switch ($Server."Role") {
				# Get our hub/cas object lists and script text output
				{ @('HUB/CAS', 'All-In-One') -contains $_ }
					{			
						$script:2010HubCASServers += $Server
						if ($Server."Internet Facing" -eq $true)
						{
							$script:2010HubCASServersExternal += $Server
							if ($Server."NLB" -eq $true)
							{
								$script:2010HubCASServersExternalNLB += $Server
								$script:2010NLBServers += $Server                                
							}
						}
						else
						{
							$script:2010HubCASServersInternal += $Server
							if ($Server."NLB" -eq $true)
							{
								$script:2010HubCASServersInternalNLB += $Server
								$script:2010NLBServers += $Server
							}
						}				
					}
				{@('Database','All-In-One') -contains $_}
					{
						$script:2010DatabaseServers += $Server
						If ($Server."DAG Replication" -eq $true)
						{
							$script:2010DatabaseDAGServers += $Server
						}
					}
				'Edge' 
					{
						$script:2010EdgeServers += $Server
						#<code>
					}
				}
			}
			'Exchange 2007'
				{
					$script:2007Servers += $Server
				}
			'Exchange 2003'
				{
					$script:2003Servers += $Server
				}
			}
			switch ($Server."Windows Version")
				{
				'Windows 2012'
					{
						$Script:2012Servers += $Server
					}
				'Windows 2008 R2'
					{
						$Script:2008R2Servers += $Server
					}
				}
		}
	    
	    # Determine our HUB/CAS internal/external sets (
	    if ($2010HubCASServersExternal.Count -ge 1)
	    {
	        # I'm sure there is a much more effective way of doing this
	        foreach ($External in $2010HubCASServersExternal)
	        {
	            if ($2010HubCASServersExternal.Count -ge 1)
	            {
	                foreach ($Internal in $2010HubCASServersInternal)
	                {
	            	    if ($External."Server Name" -eq $Internal."Server Name")
	                    {
	    		            $MatchedSet = New-Object PSObject -Property @{
	                		    'ServerName' = $External.'Server Name'
	                		    'ExchangeVersion' = $External.'Exchange Version'
	                		    'Site' = $External.'Site'
	                		    'ExternalAutodiscoverURL' = $External.'Autodiscover URL'
	                            'InternalAutodiscoverURL' = $Internal.'Autodiscover URL'
	                		    'ExternalCASURL' = $External.'CAS URL'
	                		    'InternalCASURL' = $Internal.'CAS URL'
							}
	                        $Script:2010HubCASSets += $MatchedSet
	    				}
	    			}
				}
		    }
		}
		# Possibly a 2003 to 2010 migration
		if (($2003Servers) -and ($2010Servers))
		{
			$Script:2003to2010 = $true
		}
		# Possibly a 2003 to 2012 migration
		if (($2003Servers) -and ($2012Servers))
		{
			$Script:2003to2012 = $true
		}
		# Possibly a 2007 to 2010 migration
		if (($2007Servers) -and ($2010Servers))
		{
			$Script:2007to2010 = $true
		}
		# Possibly a 2007 to 2012 migration
		if (($2007Servers) -and ($2012Servers))
		{
			$Script:2007to2012 = $true
		}
		# Possibly a 2010 to 2012 migration
		if (($2010Servers) -and ($2012Servers))
		{
			$Script:2010to2012 = $true
		}
	}
	#endregion Server Classification
	
	#region Exchange Prerequisites
	function Generate-Preqs
	{
		$ScriptOutput += @"

############################################## 
## PREREQUISITES DIRECTIONS/SCRIPTS         ##
##############################################

"@
		
		# Start with general items to be aware of
		$ScriptOutput += @"
<# General Notes (some snagged from http://msunified.net/2009/10/30/exchange-2010-prerequisites-on-server-2008-r2/ 
    others from multiple sources and experience)
	- Make sure that the functional level of your forest is at least Windows Server 2003
    - Also make sure that the Schema Master is running Windows Server 2003 with Service Pack 1 or later
"@
		if (($2010DatabaseDAGServers) -or ($2012DatabaseDAGServers))
		{
			$ScriptOutput += @"

    - If Database Availability Groups (DAG) is going to be used install Server 2008 R2 Enterprise Edition
        - Exchange 2010 Standard Edition supports DAG with up to 5 databases
        - Exchange 2010 Enterprise Edition supports up to 100 databases per server
        - You can install all server roles on the same server when using DAG
        - But then you need a hardware load balancer  for redundant CAS and HUB due to a Windows limitation preventing you from using Windows NLB and Clustering Services on the same Windows box
        - Two node DAG requires a witness that is not on a server within the DAG
            - Exchange 2010 automatically takes care of FSW creation; though you do have to specify the location of the FSW
            - It is recommended to specify the FSW to be created on a Hub Transport Server
            - Alternatively, you can put the witness on a non-Exchange Server after some prerequisites have been completed
            - You can follow these steps to get your member server to act as FSW
                add the “Exchange Trusted Subsystem” group to our Local Administrators Group on that member server
            - On servers that will host the Hub Transport or Mailbox server role, install the Microsoft Filter Pack. For details, see 2007 Office System Converter: Microsoft Filter Pack (this allows office attachments content to be searched and indexed)
    - Set Pagefile size, RAM + 10MB (for systems with 8 GB of RAM or less, set pagefile to RAM * 1,5)
    - Disable IPv6 by using this guide: http://support.microsoft.com/kb/929852
    - Run Windows Update untill all updates are installed
  
   DAG Network Configuration Best Practices ##
   Adaptor Component                                  MAPI NIC       Replication NIC
   ______________________________________________     ________       _______________
   Client for Microsoft Networks                      Enabled        Disabled
   QoS Packet Scheduler                               Optional       Optional
   File and Printer Sharing for Microsoft Networks    Enabled        Disabled
   Internet Protocol Version 6 (TCP/IP v6)            Optional       Optional
   Internet Protocol Version 4 (TCP/IP v4)            Enabled        Enabled
   Link-Layer Topology Discovery Mapper I/O Driver    Enabled        Enabled
   Link-Layer Topology Discovery Responder            Enabled        Enabled
   Register Connection in DNS                         Enabled        Disabled
   Default Gateway                                    Enabled        Disabled (use static routes)
#> 
"@
		}
	    if ($2003to2010)
		{
	     	$ScriptOutput += @"

<# 
**** Exchange 2003 to 2010 migration tips and requirements ****

#>		
"@
		}
		if ($2003to2012)
		{
	     	$ScriptOutput += @"

<# 
**** Exchange 2003 to 2012 migration tips and requirements ****

#>
"@
		}
		if ($2007to2010)
		{
	     	$ScriptOutput += @"
    
<#
**** Exchange 2007 to 2010 migration tips and requirements ****

#>
"@
		}
		if ($2007to2012)
		{
	     	$ScriptOutput += @"

<# 
**** Exchange 2007 to 2012 migration tips and requirements ****

#>    
"@
		}
		if ($2010to2012)
		{
	     	$ScriptOutput += @"
    
<# 
**** Exchange 2010 to 2012 migration tips and requirements ****

#>
"@
		}
	
	    # Process all Exchange 2010/Windows 2008R2 Combos
		If (($2008R2Servers) -and ($2010Servers))
		{
	        $ScriptOutput += @"

# **** EXCHANGE 2010 - Windows 2008 R2 ****
"@
			# Start with the HUB/CAS servers
			If ($2010HubCASServers)	
			{
				$ScriptOutput += @"

# **** EXCHANGE 2010 - HUB/CAS SERVERS ****

"@
				$Unique2010HUBCASServers = ($2010HubCASServers | Select -Unique -ExpandProperty "Server Name")
				foreach ($Server in $Unique2010HUBCASServers)
				{
					$ScriptOutput += @"

#--------------------------------------------- 
# $($Server.ToUpper()) - Prerequisites
#---------------------------------------------
# Run this via an administrative powershell console

Set-ExecutionPolicy RemoteSigned
Import-Module ServerManager
Enable-PSRemoting -Force
Add-WindowsFeature NET-Framework,RSAT-ADDS,Web-Server,Web-Basic-Auth,Web-Windows-Auth,Web-Metabase,Web-Net-Ext,Web-Lgcy-Mgmt-Console,WAS-Process-Model,RSAT-Web-Server,Web-ISAPI-Ext,Web-Digest-Auth,Web-Dyn-Compression,NET-HTTP-Activation,Web-Asp-Net,Web-Client-Auth,Web-Dir-Browsing,Web-Http-Errors,Web-Http-Logging,Web-Http-Redirect,Web-Http-Tracing,Web-ISAPI-Filter,Web-Request-Monitor,Web-Static-Content,Web-WMI,RPC-Over-HTTP-Proxy -Restart

# After the server has rebooted run this via an administrative powershell console
	Set-Service NetTcpPortSharing -StartupType Automatic

<# Per Technet documentation (http://technet.microsoft.com/en-us/library/bb691354.aspx) do the following:
	
 The following hotfix is required for Windows Server 2008 R2 and must be installed after the operating system 
 prerequisites have been installed:
   - Install the hotfix described in Knowledge Base article 982867, WCF services that are hosted by computers together 
    with a NLB fail in .NET Framework 3.5 SP1. For more information, see these MSDN Code Gallery pages:
        http://go.microsoft.com/fwlink/?linkid=3052&kbid=982867
        For additional background information, see KB982867 - WCF: Enable WebHeader settings on the RST/SCT.
            http://go.microsoft.com/fwlink/?LinkId=199857
        For the available downloads, see KB982867 - WCF: Enable WebHeader settings on the RST/SCT.
            http://go.microsoft.com/fwlink/?LinkId=199858

 Download and install the Office 2010 Filter Packs
  http://www.microsoft.com/en-us/download/details.aspx?id=17062

 After installing the preceding prerequisites, and before installing Exchange 2010, 
 it is recommended that you install any critical or recommended updates from Microsoft Update and reboot.
#>
"@
				}
			}
			# Now lets rock the Mailbox servers
			If ($2010DatabaseServers)
			{
				$ScriptOutput += @"
            
# **** EXCHANGE 2010 - MAILBOX SERVERS ****

"@
			$Unique2010DatabaseServers = ($2010DatabaseServers | Select -Unique -ExpandProperty "Server Name")
			foreach ($Server in $Unique2010DatabaseServers)
			{
				$ScriptOutput += @"

#--------------------------------------------- 
# $($Server.ToUpper()) - Prerequisites
#---------------------------------------------
#
# Run this via an administrative powershell console
Set-ExecutionPolicy RemoteSigned
Import-Module ServerManager
Enable-PSRemoting -Force
Add-WindowsFeature NET-Framework,RSAT-ADDS,Web-Server,Web-Basic-Auth,Web-Windows-Auth,Web-Metabase,Web-Net-Ext,Web-Lgcy-Mgmt-Console,WAS-Process-Model,RSAT-Web-Server -Restart

<# Per Technet documentation (http://technet.microsoft.com/en-us/library/bb691354.aspx) do the following:
	
The following hotfix is required for Windows Server 2008 R2 and must be installed after the operating system prerequisites have been installed:
   Install the hotfix described in Knowledge Base article 982867, WCF services that are hosted by computers together with a NLB 
   fail in .NET Framework 3.5 SP1. For more information, see these MSDN Code Gallery pages:
      http://go.microsoft.com/fwlink/?linkid=3052&kbid=982867
        For additional background information, see KB982867 - WCF: Enable WebHeader settings on the RST/SCT.
            http://go.microsoft.com/fwlink/?LinkId=199857
        For the available downloads, see KB982867 - WCF: Enable WebHeader settings on the RST/SCT.
            http://go.microsoft.com/fwlink/?LinkId=199858

After installing the preceding prerequisites, and before installing Exchange 2010, it is recommended 
that you install any critical or recommended updates from Microsoft Update and reboot.
#>
"@
				}
			}
			# Finally lets do the All-In-One servers
			If ($2010AllInOneServers)
			{
				$ScriptOutput += @"
            
# **** EXCHANGE 2010 - All-In-One SERVERS ****

"@
	
			$Unique2010AllInOneServers = ($2010AllInOneServers | Select -Unique -ExpandProperty "Server Name")
			foreach ($Server in $Unique2010AllInOneServers)
			{
				$ScriptOutput += @"

#--------------------------------------------- 
# $($Server.ToUpper()) - Prerequisites
#---------------------------------------------
#
# Run this via an administrative powershell console
Set-ExecutionPolicy RemoteSigned
Import-Module ServerManager
Enable-PSRemoting -Force
Add-WindowsFeature NET-Framework,RSAT-ADDS,Web-Server,Web-Basic-Auth,Web-Windows-Auth,Web-Metabase,Web-Net-Ext,Web-Lgcy-Mgmt-Console,WAS-Process-Model,RSAT-Web-Server,Web-ISAPI-Ext,Web-Digest-Auth,Web-Dyn-Compression,NET-HTTP-Activation,RPC-Over-HTTP-Proxy,Desktop-Experience -Restart

<# Per Technet documentation (http://technet.microsoft.com/en-us/library/bb691354.aspx) do the following:
	
The following hotfix is required for Windows Server 2008 R2 and must be installed after the operating system prerequisites have been installed:
   Install the hotfix described in Knowledge Base article 982867, WCF services that are hosted by computers together with a NLB 
   fail in .NET Framework 3.5 SP1. For more information, see these MSDN Code Gallery pages:
      http://go.microsoft.com/fwlink/?linkid=3052&kbid=982867
        For additional background information, see KB982867 - WCF: Enable WebHeader settings on the RST/SCT.
            http://go.microsoft.com/fwlink/?LinkId=199857
        For the available downloads, see KB982867 - WCF: Enable WebHeader settings on the RST/SCT.
            http://go.microsoft.com/fwlink/?LinkId=199858

After installing the preceding prerequisites, and before installing Exchange 2010, it is recommended 
that you install any critical or recommended updates from Microsoft Update and reboot.
#>
"@
				}
			}
		}
		Return $ScriptOutput
	}
	#endregion Exchange Prerequisites
	
	#region Exchange Coexistence
	function Generate-Coexistence
	{
		$ScriptOutput = @"
#region Migration/Coexistence Tips
# **** EXCHANGE Migration/Coexistence Tips ****
"@
		if ($2003to2010)
		{
			$ScriptOutput += @"

<# Exchange 2003 to 2010 Migration/Coexistence Tips
	Recommended Guides:
		http://technet.microsoft.com/en-us/library/ff805040.aspx
		http://milindn.files.wordpress.com/2010/01/rapid-transition-guide-from-exchange-2003-to-exchange-2010.pdf
		
	Pre-deployment:
		- Ensure that all blackberry enterprise servers are at a minimum of 5.0 SP2 MR5 or 5.0 SP3 MR3
		Reference(s):
			http://btsc.webapps.blackberry.com/btsc/viewdocument.do?noCount=true&externalId=KB22601&sliceId=2&cmd=displayKC&dialogID=932409796&docType=kc&stateId=1+0+206838079&docTypeID=DT_SUPPORTISSUE_1_1&ViewedDocsListHelper=com.kanisa.apps.common.BaseViewedDocsListHelperImpl
		To Download:
			https://swdownloads.blackberry.com/Downloads/entry.do?code=7B66B4FD401A271A1C7224027CE111BC
			
		- On all Exchange 2003 servers ensure that link state updates are suppressed
		Reference(s): 
			http://technet.microsoft.com/en-us/library/aa996728.aspx
		- All Exchange 2003 servers must be running Exchange 2003 SP2 or above
		- Run the following hotfix on all 2003 front-end (or all-in-one) servers to allow for activesync coexistence
			http://support.microsoft.com/?kbid=937031
			After this has been installed on existing servers follow the directions to enable windows integrated authentication
			for Microsoft-Server-ActiveSync on the 2003 servers.
			
	Migration:
		- Convert default object from LDAP to OPATH filters
		Reference(s):
			http://technet.microsoft.com/en-us/library/cc164375.aspx
			http://technet.microsoft.com/en-us/library/dd335105.aspx
		
		Upgrading the default objects
			Set-AddressList "All Users" -IncludedRecipients MailboxUsers
			Set-AddressList "All Groups" -IncludedRecipients MailGroups
			Set-AddressList "All Contacts" -IncludedRecipients MailContacts
			Set-AddressList "Public Folders" -RecipientFilter { RecipientType -eq 'PublicFolder' }
			Set-GlobalAddressList "Default Global Address List" -RecipientFilter {(Alias -ne $null -and (ObjectClass -eq 'user' -or ObjectClass -eq 'contact' -or ObjectClass -eq 'msExchSystemMailbox' -or ObjectClass -eq 'msExchDynamicDistributionList' -or ObjectClass -eq 'group' -or ObjectClass -eq 'publicFolder'))}
		
		- Convert custom object LDAP to OPATH filters
		Reference(s):
			http://technet.microsoft.com/en-us/library/cc164375.aspx
			http://gallery.technet.microsoft.com/scriptcenter/7c04b866-f83d-4b34-98ec-f944811dd48d
			
		Custom LDAP filters can be created for the following Exchange objects:
    		Address lists
    		E-mail address policies
    		Dynamic distribution groups
		
		To convert these you can use the LDAP to OPATH conversion script (see references).
		Example of upgrading dynamic distribution groups from LDAP to OPath (assuming .\ldap2opath.ps1 exists)
			Get-DynamicDistributionGroup -OrganizationalUnit "contoso.com" -Filter {RecipientFilter -eq $null} | foreach {Set-DynamicDistributionGroup $_.Name -RecipientFilter (.\ldap2opath.ps1 $_.ldapRecipientFilter ) -ForceUpgrade}
		
	Decomissioning:
#>
"@
		}
		if ($2003to2012)
		{
			$ScriptOutput += @"

<# Exchange 2003 to 2012 Migration/Coexistence Tips
There currently is no migration path from Exchange 2003 to Exchange 2012
#>
"@
		}
		if ($2007to2010)
		{
			$ScriptOutput += @"

<# Exchange 2007 to 2010 Migration/Coexistence Tips

#>
"@
		}
		if ($2007to2012)
		{
			$ScriptOutput += @"

<# Exchange 2007 to 2012 Migration/Coexistence Tips
There currently is no migration path from Exchange 2007 to Exchange 2012
#>
"@
		}
		if ($2010to2012)
		{
			$ScriptOutput += @"

<# Exchange 2010 to 2012 Migration/Coexistence Tips
The migration path from Exchange 2010 to Exchange 2012 is dependant on service pack 3 for Exchange 2010 which has not been released as of yet.
#>
"@
		}
		if (!($2003to2010 -or $2003to2012 -or $2007to2010 -or $2007to2012 -or $2010to2012))
		{
			$ScriptOutput += @"
		
# This is a fresh deployment!
"@
		}
		$ScriptOutput += @"
		
#endregion Migration/Coexistence Tips
"@
		Return $ScriptOutput
	}
	#endregion Exchange Coexistence
	
	#region Exchange CAS Configuration
	function Generate-CASConfiguration 
	{
		$ScriptOutput = ''
		# Externally facing HUB/CAS Servers
		If ($2010HubCASServersExternal)	
		{
			$ScriptOutput += @"

#region Hub/CAS - External
		
# **** EXCHANGE 2010 - EXTERNAL HUB/CAS SERVERS ****

"@
			$Unique2010HUBCASServersExternal = $2010HubCASServersExternal | sort-object -Property "Server Name" -Unique
			ForEach ($Server in $Unique2010HUBCASServersExternal)
			{            
				$ScriptOutput += @"

#------------------------------------------------ 
# $(($Server."Server Name").ToUpper()) - External Hub/CAS Config
#-------------------------------------------------
"@
	    
		        if ($Redirection)
	            {
	                $ScriptOutput += @"

`$remotecmd = {   
    # Functions
    function Get-Exchange2010InstallPath	
    {
        # Get the root setup entries.
        `$setupRegistryPath = "HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup"
        `$setupEntries = Get-ItemProperty `$setupRegistryPath
        if(`$setupEntries -eq `$null) {
            return `$null
        }

        # Try to get the Install Path.
        `$InstallPath = `$setupEntries.MsiInstallPath
        return `$InstallPath
    }

    # Module/Snapin/Dot Sourcing
    if (! (Get-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction:SilentlyContinue) )
    { 
    	Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 
    }
    Import-Module WebAdministration
    
    Backup-WebConfiguration -Name $($Server."Server Name")-IISBackup

    # Setup http redirection on external cas
    Write-Host "``n"
	Write-Host "$($Server."Server Name") - Setting SSL configuration" -ForegroundColor green	
    Set-WebConfigurationProperty -Filter //security/access -name sslflags -Value "" -PSPath IIS:\ -Location "Default Web Site/"
	Set-WebConfigurationProperty -Filter //httpRedirect -Name enabled -Value "true" -PSPath IIS:\ -Location "Default Web Site/"
	Set-WebConfigurationProperty -Filter //httpRedirect -Name destination -Value "https://$($Server."CAS URL")/owa" -PSPath IIS:\ -Location "Default Web Site/"
	Set-WebConfigurationProperty -Filter //httpRedirect -Name ChildOnly -Value "true" -PSPath IIS:\ -Location "Default Web Site/"
	Set-WebConfigurationProperty -Filter //httpRedirect -Name httpResponseStatus -Value "Found" -PSPath IIS:\ -Location "Default Web Site/"
	
	# Fix redirection of non owa external directories
	Set-WebConfiguration system.webserver/httpRedirect "IIS:\sites\Default Web Site\Autodiscover" -Value @{enabled="false";destination="";ChildOnly="true";httpResponseStatus="Found"}
	Set-WebConfiguration system.webserver/httpRedirect "IIS:\sites\Default Web Site\ews" -Value @{enabled="false";destination="";ChildOnly="true";httpResponseStatus="Found"}
	Set-WebConfiguration system.webserver/httpRedirect "IIS:\sites\Default Web Site\Microsoft-Server-ActiveSync" -Value @{enabled="false";destination="";ChildOnly="true";httpResponseStatus="Found"}
	Set-WebConfiguration system.webserver/httpRedirect "IIS:\sites\Default Web Site\OAB" -Value @{enabled="false";destination="";ChildOnly="true";httpResponseStatus="Found"}

    # Fix OAB virtual directory permission issues
    `$PathToFile = (Get-Exchange2010InstallPath)+"ClientAccess\OAB\web.config"
    `$runcmd = icacls `$PathToFile /grant:R 'NT Authority\Authenticated Users:RX'
    `$runcmd

    # Be nice and restart IIS
    `$runcmd = iisreset
    `$runcmd
"@
				}
	            $ScriptOutput += @"
            
}

Invoke-Command -ComputerName $($Server."IP") -ScriptBlock `$remotecmd
"@
	            if ($ReverseProxy)
	            {
	                $ScriptOutput += @"
                
# External Authentication
Set-OwaVirtualDirectory "$($Server."Server Name")(default web site)\owa" -ExternalUrl "https://$($Server."CAS URL")/owa" -ExternalAuthenticationMethods "Basic" -BasicAuthentication $true -FormsAuthentication $false
Set-ECPVirtualDirectory "$($Server."Server Name")(default web site)\ecp" -ExternalUrl "https://$($Server."CAS URL")/owa" -ExternalAuthenticationMethods "Basic" -BasicAuthentication $true -FormsAuthentication $false
"@
				}
			}
	        $ScriptOutput += @"

#endregion Hub/CAS - External
"@
		}
	
		# Process Internal Hub/CAS Information
		If ($2010HubCASServersInternal)	
		{
			$ScriptOutput += @"

#region Hub/CAS - Internal

# **** EXCHANGE 2010 - INTERNAL HUB/CAS SERVERS ****

"@
			$Unique2010HUBCASServersInternal = $2010HubCASServersInternal | sort-object -Property "Server Name" -Unique
	
			# First process the individual Internal CAS servers
			ForEach ($Server in $Unique2010HubCASServersInternal)
			{
				$ScriptOutput += @"

#------------------------------------------------ 
# $(($Server."Server Name").ToUpper()) - Internal Hub/CAS Config
#-------------------------------------------------
                
`$remotecmd = {   
    # Functions
    function Get-Exchange2010InstallPath	
    {
        # Get the root setup entries.
        `$setupRegistryPath = "HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup"
        `$setupEntries = Get-ItemProperty `$setupRegistryPath
        if(`$setupEntries -eq `$null) {
            return `$null
        }

        # Try to get the Install Path.
        `$InstallPath = `$setupEntries.MsiInstallPath
        return `$InstallPath
    }

    # Module/Snapin/Dot Sourcing
    if (! (Get-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction:SilentlyContinue) )
    { 
    	Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 
    }
    Import-Module WebAdministration

    # Create folder on $($Server."Server Name") for internal facing website
    New-Item -Path $($vDirLocation) -type directory -Force 
    Write-Host "Folder creation complete"
    
    Backup-WebConfiguration -Name $($Server."Server Name")-IISBackup
	# Create the new internal site and virtual directories
	New-Website -Name "Exchange-Internal" -IPAddress $($Server."IP") -PhysicalPath $($vDirLocation) -Port 443 -Ssl
	Set-WebBinding -Name "Exchange-Internal" -IPAddress $($Server."IP") -Port 80
	Remove-WebBinding -Name "Exchange-Internal" -IPAddress "127.0.0.1" -Port 443
	Remove-WebBinding -Name "Exchange-Internal" -IPAddress "127.0.0.1" -Port 80
	New-ActiveSyncVirtualDirectory -WebSiteName "Exchange-Internal" -InternalURL "https://$($Server."CAS URL")/Microsoft-Server-ActiveSync"
	New-WebServicesVirtualDirectory -WebSiteName "Exchange-Internal" -InternalURL "https://$($Server."CAS URL")/ews/exchange.asmx" -Force
	New-AutodiscoverVirtualDirectory -WebSiteName "Exchange-Internal" -InternalURL "https://$($Server."Autodiscover URL")/autodiscover"
	New-OWAVirtualDirectory -WebSiteName "Exchange-Internal" -InternalURL "https://$($Server."CAS URL")/owa"
	New-OABVirtualDirectory -WebSiteName "Exchange-Internal" -InternalURL "https://$($Server."CAS URL")/OAB"
	New-ECPVirtualDirectory -WebSiteName "Exchange-Internal" -InternalURL "https://$($Server."CAS URL")/ECP"

"@
	            if ($Redirection)
	            {
	                $ScriptOutput += @"

	# Make ssl a requirement on the root directory
	Set-WebConfigurationProperty -Filter //security/access -name sslflags -Value "Ssl" -PSPath IIS:\ -Location "Exchange-Internal/"
	
	#Setup http redirection
	Set-WebConfigurationProperty -Filter //httpRedirect -Name enabled -Value "true" -PSPath IIS:\ -Location "Exchange-Internal/"
	Set-WebConfigurationProperty -Filter //httpRedirect -Name destination -Value "https://$($Server."CAS URL")/owa" -PSPath IIS:\ -Location "Exchange-Internal/"
	Set-WebConfigurationProperty -Filter //httpRedirect -Name ChildOnly -Value "true" -PSPath IIS:\ -Location "Exchange-Internal/"
	Set-WebConfigurationProperty -Filter //httpRedirect -Name httpResponseStatus -Value "Found" -PSPath IIS:\ -Location "Exchange-Internal/"

	# Fix redirection of non owa-internal directories
	Set-WebConfiguration system.webserver/httpRedirect "IIS:\sites\Exchange-Internal\Autodiscover" -Value @{enabled="false";destination="";ChildOnly="true";httpResponseStatus="Found"}
	Set-WebConfiguration system.webserver/httpRedirect "IIS:\sites\Exchange-Internal\ews" -Value @{enabled="false";destination="";ChildOnly="true";httpResponseStatus="Found"}
	Set-WebConfiguration system.webserver/httpRedirect "IIS:\sites\Exchange-Internal\Microsoft-Server-ActiveSync" -Value @{enabled="false";destination="";ChildOnly="true";httpResponseStatus="Found"}
	Set-WebConfiguration system.webserver/httpRedirect "IIS:\sites\Exchange-Internal\OAB" -Value @{enabled="false";destination="";ChildOnly="true";httpResponseStatus="Found"}
    
    # Fix OAB virtual directory permission issues
    `$PathToFile = (Get-Exchange2010InstallPath)+"ClientAccess\OAB\web.config"
    `$runcmd = icacls `$PathToFile /grant:R 'NT Authority\Authenticated Users:RX'
    `$runcmd

    # Be nice and restart IIS
    `$runcmd = iisreset
    `$runcmd
"@
				}
	            $ScriptOutput += @"
            
}

Invoke-Command -ComputerName $($Server."IP") -ScriptBlock `$remotecmd
"@
			}
	        $ScriptOutput += @"
        
#endregion Hub/CAS - Internal
"@
		}
		If ($2010NLBServers.Count -ge 1)
		{
			$ScriptOutput += @"

#region NLB Configuration
<#
**** EXCHANGE 2010 - NLB CONFIG ****
#>

"@
			$Unique2010NLBServers = $2010NLBServers | sort-object -Property "Server Name" -Unique
	
			# First process the individual Internal CAS servers
			ForEach ($Server in $Unique2010NLBServers)
			{
				$ScriptOutput += @"

#------------------------------------------------ 
# $(($Server."Server Name").ToUpper()) - Add NLB Feature
#-------------------------------------------------
Invoke-Command -ComputerName $($Server."IP") -ScriptBlock {
	Import-Module ServerManager
	Add-WindowsFeature NLB, RSAT-NLB	
}
"@
			}
			#region ExternalNLB
			if ($2010HubCASServersExternalNLB)
			{
				$Unique2010HubCASServersExternalNLB = $2010HubCASServersExternalNLB | sort-object -Property "Server Name" -Unique
				$count = 0
				# First process the individual External CAS servers		
				ForEach ($Server in $Unique2010HubCASServersExternalNLB)
				{
					$count++
					if ($count -eq 1)
					{
						$firstclusternode = $Server
						$ScriptOutput += @"

#-------------------------------------------------
# $(($Server."Server Name").ToUpper()) - Configure External NLB Cluster
#-------------------------------------------------
Invoke-Command -ComputerName $($Server."IP") -ScriptBlock {
	`$runcmd = Netsh interface ipv4 set int $($Server."Interface Name") forwarding=enabled
	`$runcmd
	if(`$runcmd -match 'ok'){write-Host "forwarding is enabled"}
    else{write-Host "Failed to configure forwarding on interface"}
	Import-Module NetworkLoadBalancingClusters
	Write-Host "Creating NLB Cluster" -ForegroundColor yellow
	New-NlbCluster -InterfaceName $($Server."Interface Name") -ClusterName CAS-External -HostName $($Server."Server Name") -ClusterPrimaryIP $($Server."NLB or DAG IP")
	Get-NlbClusterPortRule -HostName . | Remove-NlbClusterPortRule -Force
	Write-Host "Adding port rules for Exchange CAS" -ForegroundColor yellow
	Add-NlbClusterPortRule -StartPort 25 -EndPort 25 -Protocol TCP -Affinity Single -Mode Multiple -InterfaceName $($Server."Interface Name") | Out-Null
	Write-Host "Added port rule for SMTP (tcp 25)" -ForegroundColor yellow
	Add-NlbClusterPortRule -StartPort 80 -EndPort 80 -Protocol TCP -Affinity Single -Mode Multiple -InterfaceName $($Server."Interface Name") | Out-Null
	Write-Host "Added port rule for http (tcp 80)" -ForegroundColor yellow
	Add-NlbClusterPortRule -StartPort 110 -EndPort 110 -Protocol TCP -Affinity Single -Mode Multiple -InterfaceName $($Server."Interface Name") | Out-Null
	Write-Host "Added port rule for POP3 (tcp 110)" -ForegroundColor yellow
	Add-NlbClusterPortRule -StartPort 135 -EndPort 135 -Protocol TCP -Affinity Single -Mode Multiple -InterfaceName $($Server."Interface Name") | Out-Null
	Write-Host "Added port rule for RPC (tcp 135)" -ForegroundColor yellow
	Add-NlbClusterPortRule -StartPort 143 -EndPort 143 -Protocol TCP -Affinity Single -Mode Multiple -InterfaceName $($Server."Interface Name") | Out-Null
	Write-Host "Added port rule for IMAP4 (tcp 143)" -ForegroundColor yellow
	Add-NlbClusterPortRule -StartPort 443 -EndPort 443 -Protocol TCP -Affinity Single -Mode Multiple -InterfaceName $($Server."Interface Name") | Out-Null
	Write-Host "Added port rule for https (tcp 443)" -ForegroundColor yellow
	Add-NlbClusterPortRule -StartPort 59531 -EndPort 59531 -Protocol Both -Affinity Single -Mode Multiple -InterfaceName $($Server."Interface Name") | Out-Null
	Write-Host "Added port rule for MSExchange RPC (tcp 59531)" -ForegroundColor yellow
	Add-NlbClusterPortRule -StartPort 59532 -EndPort 59532 -Protocol Both -Affinity Single -Mode Multiple -InterfaceName $($Server."Interface Name") | Out-Null
	Write-Host "Added port rule for MSExchange Address Book service (tcp 59532)" -ForegroundColor yellow
}
"@
					}
					else
					{
						$ScriptOutput += @"

#------------------------------------------------ 
# $(($Server."Server Name").ToUpper()) - Configure External NLB Members
#-------------------------------------------------
Invoke-Command -ComputerName $($Server."IP") -ScriptBlock {
	`$runcmd = netsh interface ipv4 set int $($Server."Interface Name") forwarding=enabled
	`$runcmd
	if(`$runcmd -match 'ok'){write-Host "forwarding is enabled"}
    else{write-Host "Failed to configure forwarding on interface"}
	Import-Module NetworkLoadBalancingClusters
	Get-NlbCluster -Hostname $($firstclusternode."Server Name")  -InterfaceName $($firstclusternode."Interface Name") | Add-NlbClusterNode -NewNodeName $($Server."Server Name") -NewNodeInterface $($Server."Interface Name")
}
"@
					}
				}
			}
			#endregion ExternalNLB
			
			#region InternalNLB
			if ($2010HubCASServersInternalNLB)
			{
				$Unique2010HubCASServersInternalNLB = $2010HubCASServersInternalNLB | sort-object -Property "Server Name" -Unique
				$count = 0
				# First process the individual Internal CAS servers
				ForEach ($Server in $Unique2010HubCASServersInternalNLB)
				{
					$count++
					if ($count -eq 1)
					{
						$firstclusternode = $Server
						$ScriptOutput += @"

#------------------------------------------------ 
# $(($Server."Server Name").ToUpper()) - Configure Internal NLB Cluster
#-------------------------------------------------
Invoke-Command -ComputerName $($Server."IP") -ScriptBlock {
	`$runcmd = Netsh interface ipv4 set int $($Server."Interface Name") forwarding=enabled
	`$runcmd
	if(`$runcmd -match 'ok'){write-Host "forwarding is enabled"}
    else{write-Host "Failed to configure forwarding on interface"}
	Import-Module NetworkLoadBalancingClusters
	New-NlbCluster -InterfaceName $($Server."Interface Name") -ClusterName CAS-Internal -HostName $($Server."Server Name") -ClusterPrimaryIP $($Server."NLB or DAG IP")
	Get-NlbClusterPortRule -HostName . | Remove-NlbClusterPortRule -Force
	Write-Host "Adding port rules for Exchange CAS" -ForegroundColor yellow
	Add-NlbClusterPortRule -StartPort 25 -EndPort 25 -Protocol TCP -Affinity Single -Mode Multiple -InterfaceName $($Server."Interface Name") | Out-Null
	Write-Host "Added port rule for SMTP (tcp 25)" -ForegroundColor yellow
	Add-NlbClusterPortRule -StartPort 80 -EndPort 80 -Protocol TCP -Affinity Single -Mode Multiple -InterfaceName $($Server."Interface Name") | Out-Null
	Write-Host "Added port rule for http (tcp 80)" -ForegroundColor yellow
	Add-NlbClusterPortRule -StartPort 110 -EndPort 110 -Protocol TCP -Affinity Single -Mode Multiple -InterfaceName $($Server."Interface Name") | Out-Null
	Write-Host "Added port rule for POP3 (tcp 110)" -ForegroundColor yellow
	Add-NlbClusterPortRule -StartPort 135 -EndPort 135 -Protocol TCP -Affinity Single -Mode Multiple -InterfaceName $($Server."Interface Name") | Out-Null
	Write-Host "Added port rule for RPC (tcp 135)" -ForegroundColor yellow
	Add-NlbClusterPortRule -StartPort 143 -EndPort 143 -Protocol TCP -Affinity Single -Mode Multiple -InterfaceName $($Server."Interface Name") | Out-Null
	Write-Host "Added port rule for IMAP4 (tcp 143)" -ForegroundColor yellow
	Add-NlbClusterPortRule -StartPort 443 -EndPort 443 -Protocol TCP -Affinity Single -Mode Multiple -InterfaceName $($Server."Interface Name") | Out-Null
	Write-Host "Added port rule for https (tcp 443)" -ForegroundColor yellow
	Add-NlbClusterPortRule -StartPort 59531 -EndPort 59531 -Protocol Both -Affinity Single -Mode Multiple -InterfaceName $($Server."Interface Name") | Out-Null
	Write-Host "Added port rule for MSExchange RPC (tcp 59531)" -ForegroundColor yellow
	Add-NlbClusterPortRule -StartPort 59532 -EndPort 59532 -Protocol Both -Affinity Single -Mode Multiple -InterfaceName $($Server."Interface Name") | Out-Null
	Write-Host "Added port rule for MSExchange Address Book service (tcp 59532)" -ForegroundColor yellow
}
"@
					}
					else
					{
						$ScriptOutput += @"

#------------------------------------------------ 
# $(($Server."Server Name").ToUpper()) - Configure Internal NLB Members
#-------------------------------------------------
Invoke-Command -ComputerName $($Server."IP") -ScriptBlock {
	`$runcmd = netsh interface ipv4 set int $($Server."Interface Name") forwarding=enabled
	`$runcmd
	if(`$runcmd -match 'ok'){write-Host "forwarding is enabled"}
    else{write-Host "Failed to configure forwarding on interface"}
	Import-Module NetworkLoadBalancingClusters
	Get-NlbCluster -Hostname $($firstclusternode."Server Name")  -InterfaceName $($firstclusternode."Interface Name") | Add-NlbClusterNode -NewNodeName $($Server."Server Name") -NewNodeInterface $($Server."Interface Name")
}
"@
					}
				}
			}
			#endregion InternalNLB
		$ScriptOutput += @"

#endregion NLB Configuration
"@
		}
		If ($GenerateCASArrays)
		{
	        	$ScriptOutput += @"

#region CAS Array Generation
# **** EXCHANGE 2010 - CAS Array Generation ****
## NOTE: If you are setting this up then you really
##       need to look and make sure these CAS array
##       names are what works in your envrironment and
##       are on your SAN certificates.
##       (it is VERY unlikely that they are!).
##
##       It is best practice for these array names
##       to differ from the external access name.
##       this ensures that there are not long delays
##       when accessing mail externally with outlook
##       anywhere.
"@
			$2010UniqueSites = $2010Servers | sort-object -Property "Site" -Unique
			
			Foreach ($SiteName in $2010UniqueSites)
			{
				If ($UseCASInternalDomainNameSpace)
	            {
					$ScriptOutput += @"


#------------------------------------------------ 
# $(($SiteName."Site").ToUpper()) - CAS Array
#-------------------------------------------------
New-ClientAccessArray -fqdn $($SiteName."Site")-Array.$($SiteName."Internal Domain") -site $($SiteName."Site")
"@
				}
				Else
				{
					$ScriptOutput += @"
			
#------------------------------------------------ 
# $(($SiteName."Site").ToUpper()) - CAS Array
#-------------------------------------------------
	New-ClientAccessArray -fqdn $($SiteName."Site")-Array.$($ExternalNamespace) -site $($SiteName."Site")
"@
				}
			}
		}
	    $ScriptOutput += @"

#endregion CAS Array Generation

#region CAS Static Ports
`$storageDir = `$pwd
`$webclient = New-Object System.Net.WebClient
`$url = "http://cid-14adc5cf1e0cbccf.skydrive.live.com/download.aspx/.Public/Exchange%202010/Scripts/Set-StaticPorts.ps1"
`$setstaticfile = "`$storageDir\Set-StaticPorts.ps1"
`$webclient.DownloadFile(`$url,`$setstaticfile)
"@
	
		$2010UniqueHubCASServers = $2010HubCASServers | sort-object -Property "Server Name" -Unique 
		ForEach ($Server in $2010UniqueHubCASServers)
		{
			$ScriptOutput += @"
		
`&`$setstaticfile -Server $($Server."Server Name") 
"@
		}
		    $ScriptOutput += @"

#endregion CAS Static Ports
"@
		Return $ScriptOutput
	}
	#endregion Exchange CAS Configuration
#endregion Source: Globals.ps1

#region Source: ChildForm.pff
function Call-ChildForm_pff
{
#region File Recovery Data (DO NOT MODIFY)
<#RecoveryData:
gwgAAB+LCAAAAAAABAC9VluPmkAUfm/S/0B4tiAi6iZIYrDbbHozi2n71oxw0OkOM2QY3NJf3+Eq
CO5uTLsxMZ7vXOebOedo34PPjsCzNRLIeftGUeyvHO8xReQWE/iCInDcAybBLeORFoehrff0hVcu
fQOeYEYdQ5vYehso4+5+gS8UkcWwVL0sERBp3zEN2GOi5dHL75EypBopVajlVBvnn5HipkSkHJYU
UsERGSmbdEew/xGyLXsAutzN58jyrZlxY05hvLhRFSqLXaqhjGeoip8fiks71WVUcEYStShTFrrh
LAYussrBJRio8PAfUJ3JYjpSJjN5vNroglNOjOoUuZ613cJvoTobjiNElPy4fY/3R1lCZf6JoUB1
crsC/ZnLtl78LonWS6ZLYZUkEEliIKljVUjmVER/RhTtIZLu2ioVLEJC8nwi3Hwp4aaxC82FNUOB
OZuCadl6k2k48ynH5NpLfTbHpWf0HzOuOXrEdH9NrrEZWuE8NIzAGiMTPZErSnzGCd79g8Z4wYnk
bHgV6n5E5HWuCHPZIIxnHvAj9uGqh/H0ZTVi3Xn2BvkPstF4XdQHoCADV+LJoJmpBagPohV451fD
tWPaRWV74xAS4XIoOtuRA6mHNdZumsgRUOtbsS8pNgSJfMw57/JBVwuN+j6l3nbljG29+nVyZI/A
vQMQUm8JUwbogRVZepct2wM/5ViOUL1Gug5nG+l0jkFUjtsg9UXf+pKiws+vaghdQ+JzHHdp1gdR
l0Uxolmb9XPEZXEm9++hczV97I4K4HJHnxU4DA8u9NLhskpWFtUbp6m1hdh6Z/nrrdcv30L7T8df
pqdH/IMIAAA=#>
#endregion
	#----------------------------------------------
	#region Import the Assemblies
	#----------------------------------------------
	[void][reflection.assembly]::Load("System, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
	[void][reflection.assembly]::Load("mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Xml, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.DirectoryServices, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
	#endregion Import Assemblies

	#----------------------------------------------
	#region Generated Form Objects
	#----------------------------------------------
	[System.Windows.Forms.Application]::EnableVisualStyles()
	$form1 = New-Object 'System.Windows.Forms.Form'
	$InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'
	#endregion Generated Form Objects

	#----------------------------------------------
	# User Generated Script
	#----------------------------------------------
	
	
	$FormEvent_Load={
		#TODO: Initialize Form Controls here
		
	}
	
		# --End User Generated Script--
	#----------------------------------------------
	#region Generated Events
	#----------------------------------------------
	
	$Form_StateCorrection_Load=
	{
		#Correct the initial state of the form to prevent the .Net maximized form issue
		$form1.WindowState = $InitialFormWindowState
	}
	
	$Form_StoreValues_Closing=
	{
		#Store the control values
	}

	
	$Form_Cleanup_FormClosed=
	{
		#Remove all event handlers from the controls
		try
		{
			$form1.remove_Load($FormEvent_Load)
			$form1.remove_Load($Form_StateCorrection_Load)
			$form1.remove_Closing($Form_StoreValues_Closing)
			$form1.remove_FormClosed($Form_Cleanup_FormClosed)
		}
		catch [Exception]
		{ }
	}
	#endregion Generated Events

	#----------------------------------------------
	#region Generated Form Code
	#----------------------------------------------
	#
	# form1
	#
	$form1.ClientSize = '284, 262'
	$form1.Name = "form1"
	$form1.Text = "Primal Form"
	$form1.add_Load($FormEvent_Load)
	#endregion Generated Form Code

	#----------------------------------------------

	#Save the initial state of the form
	$InitialFormWindowState = $form1.WindowState
	#Init the OnLoad event to correct the initial state of the form
	$form1.add_Load($Form_StateCorrection_Load)
	#Clean up the control events
	$form1.add_FormClosed($Form_Cleanup_FormClosed)
	#Store the control values when form is closing
	$form1.add_Closing($Form_StoreValues_Closing)
	#Show the Form
	return $form1.ShowDialog()

}
#endregion Source: ChildForm.pff

#Start the application
Main ($CommandLine)