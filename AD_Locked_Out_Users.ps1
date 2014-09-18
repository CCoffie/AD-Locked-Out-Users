#Caleb Coffie
#https://www.CalebCoffie.com


#TODO No accounts locked errors


#Imports
Import-Module ActiveDirectory
Add-Type -AssemblyName "System.DirectoryServices"
Add-Type -AssemblyName "System.Drawing"
Add-Type -AssemblyName "System.Windows.Forms"

$GUI = New-Object System.Windows.Forms.Form
$lockedOutUsersList = New-Object Windows.Forms.DataGridView
$RefreshButton = New-Object Windows.Forms.Button
$OnTopCheckbox = New-Object Windows.Forms.Checkbox
$AutoUpdateTimer = New-Object Windows.Forms.Timer
$AutoUpdateCheckbox = New-Object Windows.Forms.Checkbox
$RefreshingLabel = New-Object Windows.Forms.Label
$NotificationIcon = New-Object Windows.Forms.NotifyIcon

$LockedOutUsersGridData = New-Object System.Collections.ArrayList
$LastLockedOutUsersGridData = New-Object System.Collections.ArrayList

#Get Directory in which script is executing
$fullPathIncFileName = $MyInvocation.MyCommand.Definition
$currentScriptName = $MyInvocation.MyCommand.Name
$currentExecutingPath = $fullPathIncFileName.Replace($currentScriptName, "")

#Set Resources Directory
$resourcesDirectory = $currentExecutingPath + "Resources\"

#NotificationIcon Attributes setup
#Set NotifyIcon Icon
$Icon = New-Object System.Drawing.Icon($($resourcesDirectory + "icon.ico"))
$NotificationIcon.Icon = $Icon

#Main GUI Settings
$GUI.Text = "Active Directory Locked-out Users"
$GUI.NAME = "Active Directory Locked-out Users"
$GUI.Icon = $Icon
#$GUI.AutoSize = $True
$GUI.ClientSize = '700,200'
$GUI.StartPosition = "CenterScreen"

#Timer Settings
$AutoUpdateTimer.Interval = 30000

#Refresh Button
$GUI.Controls.Add($RefreshButton)
$RefreshButton.text = "Refresh Now"
$RefreshButton.AutoSize = $True
$RefreshButton.Anchor = 'Bottom, Right'
$RefreshButton.Location = '610,165'

#On Top Checkbox
$GUI.Controls.Add($OnTopCheckbox)
$OnTopCheckbox.text = "Always On Top"
$OnTopCheckbox.Location = '10,170'
$OnTopCheckbox.AutoSize = $True
$OnTopCheckbox.Anchor = 'Bottom, Left'

#Auto Update Checkbox
$GUI.Controls.Add($AutoUpdateCheckbox)
$AutoUpdateCheckbox.text = "Auto Update"
$AutoUpdateCheckbox.Location = '500,170'
$AutoUpdateCheckbox.AutoSize = $True
$AutoUpdateCheckbox.Anchor = 'Bottom, Right'

#Refreshing Label
$GUI.Controls.Add($RefreshingLabel)
$RefreshingLabel.Text = "Refreshing..."
$RefreshingLabel.AutoSize = $True
$RefreshingLabel.Anchor = 'Left, Right, Bottom'
$RefreshingLabel.Location = '300,170'
$RefreshingLabel.Visible = $False

#Main Table View
$GUI.Controls.Add($lockedOutUsersList)
$lockedOutUsersList.Name = "Locked-out Users"
$lockedOutUsersList.Location = '10,10'
$lockedOutUsersList.Size = '680,150'
$lockedOutUsersList.Anchor = 'Top, Bottom, Left, Right'
$lockedOutUsersList.AllowUserToAddRows = $False
$lockedOutUsersList.AllowUsertoDeleteRows = $False
$lockedOutUsersList.ReadOnly = $True
$lockedOutUsersList.ColumnHeadersHeightSizeMode = 'DisableResizing'
$lockedOutUsersList.RowHeadersVisible = $False
$lockedOutUsersList.SelectionMode = 'FullRowSelect'
$lockedOutUsersList.MultiSelect = $False
$lockedOutUsersList.AllowUserToResizeRows = $False
$lockedOutUsersList.AllowUserToResizeColumns = $False

#On Shown Refresh Data Event handler
function GUI_RefreshData {
	#Make Refreshing Label Visible
	$RefreshingLabel.Visible = $True
	$GUI.Refresh()
	
	#Run command to find locked out users
	$LockedOutUsersData = Search-ADAccount -LockedOut
	
	#Keep a history for notifications
    $LastLockedOutUsersGridData.Clear()
	$LastLockedOutUsersGridData = $($LockedOutUsersGridData) #Fill with last run data $() is needed to get values and not reference
	
	If(($($LastLockedOutUsersGridData -is [array]) -eq $False) -and $LastLockedOutUsersGridData -ne $null)
	{
		$tempdata = $($LastLockedOutUsersGridData)
		$LastLockedOutUsersGridData = New-Object System.Collections.ArrayList
		$LastLockedOutUsersGridData.Add($tempdata)
	}
	
	if($LockedOutUsersData -is [array])
	{		
		$LockedOutUsersGridData.Clear() #Clear Old data
		$LockedOutUsersGridData.AddRange($LockedOutUsersData)	
	}
	else
	{
		#Exception for when only one user is locked out
		$tempArray = @($($LockedOutUsersData),$($LockedOutUsersData))
		$LockedOutUsersGridData.Clear() #Clear Old data
		$LockedOutUsersGridData.AddRange($tempArray)
		$LockedOutUsersGridData.Remove($LockedOutUsersData)
	}
	
	$lockedOutUsersList.DataSource = $NULL
	$lockedOutUsersList.DataSource = $LockedOutUsersGridData
	$lockedOutUsersList.AutoResizeColumns("AllCells")
	
	#Format Data Table
	FormatDataForTable
	
	#Check for changes since last refresh
	If(($LastLockedOutUsersGridData -ne $null) -and ($LockedOutUsersGridData -ne $null) -and ($(Compare-Object $LockedOutUsersGridData $LastLockedOutUsersGridData) -ne $null))
	{
		$ChangesSinceLastRefresh = $(Compare-Object $LockedOutUsersGridData $LastLockedOutUsersGridData)
		
		foreach ($Change in $ChangesSinceLastRefresh) {
			If ($Change.SideIndicator -eq "=>")
			{
				#Account is no longer locked
				If ($Change.InputObject.Name -eq $null)
				{continue}
				$NotificationIcon.BalloonTipText = $Change.InputObject.Name
				$NotificationIcon.BalloonTipTitle = "No longer locked." 
				$NotificationIcon.BalloonTipIcon = "Info" 
				$NotificationIcon.Visible = $True 
				$NotificationIcon.ShowBalloonTip(10000)
			}
			Else
			{
				#Account is now locked
				if ($Change.InputObject.Name -eq $null)
				{continue}
				$NotificationIcon.BalloonTipText = $Change.InputObject.Name
				$NotificationIcon.BalloonTipTitle = "Locked"
				$NotificationIcon.BalloonTipIcon = "Warning" 
				$NotificationIcon.Visible = $True 
				$NotificationIcon.ShowBalloonTip(10000)
			}
		}
	}
	
	#Make Refreshing Label Invisible
	$RefreshingLabel.Visible = $False
	$GUI.Refresh()
}

function FormatDataForTable {
	
	#Hide Un-needed Columns
	$lockedOutUsersList.Columns["DistinguishedName"].Visible = $False
	$lockedOutUsersList.Columns["AccountExpirationDate"].Visible = $False
	$lockedOutUsersList.Columns["ObjectClass"].Visible = $False
	$lockedOutUsersList.Columns["ObjectGUID"].Visible = $False
	$lockedOutUsersList.Columns["PasswordNeverExpires"].Visible = $False
	$lockedOutUsersList.Columns["SID"].Visible = $False
	$lockedOutUsersList.Columns["PropertyNames"].Visible = $False
	$lockedOutUsersList.Columns["PropertyCount"].Visible = $False
	$lockedOutUsersList.Columns["SamAccountName"].Visible = $False
	
	#Set order of columns
	$lockedOutUsersList.Columns["Name"].DisplayIndex = 0
	$lockedOutUsersList.Columns["UserPrincipalName"].DisplayIndex = 1
	$lockedOutUsersList.Columns["LockedOut"].DisplayIndex = 2
	$lockedOutUsersList.Columns["Enabled"].DisplayIndex = 3
	$lockedOutUsersList.Columns["PasswordExpired"].DisplayIndex = 4
	$lockedOutUsersList.Columns["LastLogonDate"].DisplayIndex = 5

	#Set Column Titles
	$lockedOutUsersList.Columns["UserPrincipalName"].HeaderText = "Email"
	$lockedOutUsersList.Columns["LockedOut"].HeaderText = "Locked"
	$lockedOutUsersList.Columns["PasswordExpired"].HeaderText = "Password Expired"
	$lockedOutUsersList.Columns["LastLogonDate"].HeaderText = "Last Logon Date"
	
	#Set to have columns fill the whole width of the window
	$lockedOutUsersList.Columns["Name"].AutoSizeMode = 'Fill'
	$lockedOutUsersList.Columns["UserPrincipalName"].AutoSizeMode = 'Fill'
	$lockedOutUsersList.Columns["LockedOut"].AutoSizeMode = 'Fill'
	$lockedOutUsersList.Columns["Enabled"].AutoSizeMode = 'Fill'
	$lockedOutUsersList.Columns["PasswordExpired"].AutoSizeMode = 'Fill'
	$lockedOutUsersList.Columns["LastLogonDate"].AutoSizeMode = 'Fill'
	
	#Set Column Fill Weight
	$lockedOutUsersList.Columns["Name"].FillWeight = 100
	$lockedOutUsersList.Columns["UserPrincipalName"].FillWeight = 125
	$lockedOutUsersList.Columns["LockedOut"].FillWeight = 50
	$lockedOutUsersList.Columns["Enabled"].FillWeight = 50
	$lockedOutUsersList.Columns["PasswordExpired"].FillWeight = 50
	$lockedOutUsersList.Columns["LastLogonDate"].FillWeight = 100
	
	#Set Columns Minimum Width
	$lockedOutUsersList.Columns["Name"].MinimumWidth = 100
	$lockedOutUsersList.Columns["UserPrincipalName"].MinimumWidth = 100
	$lockedOutUsersList.Columns["LockedOut"].MinimumWidth = 50
	$lockedOutUsersList.Columns["Enabled"].MinimumWidth = 50
	$lockedOutUsersList.Columns["PasswordExpired"].MinimumWidth = 100
	$lockedOutUsersList.Columns["LastLogonDate"].MinimumWidth = 100
	}

#On Top Checkbox event handler
function OnTopCheckbox_ChangedState {
	if ($OnTopCheckbox.Checked -eq $True)
	{$GUI.Topmost = $True}
	else
	{$GUI.Topmost = $False}
	$GUI.Refresh()
}

#Testing
#Auto Update Checkbox event handler
function AutoUpdateCheckbox_ChangedState {
	if ($AutoUpdateCheckbox.Checked -eq $True)
	{
		$AutoUpdateTimer.Start()
	}
	else
	{
		$AutoUpdateTimer.Stop()
	}
}

#Unlock User
function UnlockUser {
	$UserDistinguishedName = $lockedOutUsersList.Rows[$lockedOutUsersList.SelectedCells[0].RowIndex].Cells["DistinguishedName"].Value
	#Write-Host $UserDistinguishedName
	Unlock-ADAccount -Identity "$UserDistinguishedName"
	#Error Checking
	If($?)
	{
		#Successful
		[System.Windows.Forms.MessageBox]::Show("`"" + $lockedOutUsersList.Rows[$lockedOutUsersList.SelectedCells[0].RowIndex].Cells["Name"].Value + "`"" + " was unlocked." , "Unlocked")
	}
	else
	{
		#Error
		[System.Windows.Forms.MessageBox]::Show("`"" + $lockedOutUsersList.Rows[$lockedOutUsersList.SelectedCells[0].RowIndex].Cells["Name"].Value + "`"" + " was not unlocked." , "Error")
	}
	GUI_RefreshData
}

#Add Event Handlers
$RefreshButton.add_click({GUI_RefreshData})
$OnTopCheckbox.Add_CheckStateChanged({OnTopCheckbox_ChangedState})
$lockedOutUsersList.Add_DoubleClick({UnlockUser})
$AutoUpdateCheckbox.Add_CheckStateChanged({AutoUpdateCheckbox_ChangedState})
$AutoUpdateTimer.add_Tick({GUI_RefreshData})

#Show form
#GUI_RefreshData
$GUI.Add_Shown({GUI_RefreshData})
$GUI.Topmost = $False
[void] $GUI.ShowDialog()
