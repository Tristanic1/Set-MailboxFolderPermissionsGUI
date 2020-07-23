<# 
    .NAME
    Set-MailboxPermissionsGUI

    .SYNOPSIS
    Recursively check, change or remove permissions to Outlook folders

    .DESCRIPTION 
    A GUI script for recursively check, change or remove permissions to Outlook folders.
    This form was created using POSHGUI.com - a free online gui designer for PowerShell

    .NOTES
    Written by: Tomas Cerniauskas
    
    Change Log:
    V0.1, 26/04/2019  - Initial version
    V0.2, 02/05/2019  - Added dynamic resizing of form. Added possibility to save Log.
    V0.3, 15/07/2020  - Added possibility to view and edit "Default" and "Anonymous" permissions
	v0.3.1 23/07/2020 - Fix for German "Default/Anonymous" permission check, Remove is working again
#>

Add-Type -AssemblyName System.Windows.Forms

# Powershell Remoting to Exchange
#$Credential = Get-Credential
$ExchangeServer = 'YourExchangeServer'
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$ExchangeServer/PowerShell/ -Authentication Kerberos -AllowRedirection #-Credential $Credential
Import-PSSession $Session -DisableNameChecking -CommandName Get-Mailbox,Get-MailboxFolderPermission,Set-Mailbox,Set-MailboxFolderPermission,Add-MailboxFolderPermission,Remove-MailboxFolderPermission,Get-ADPermission,Add-ADPermission,Remove-ADPermission,Remove-MailboxPermission,Get-MailboxPermission,Add-MailboxPermission,Get-distributiongroup,Set-Mailbox,Set-MailboxPermission,Get-MailboxFolderStatistics -FormatTypeName *

$ErrorActionPreference = 'SilentlyContinue'
$WarningPreference = 'SilentlyContinue'

# Exchange 2010 Snapin
#Add-PSsnapin Microsoft.Exchange.Management.PowerShell.E2010

# Exchange 2013 Snapin
#Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn

[System.Windows.Forms.Application]::EnableVisualStyles()

# ============== GUI ELEMENTS ================
$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '800,600'
$Form.text                       = 'Set-MailboxFolderPermissionsGUI'
#$Form.FormBorderStyle            = [System.Windows.Forms.FormBorderStyle]::Fixed3D
$form.StartPosition              = [System.Windows.Forms.FormStartPosition]::CenterScreen
#$Form.TopMost                    = $true
$Form.MinimumSize                = '800,600' 

$Panel                 = New-Object 'System.Windows.Forms.TableLayoutPanel'
$Panel.Dock            = 'Fill'
$Panel.RowCount        = 3
$Panel.ColumnCount     = 5
$Panel.RowStyles.Add((New-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 19)))
$Panel.RowStyles.Add((New-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 78)))
$Panel.RowStyles.Add((New-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 3)))
$Panel.ColumnStyles.Add((New-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 19)))
$Panel.ColumnStyles.Add((New-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 14)))
$Panel.ColumnStyles.Add((New-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 35)))
$Panel.ColumnStyles.Add((New-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 19)))
$Panel.ColumnStyles.Add((New-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 13)))
$Panel.CellBorderStyle = 'None'

$MainTextBox                 = New-Object system.Windows.Forms.TextBox
$MainTextBox.multiline       = $true
$MainTextBox.ReadOnly        = $true
$MainTextBox.BackColor       = 'White'
$MainTextBox.width           = 784
$MainTextBox.height          = 670
$MainTextBox.ScrollBars      = 'Vertical'
$MainTextBox.location        = New-Object System.Drawing.Point(1,1)
$MainTextBox.Font            = 'Consolas,11'
$MainTextBox.Dock            = 'Fill'

$MailboxGroupbox                 = New-Object system.Windows.Forms.Groupbox
$MailboxGroupbox.height          = 110
$MailboxGroupbox.width           = 150
$MailboxGroupbox.text            = 'Mailbox to modify'
$MailboxGroupbox.location        = New-Object System.Drawing.Point(1,1)
$MailboxGroupbox.Dock            = 'Fill'

$ProgressBar                    = New-Object system.Windows.Forms.ProgressBar
$ProgressBar.width              = 784
$ProgressBar.height             = 15
$ProgressBar.location           = New-Object System.Drawing.Point(1,1)
$ProgressBar.Dock               = 'Fill'

$AddRemoveGroupbox               = New-Object system.Windows.Forms.Groupbox
$AddRemoveGroupbox.height        = 110
$AddRemoveGroupbox.width         = 105
$AddRemoveGroupbox.text          = 'Modification'
$AddRemoveGroupbox.location      = New-Object System.Drawing.Point(1,1)
$AddRemoveGroupbox.Dock          = 'Fill'


$FolderGroupbox                  = New-Object system.Windows.Forms.Groupbox
$FolderGroupbox.height           = 110
$FolderGroupbox.width            = 274
$FolderGroupbox.text             = 'Folder access rights to add'
$FolderGroupbox.location         = New-Object System.Drawing.Point(1,1)
$FolderGroupbox.Enabled          = $false
$FolderGroupbox.Dock             = 'Fill'

$AccessRightsGroupbox            = New-Object system.Windows.Forms.Groupbox
$AccessRightsGroupbox.height     = 110
$AccessRightsGroupbox.width      = 150
$AccessRightsGroupbox.text       = 'Access Rights'
$AccessRightsGroupbox.location   = New-Object System.Drawing.Point(1,1)
$AccessRightsGroupbox.Enabled    = $false
$AccessRightsGroupbox.Dock       = 'Fill'

$ButtonGroupbox               = New-Object system.Windows.Forms.Groupbox
$ButtonGroupbox.height        = 100
$ButtonGroupbox.width         = 90
$ButtonGroupbox.location      = New-Object System.Drawing.Point(1,1)
$ButtonGroupbox.Anchor        = 'Top'
$ButtonGroupbox.Dock          = 'Fill'

$Username                        = New-Object system.Windows.Forms.Label
$Username.text                   = 'User to add/remove:'
$Username.AutoSize               = $true
$Username.width                  = 25
$Username.height                 = 10
$Username.location               = New-Object System.Drawing.Point(10,43)
$Username.Font                   = 'Microsoft Sans Serif,8'

$DefaultAnonymousCheckBox           = New-Object system.Windows.Forms.CheckBox
$DefaultAnonymousCheckBox.text      = 'Show Default/Anonymous'
$DefaultAnonymousCheckBox.AutoSize  = $true
$DefaultAnonymousCheckBox.width     = 25
$DefaultAnonymousCheckBox.height    = 8
$DefaultAnonymousCheckBox.location  = New-Object System.Drawing.Point(10,85)
$DefaultAnonymousCheckBox.Font      = 'Microsoft Sans Serif,7'
$DefaultAnonymousCheckBox.Enabled   = $true

$CheckButton                          = New-Object system.Windows.Forms.Button
$CheckButton.text                     = 'Check'
$CheckButton.width                    = 88
$CheckButton.height                   = 25
$CheckButton.location                 = New-Object System.Drawing.Point(5,10)
$CheckButton.Font                     = 'Microsoft Sans Serif,10,style=Bold'
$CheckButton.Enabled                  = $true
$CheckButton.FlatStyle                = 'System'

$ModifyButton                          = New-Object system.Windows.Forms.Button
$ModifyButton.text                     = 'Modify'
$ModifyButton.width                    = 88
$ModifyButton.height                   = 25
$ModifyButton.location                 = New-Object System.Drawing.Point(5,40)
$ModifyButton.Font                     = 'Microsoft Sans Serif,10,style=Bold'
$ModifyButton.Enabled                  = $false
$ModifyButton.FlatStyle                = 'System'

$SaveLogButton                          = New-Object system.Windows.Forms.Button
$SaveLogButton.text                     = 'Save Log'
$SaveLogButton.width                    = 88
$SaveLogButton.height                   = 25
$SaveLogButton.location                 = New-Object System.Drawing.Point(5,70)
$SaveLogButton.Font                     = 'Microsoft Sans Serif,10,style=Bold'
$SaveLogButton.Enabled                  = $false
$SaveLogButton.FlatStyle                = 'System'

$MailboxTextBox                  = New-Object system.Windows.Forms.TextBox
$MailboxTextBox.multiline        = $false
$MailboxTextBox.width            = 115
$MailboxTextBox.height           = 20
$MailboxTextBox.location         = New-Object System.Drawing.Point(10,15)
$MailboxTextBox.Font             = 'Microsoft Sans Serif,10'
$MailboxTextBox.Text             = $null

$UserTextBox                     = New-Object system.Windows.Forms.TextBox
$UserTextBox.multiline           = $false
$UserTextBox.width               = 115
$UserTextBox.height              = 20
$UserTextBox.location            = New-Object System.Drawing.Point(10,60)
$UserTextBox.Font                = 'Microsoft Sans Serif,10'
$UserTextBox.Text                = $null

$AddRadioButton                  = New-Object system.Windows.Forms.RadioButton
$AddRadioButton.text             = 'Add'
$AddRadioButton.AutoSize         = $true
$AddRadioButton.width            = 104
$AddRadioButton.height           = 20
$AddRadioButton.location         = New-Object System.Drawing.Point(10,25)
$AddRadioButton.Font             = 'Microsoft Sans Serif,10'

$RemoveRadioButton               = New-Object system.Windows.Forms.RadioButton
$RemoveRadioButton.text          = 'Remove'
$RemoveRadioButton.AutoSize      = $true
$RemoveRadioButton.width         = 104
$RemoveRadioButton.height        = 20
$RemoveRadioButton.location      = New-Object System.Drawing.Point(10,45)
$RemoveRadioButton.Font          = 'Microsoft Sans Serif,10'

$FullAccessRadioButton               = New-Object system.Windows.Forms.RadioButton
$FullAccessRadioButton.text          = 'Full Access'
$FullAccessRadioButton.AutoSize      = $true
$FullAccessRadioButton.width         = 100
$FullAccessRadioButton.height        = 20
$FullAccessRadioButton.location      = New-Object System.Drawing.Point(10,65)
$FullAccessRadioButton.Font          = 'Microsoft Sans Serif,10'

$CompleteMailboxRadioButton           = New-Object system.Windows.Forms.RadioButton
$CompleteMailboxRadioButton.text      = 'Complete mailbox'
$CompleteMailboxRadioButton.AutoSize  = $true
$CompleteMailboxRadioButton.width     = 100
$CompleteMailboxRadioButton.height    = 20
$CompleteMailboxRadioButton.location  = New-Object System.Drawing.Point(10,25)
$CompleteMailboxRadioButton.Font      = 'Microsoft Sans Serif,10'

$SpecificFolderRadioButton          = New-Object system.Windows.Forms.RadioButton
$SpecificFolderRadioButton.text     = 'Specific folder and subfolders'
$SpecificFolderRadioButton.AutoSize = $true
$SpecificFolderRadioButton.width    = 104
$SpecificFolderRadioButton.height   = 20
$SpecificFolderRadioButton.location = New-Object System.Drawing.Point(10,45)
$SpecificFolderRadioButton.Font     = 'Microsoft Sans Serif,10'

$SpecificFolderComboBox               = New-Object system.Windows.Forms.ComboBox
$SpecificFolderComboBox.text          = ''
$SpecificFolderComboBox.width         = 240
$SpecificFolderComboBox.height        = 20
$SpecificFolderComboBox.enabled       = $false
$SpecificFolderComboBox.location      = New-Object System.Drawing.Point(10,66)
$SpecificFolderComboBox.Font          = 'Microsoft Sans Serif,8'
$SpecificFolderComboBox.DropDownStyle = 'DropDownList'

$AccessRightsComboBox               = New-Object system.Windows.Forms.ComboBox
$AccessRightsComboBox.text          = ''
$AccessRightsComboBox.width         = 115
$AccessRightsComboBox.height        = 20
$AccessRightsComboBox.location      = New-Object System.Drawing.Point(9,16)
$AccessRightsComboBox.Font          = 'Microsoft Sans Serif,8'
$AccessRightsComboBox.DropDownStyle = 'DropDownList'

$SendOnBehalfRadioButton            = New-Object system.Windows.Forms.RadioButton
$SendOnBehalfRadioButton.text       = 'Send on Behalf'
$SendOnBehalfRadioButton.AutoSize   = $false
$SendOnBehalfRadioButton.width      = 129
$SendOnBehalfRadioButton.height     = 20
$SendOnBehalfRadioButton.location   = New-Object System.Drawing.Point(9,45)
$SendOnBehalfRadioButton.Font       = 'Microsoft Sans Serif,10'

$SendAsMailboxRadioButton           = New-Object system.Windows.Forms.RadioButton
$SendAsMailboxRadioButton.text      = 'Send as Mailbox'
$SendAsMailboxRadioButton.AutoSize  = $false
$SendAsMailboxRadioButton.width     = 129
$SendAsMailboxRadioButton.height    = 20
$SendAsMailboxRadioButton.location  = New-Object System.Drawing.Point(9,65)
$SendAsMailboxRadioButton.Font      = 'Microsoft Sans Serif,10'

# ============= END OF GUI ELEMENTS =================

# Folder Exclusion list, won't be shown in Drop-down list
$FolderExclusions = @('/Sync Issues',
                      '/Sync Issues/Conflicts',
                      '/Sync Issues/Local Failures',
                      '/Sync Issues/Server Failures',
                      '/Recoverable Items',
                      '/Deletions',
                      '/Purges',
                      '/Versions')

# All possible Folder Access Rights 
$AccessRights = @('Author',
                  'Contributor',
                  'Editor',
                  'None',
                  'NonEditingAuthor',
                  'Owner',
                  'PublishingEditor',
                  'PublishingAuthor',
                  'Reviewer',
                  'CreateItems',
                  'CreateSubfolders',
                  'DeleteAllItems',
                  'DeleteOwnedItems',
                  'EditAllItems',
                  'EditOwnedItems',
                  'FolderContact',
                  'FolderOwner',
                  'FolderVisible',
                  'ReadItems',
                  'AvailabilityOnly',
                  'LimitedDetails')

# Populate the Access Rights Combobox with values from $AccessRights array
$AccessRights | ForEach-Object {[void] $AccessRightsComboBox.Items.Add($_)}

# New Line and Carriage Return
$NewLine = "`r`n"

# Error providers for the TextBox entries
$ErrorProvider1 = New-Object System.Windows.Forms.ErrorProvider
$ErrorProvider2 = New-Object System.Windows.Forms.ErrorProvider
$ErrorProvider3 = New-Object System.Windows.Forms.ErrorProvider
$ErrorProvider4 = New-Object System.Windows.Forms.ErrorProvider

$ToolTip1                        = New-Object system.Windows.Forms.ToolTip
$ToolTip1.isBalloon              = $false
$ToolTip1.SetToolTip($CheckButton,'Check what access rights user(-s) currently has(-ve) for the mailbox')

$MailboxGroupbox.Controls.AddRange(@($Username,$MailboxTextBox,$UserTextBox,$DefaultAnonymousCheckBox))
$AddRemoveGroupbox.Controls.AddRange(@($AddRadioButton,$ModifyButton,$RemoveRadioButton,$FullAccessRadioButton))
$FolderGroupbox.Controls.AddRange(@($CompleteMailboxRadioButton,$SpecificFolderRadioButton,$SpecificFolderComboBox))
$AccessRightsGroupbox.Controls.AddRange(@($AccessRightsComboBox,$SendOnBehalfRadioButton,$SendAsMailboxRadioButton))
$ButtonGroupbox.Controls.AddRange(@($CheckButton,$ModifyButton,$SaveLogButton))

$Panel.Controls.add($MailboxGroupbox,0,0)
$Panel.Controls.add($AddRemoveGroupbox,1,0)
$Panel.Controls.add($FolderGroupbox,2,0)
$Panel.Controls.add($AccessRightsGroupbox,3,0)
$Panel.Controls.add($ButtonGroupbox,4,0)
$Panel.Controls.add($MainTextBox,0,1)
$Panel.SetColumnSpan($MainTextBox, 5)
$Panel.Controls.add($ProgressBar,0,2)
$Panel.SetColumnSpan($ProgressBar, 5)

$Form.Controls.AddRange(@($Panel))

function Enable-GroupBoxes {
  $FolderGroupbox.Enabled          = $true
  $AccessRightsGroupbox.Enabled    = $true
  $AccessRightsComboBox.Enabled    = $true
  $ModifyButton.Enabled            = $true
}

function Disable-GroupBoxes {
  $FolderGroupbox.Enabled              = $false
  $AccessRightsGroupbox.Enabled        = $false
  $AccessRightsComboBox.SelectedIndex  = -1
  $SpecificFolderComboBox.ResetText()
  $SpecificFolderComboBox.Items.Clear()
  $SendOnBehalfRadioButton.Checked     = $false
  $SendAsMailboxRadioButton.Checked    = $false
  $CompleteMailboxRadioButton.Checked  = $false
  $SpecificFolderRadioButton.Checked   = $false
  $ModifyButton.Enabled                = $true
}

function Disable-FolderGroupBox {
  $FolderGroupbox.Enabled               = $false
  $SpecificFolderComboBox.ResetText()
  $SpecificFolderComboBox.Items.Clear()
  $CompleteMailboxRadioButton.Checked   = $false
  $SpecificFolderRadioButton.Checked    = $false
  $AccessRightsGroupbox.Enabled         = $true
  $AccessRightsComboBox.Enabled         = $false
  $AccessRightsComboBox.SelectedIndex   = -1
  $ModifyButton.Enabled                 = $true
}

function Enable-SpecificFolderCombobox {
  $SpecificFolderComboBox.Enabled  = $true
  
}

# Check if Mailbox exists
function Test-Mailbox {

    try {
        $mailbox = Get-mailbox -Identity $MailboxTextBox.Text
    } catch { }
     
    if ($mailbox) {
        $ErrorProvider1.Clear()
    } else { 
        $ErrorProvider1.SetError($MailboxTextBox, 'Enter a valid Mailbox name (Alias)')
    }
    
}

# Check if User exists
function Test-User {

    try {
        $user = Get-mailbox -Identity $UserTextbox.Text
    } catch { }
    
    if (($user) -or ((Get-distributiongroup $UserTextbox.Text).RecipientType -eq 'MailUniversalSecurityGroup') -or ($UserTextbox.TextLength -eq 0) -or ($UserTextbox.Text -match '^Default$|^Anonym$|^Anonymous$|^Standard$')) {
        $ErrorProvider2.Clear()
    } else { 
        $ErrorProvider2.SetError($UserTextbox, 'Enter a valid username (Alias)')
    }
    
}

# Create mailbox folder list in a drop-down menu excluding exceptions
function Get-FolderList {

  Enable-SpecificFolderCombobox
  
  try {
      $mailbox = Get-mailbox -Identity $MailboxTextBox.Text
  } catch { }
      
  if ($mailbox) {
      (Get-MailboxFolderStatistics $MailboxTextBox.Text).FolderPath | Where-Object {!($FolderExclusions -contains $_)} | ForEach-Object { [void] $SpecificFolderComboBox.Items.Add($_) }
      $ErrorProvider3.Clear()
  } else { 
      $ErrorProvider3.SetError($SpecificFolderCombobox, 'Mailbox was not found')
  }
}

function Disable-SpecificFolderComboBox {

  $SpecificFolderComboBox.enabled  = $false
  $SpecificFolderComboBox.ResetText()
  $SpecificFolderComboBox.Items.Clear()
  
}


function Get-MailboxFolderPermissions {
    
    $mailbox = $MailboxTextBox.Text
    $user = $UserTextBox.Text
    $CheckButton.Enabled = $false
    $ModifyButton.Enabled = $false
    $MainTextBox.Focus()
    $ErrorProvider4.Clear()
    $SaveLogButton.Enabled = $true

    # Check if Mailbox TextBox field is not empty
    if ($MailboxTextBox.TextLength -ne 0) {
       
        # Check if Mailbox exists
        if (Get-mailbox -Identity $Mailbox) {             

            # Initialize Progress Bar
            $Progressbar.Value = 0
            $ProgressBar.Maximum =Get-MailboxFolderStatistics -Identity $mailbox | Measure-Object | Select-Object -ExpandProperty Count
                    
            # Check if User TextBox field is empty
            if ($UserTextBox.TextLength -eq 0) {
                
                $MainTextBox.AppendText("[{0}] Checking the existing permissions for all users on mailbox $mailbox..." -f (Get-Date -Format T)+$NewLine)
                
                # Check who has "Send on Behalf" permissions
                $sendonbehalf = Get-Mailbox -Identity $mailbox | Where-Object { $_.GrantSendOnBehalfTo -ne $null } | Select-Object -ExpandProperty GrantSendOnBehalfTo

                ForEach ($user in $sendonbehalf) {
                    # Split the "domain/OU/OU/Display Name" into substring array and select last member "Display Name"
                    $usr = $user.Split('/')[-1]
                    
                    # This line works with SnappIn only
                    #$MainTextBox.AppendText("[{0}] Send on Behalf Permission: $($user.Name)" -f (Get-Date -Format T) +$NewLine)
                    
                    $MainTextBox.AppendText("[{0}] Send on Behalf Permission: $usr" -f (Get-Date -Format T) +$NewLine)
                }
                
                # Check who has "Send as Mailbox" permissions
                $sendasmailbox = Get-Mailbox -Identity $mailbox | Get-ADPermission | Where-Object { ($_.ExtendedRights -like '*send*') -and ($_.IsInherited -eq $false) -and ($_.user.tostring() -ne 'NT AUTHORITY\SELF') -and ($_.user.tostring() -notlike 'S-1-*') } | Select-Object -ExpandProperty User
                
                ForEach ($line in $sendasmailbox) {
                    
                    # Split the "DOMAIN\Username" into two substrings and select second one: "Username"
                    $user = $line.Split('\')[1]
                    
                    # Check if it's user, otherwise it's a group
                    if (Get-Mailbox $user -erroraction SilentlyContinue) {
                        $user = $(Get-Mailbox $user).name
                    }
    
                    $MainTextBox.AppendText("[{0}] Send as Mailbox Permission: $user" -f (Get-Date -Format T)+$NewLine)
                }

                
                # Check who has "Full Mailbox" permissions
                $fullmailboxpermissions = Get-Mailbox -Identity $mailbox | Get-MailboxPermission | Where-Object { ($_.IsInherited -eq $false) -and ($_.user.tostring() -ne 'NT AUTHORITY\SELF') -and ($_.user.tostring() -notlike 'S-1-*') } | Select-Object -ExpandProperty User

                ForEach ($line in $fullmailboxpermissions){ 
                    
                    # Split the "DOMAIN\Username" into two substrings and select second one: "Username"
                    $user = $line.Split('\')[1]
                        
                    # Check if it's user, otherwise it's a group
                    if (Get-Mailbox $user -erroraction SilentlyContinue) {
                        $user = $(Get-Mailbox $user).name
                    }
    
                    $MainTextBox.AppendText("[{0}] Full Mailbox Permission: $user" -f (Get-Date -Format T)+$NewLine)
                }
                

                # Show all Users and their access rights
                ForEach ($f in (Get-MailboxFolderStatistics $Mailbox)) {
                    $Progressbar.Increment(1)
                                
                    if ($DefaultAnonymousCheckBox.Checked -eq $true) {
                        $FolderAccessRights = Get-MailboxFolderPermission -Identity "$mailbox`:$($f.FolderId)" 
                    } else { 
                        $FolderAccessRights = Get-MailboxFolderPermission -Identity "$mailbox`:$($f.FolderId)" | Where-Object {$_.User.DisplayName -notmatch '^Default|^Anonym|^Standard'}
                    }

                               
                    ForEach ($entry in $FolderAccessRights) {
                        $username = $entry.user
                        $AccessRight = $entry.AccessRights
                        
                        $MainTextBox.AppendText("[{0}] " -f (Get-Date -Format T))
                        $MainTextBox.AppendText("$($f.FolderPath)  -  "+"$username  -  "+"$AccessRight "+$NewLine)
                    }
                                
                }


            } else {
                # Check if User exists
                $usercheck = Get-mailbox -Identity $User
                if (($usercheck) -or ((Get-distributiongroup $User).RecipientType -eq 'MailUniversalSecurityGroup') -or ($User -match '^Default$|^Anonym$|^Anonymous$|^Standard$')) {         
                    
                    # Check if Mailbox and User are same
                    if ($Mailbox -eq $User) {                                               
                        $MainTextBox.AppendText("[{0}] Mailbox and User have to be different" -f (Get-Date -Format T)+$NewLine)

                    } else {
                                        
                         $MainTextBox.AppendText("[{0}] Checking the existing permissions for user $user on mailbox $mailbox..."-f (Get-Date -Format T)+$NewLine)
                     
                        ForEach($f in (Get-MailboxFolderStatistics $Mailbox)) {
                            $Progressbar.Increment(1)
                            
                            if ($DefaultAnonymousCheckBox.Checked -eq $true) {
                                $FolderAccessRights = Get-MailboxFolderPermission -Identity "$mailbox`:$($f.FolderId)" 
                            } else { 
                                $FolderAccessRights = Get-MailboxFolderPermission -Identity "$mailbox`:$($f.FolderId)" | Where-Object {$_.User.DisplayName -notmatch 'Default|Anonym|Standard'}
                            }
                            
                                                
                            ForEach ($entry in $FolderAccessRights) {
                                
                                if ($usercheck.name -eq $entry.user) {
                                    $username = $entry.user
                                    $AccessRight = $entry.AccessRights
                                
                                    $MainTextBox.AppendText("[{0}] " -f (Get-Date -Format T))
                                    $MainTextBox.AppendText("$($f.FolderPath)  -  "+"$username  -  "+"$AccessRight "+$NewLine)
                                }
                            }
                        }
                         
                    }
                
                                
                } else {
                    $MainTextBox.AppendText("[{0}] User account " -f (Get-Date -Format T)+$UserTextBox.Text+' was not found'+$NewLine)
                }

            }
        } else { 
            $MainTextBox.AppendText("[{0}] Mailbox "-f (Get-Date -Format T)+$MailboxTextBox.Text+' was not found'+$NewLine)
        }

        $MainTextBox.AppendText("[{0}] COMPLETED!"-f (Get-Date -Format T)+$NewLine)    
        $MainTextBox.AppendText('--------------------------------------------------'+$NewLine)    
    
    # Closing bracket for if ($MailboxTextBox.TextLength -ne 0) {
    }
    $CheckButton.Enabled = $true
    $ModifyButton.Enabled = $true

}

function Set-MailboxFolderPermissions {
    $mailbox = $MailboxTextBox.Text
    $user = $UserTextBox.Text
    $folder = $SpecificFolderComboBox.Text
    $access = $AccessRightsComboBox.Text
    $usercheck = Get-mailbox -Identity $User
    
    $ModifyButton.Enabled  = $false   
    $CheckButton.Enabled   = $false
    $SaveLogButton.Enabled = $true  
    $MainTextBox.Focus()
    $ErrorProvider4.Clear()
        
    # check if Mailbox exists
    if (Get-mailbox -Identity $Mailbox) {
            
        $MailboxIdentity = (Get-Mailbox $mailbox).Identity 
        $MailboxName = (Get-Mailbox $mailbox).Name
            
        # check if User or Group exists              
        if ((($usercheck) -and !($Mailbox -eq $User)) -or ((Get-distributiongroup $User).RecipientType -eq 'MailUniversalSecurityGroup') -or ($User -match '^Default$|^Anonym$|^Anonymous$|^Standard$')) {   
        
                # Initialize Progress Bar
                $Progressbar.Value = 0
                $ProgressBar.Maximum =Get-MailboxFolderStatistics -Identity $mailbox | Measure-Object | Select-Object -ExpandProperty Count
                
                # Check if "Send of Behalf" is ticked    
                if ($SendOnBehalfRadioButton.Checked -eq $true){ 
                    # Needs the .Identity value, otherwise some email addresses are not getting resolved 
                    Set-Mailbox $MailboxIdentity -Grantsendonbehalfto @{add="$user"} -WarningAction SilentlyContinue
                    $MainTextBox.AppendText("[{0}] Granted Send on Behalf permission to user $user"-f (Get-Date -Format T)+$NewLine)
                }

                # Check if "Send As Mailbox" is ticked
                if ($SendAsMailboxRadioButton.Checked -eq $true){
                    Add-ADPermission -Identity $MailboxName -ExtendedRights Send-As -user $user -WarningAction SilentlyContinue
                    $MainTextBox.AppendText("[{0}] Granted Send As Mailbox permission to user $user"-f (Get-Date -Format T)+$NewLine)
                }

                # Check if "Full Access" is ticked
                if ($FullAccessRadioButton.Checked -eq $true){

                    $ErrorProvider4.Clear()
                    Add-MailboxPermission -Identity $MailboxIdentity -User $user -AccessRights FullAccess -WarningAction SilentlyContinue
                    $MainTextBox.AppendText("[{0}] Granted Full Access permission to user $user"-f (Get-Date -Format T)+$NewLine)
                }

                # Check if the the "Remove" access rights is ticked
                if ($RemoveRadioButton.Checked -eq $true) {

                    $MainTextBox.AppendText("[{0}] Removing the existing permissions for user $user on mailbox $mailbox..."-f (Get-Date -Format T)+$NewLine)
                        
                    # Try to remove Send on Behalf Permission                                                                       
                    Set-Mailbox -Identity $MailboxIdentity -Grantsendonbehalfto @{remove="$User"} -WarningAction SilentlyContinue
                    $MainTextBox.AppendText("[{0}] Removing Send on Behalf permission..."-f (Get-Date -Format T)+$NewLine)

                    # Try to remove Send As Mailbox Permission
                    Remove-ADPermission -Identity $MailboxIdentity -User $User -ExtendedRights Send-As -Confirm:$False -WarningAction SilentlyContinue
                    $MainTextBox.AppendText("[{0}] Removing Send As Mailbox permission..."-f (Get-Date -Format T)+$NewLine)
                                                                      
                    # Try to remove Full Access
                    Remove-MailboxPermission -Identity $MailboxIdentity -User $user -AccessRights FullAccess -InheritanceType All -Confirm:$False -WarningAction SilentlyContinue
                    $MainTextBox.AppendText("[{0}] Removing Full Access permissions..."-f (Get-Date -Format T)+$NewLine)


                    ForEach($f in (Get-MailboxFolderStatistics $Mailbox)) {
                        $Progressbar.Increment(1)
                            
                        Remove-MailboxFolderPermission "$mailbox`:$($f.FolderId)" -User $User -Confirm:$False
                            
                        # Has to be separate line, because there are Calendar folders like /Calendar/{091029301293...} and the -f Operator doesn't work properly
                        $MainTextBox.AppendText("[{0}] " -f (Get-Date -Format T))
                        $MainTextBox.AppendText("Removing access rights from folder $($f.FolderPath)"+$NewLine)
                    }

                }

                # Check if the "Add" is ticked 
                if (($AddRadioButton.checked -eq $true) -and (($CompleteMailboxRadioButton.checked -eq $true) -or ($SpecificFolderRadioButton.checked -eq $true))) { 
                      
                    # Check if the "Access Rights" combobox value is not empty
                    if ($AccessRightsComboBox.SelectedIndex -ne -1) { 
                            
                        $access = $AccessRightsComboBox.SelectedItem
                                                                                                        
                        # Check if the "Complete Mailbox" is ticked   
                        if ($CompleteMailboxRadioButton.checked -eq $true) { 
                                                
                            ForEach ($f in (Get-MailboxFolderStatistics $mailbox)) {
                                $Progressbar.Increment(1)

                                Add-MailboxFolderPermission "$mailbox`:$($f.FolderId)" -User $user -AccessRights $access
                                
                                # if existing permission entry already exists - use Set command instead
                                if (!($?)) {
                                    Set-MailboxFolderPermission "$mailbox`:$($f.FolderId)" -User $user -AccessRights $access
                                }
                                $MainTextBox.AppendText("[{0}] " -f (Get-Date -Format T))  
                                $MainTextBox.AppendText("Adding $access access rights to folder $($f.FolderPath) to $user"+$NewLine)
                            }   
                        }

                        # Check if the "Specific Folder" is ticked   
                        if (($SpecificFolderRadioButton.checked -eq $true) -and ($SpecificFolderComboBox.SelectedIndex -ne -1)) { 
                                
                            $folder = $SpecificFolderComboBox.SelectedItem
                            $ProgressBar.Maximum =Get-MailboxFolderStatistics -Identity $mailbox | Where-Object { $_.FolderPath.Contains("$folder") -eq $True } | Measure-Object | Select-Object -ExpandProperty Count
                            $subfolders = $folder.Split('/')
                            $count = $folder.Split('/').count

                            # For Calendar and /Top of Information Store this access right (FolderVisible) is not needed
                            if ($folder -notmatch 'Calendar|Top of Inf') {
                                Add-MailboxFolderPermission $mailbox -User $user -AccessRights FolderVisible

                                # if existing permission entry already exists - use Set command instead
                                if (!($?)) {
                                    Set-MailboxFolderPermission $mailbox -User $user -AccessRights FolderVisible
                                }
                                
                                $MainTextBox.AppendText("[{0}] Adding FolderVisible access rights to folder /Top of Information Store to $user"-f (Get-Date -Format T)+$NewLine)
                            }
                                                                                         
                            $fname = "$mailbox`:"

                            # Recursively grant FolderVisible to all upper level folders
                            for($i=1; $i -lt $count-1; $i++){
    
                                $path = $subfolders[$i].Replace([char]63743,'/')
                                $fname = "$fname" +'\'+"$path"
                                
                                Add-MailboxFolderPermission $fname -User $user -AccessRights FolderVisible
                                # if existing permission entry already exists - use Set command instead
                                if (!($?)) {
                                    Set-MailboxFolderPermission $fname -User $user -AccessRights FolderVisible
                                }
                                
                                $output = $fname.Split(':')[1].Replace('\','/')
                                
                                # Has to be separate line, because there are Calendar folders like /Calendar/{091029301293...} and the -f Operator doesn't work properly
                                $MainTextBox.AppendText("[{0}] " -f (Get-Date -Format T))
                                $MainTextBox.AppendText("Adding FolderVisible access rights to folder $output to $user"+$NewLine)
                            }
    
                            # Recursively grant selected Access Rights to specified folder and subfolders
                            ForEach ($f in (Get-MailboxFolderStatistics $mailbox | Where-Object { $_.FolderPath.Contains("$folder") -eq $True } ) ) {
                                $Progressbar.Increment(1)

                                Add-MailboxFolderPermission "$mailbox`:$($f.FolderId)" -User $user -AccessRights $access
                                # if existing permission entry already exists - use Set command instead
                                if (!($?)) {
                                    Set-MailboxFolderPermission "$mailbox`:$($f.FolderId)" -User $user -AccessRights $access
                                }
                                $output = $($f.FolderPath).Replace([char]63743,'/')
                                    
                                # Has to be separate line, because there are Calendar folders like /Calendar/{091029301293...} and the -f Operator doesn't work properly
                                $MainTextBox.AppendText("[{0}] " -f (Get-Date -Format T))
                                $MainTextBox.AppendText("Adding $access access rights to folder $output to $user"+$NewLine)
                            }
                        }
                            
                    } else {
                        $ErrorProvider4.SetError($AccessRightsCombobox, 'Access Rights not chosen')

                    }
                }

                 
        } else {
            $MainTextBox.AppendText("[{0}] User account "-f (Get-Date -Format T)+$UserTextBox.Text+' was not found'+$NewLine)
        }
    } else { 
        $MainTextBox.AppendText("[{0}] Mailbox "-f (Get-Date -Format T)+$MailboxTextBox.Text+' was not found'+$NewLine)
    }
    $MainTextBox.AppendText("[{0}] COMPLETED!"-f (Get-Date -Format T)+$NewLine)
    $MainTextBox.AppendText('--------------------------------------------------'+$NewLine)  

        
    $ModifyButton.Enabled = $true
    $CheckButton.Enabled = $true
    $CompleteMailboxRadioButton.Checked = $false
}

Function Save-Log {
    $date = Get-Date -Format FileDate
    $Filename = 'Log_'+$date+'.txt'
    $LogFile = Join-Path $PSScriptRoot ($Filename)
    
    if (Test-Path -Path $LogFile){
        Add-Content -Path $LogFile -Value $MainTextBox.Text
        $MainTextBox.AppendText("Log has been saved to file: $Logfile"+"$NewLine")
        $MainTextBox.AppendText('--------------------------------------------------'+"$NewLine")  
    } else {
        New-Item -Path $PSScriptRoot -Name $Filename
        Add-Content -Path $LogFile -Value $MainTextBox.Text
        $MainTextBox.AppendText("Log has been saved to file: $Logfile"+"$NewLine")
        $MainTextBox.AppendText('--------------------------------------------------'+"$NewLine")  
    }
      
}  

$MailboxTextBox.Add_Leave({ Test-Mailbox })
$UserTextBox.Add_Leave({ Test-User })
$AddRadioButton.Add_CheckedChanged({ Enable-GroupBoxes })
$RemoveRadioButton.Add_CheckedChanged({ Disable-GroupBoxes })
$CompleteMailboxRadioButton.Add_CheckedChanged({ Disable-SpecificFolderComboBox })
$SpecificFolderRadioButton.Add_CheckedChanged({ Get-FolderList })
$FullAccessRadioButton.Add_CheckedChanged({ Disable-FolderGroupBox })
$CheckButton.Add_Click({ Get-MailboxFolderPermissions })
$ModifyButton.Add_Click({ Set-MailboxFolderPermissions })
$SaveLogButton.Add_Click({ Save-Log })

[void]$Form.ShowDialog()

Remove-PSSession $Session