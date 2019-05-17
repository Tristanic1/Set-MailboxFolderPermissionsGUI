# Set-MailboxFolderPermissionsGUI
Recursively Set-MailboxFolderPermission with a GUI

# How to Use
Download script https://raw.githubusercontent.com/Tristanic1/Set-MailboxFolderPermissionsGUI/master/Set-MailboxFolderPermissionsGUI.ps1

Modify line 24 of "Set-MailboxFolderPermissionGUI.ps1" with your own Exchange Server name

Can be run in normal user context, but then uncomment the Credential parts on lines 23 and s5

# Features


# Screenshots
Checking all existing permissions on a mailbox:

![Check All Permissions](https://github.com/Tristanic1/Set-MailboxFolderPermissionsGUI/blob/master/img/Check2.gif)

Checking existing permissions on a mailbox for a specific user only:
![Check Single User Permissions](https://github.com/Tristanic1/Set-MailboxFolderPermissionsGUI/blob/master/img/CheckUser2.gif)

Remove user permissions from a mailbox:
![Remvoe User Permissions](https://github.com/Tristanic1/Set-MailboxFolderPermissionsGUI/blob/master/img/Remove2.gif)

Add user permissions on a mailbox:
![Add User Permissions](https://github.com/Tristanic1/Set-MailboxFolderPermissionsGUI/blob/master/img/Add2.gif)

![Add Specific Folder Permissions](https://github.com/Tristanic1/Set-MailboxFolderPermissionsGUI/blob/master/img/Modify2.gif)

# Usage
powershell -ExceutionPolicy Bypass Set-MailboxFolderPermissionsGUI.ps1
