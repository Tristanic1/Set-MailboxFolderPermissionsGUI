# Set-MailboxFolderPermissionsGUI
Recursively Set-MailboxFolderPermission with a GUI for On-Premises Exchange 2013 environment

## How to Use
Download [script](https://github.com/Tristanic1/Set-MailboxFolderPermissionsGUI/archive/master.zip)

Modify line 24 of "Set-MailboxFolderPermissionGUI.ps1" with your own Exchange Server name.

Can be run in normal user context, but then uncomment the Credential parts on lines 23 and 25.

## Features
1.	Check all existing permissions on the mailbox by just entering the mailbox name and pressing button “Check”.
This will show:
    * who has “Send on Behalf” permissions;
    * who has “Send as Mailbox” permissions;
    * who has “Full Access” to mailbox permission;
    * who has what kind of access rights to individual mailbox folders.
2.	Check all existing permissions of one specific user on the mailbox;
3.	Remove “User” access rights from the “Mailbox”;
4.	By selecting “Add” you can either give some access to complete mailbox (all folders) or specific folder with subfolders; additionally you add “Send on Behalf” or “Send as Mailbox” permissions;
5.	Pressing “SaveLog” log saves log to the same folder as the tool is currently located.

## Screenshots/Gifs
Checking all existing permissions on a mailbox:

![Check All Permissions](https://github.com/Tristanic1/Set-MailboxFolderPermissionsGUI/blob/master/img/Check.gif)



Checking existing permissions on a mailbox for a specific user only:

![Check Single User Permissions](https://github.com/Tristanic1/Set-MailboxFolderPermissionsGUI/blob/master/img/CheckUser1.gif)



Remove user permissions from a mailbox:

![Remove User Permissions](https://github.com/Tristanic1/Set-MailboxFolderPermissionsGUI/blob/master/img/Remove1.gif)



Add user permissions on a mailbox:

![Add Specific Folder Permissions](https://github.com/Tristanic1/Set-MailboxFolderPermissionsGUI/blob/master/img/Add1.gif)

![Add User Permissions](https://github.com/Tristanic1/Set-MailboxFolderPermissionsGUI/blob/master/img/Modify1.gif)


## Version history
*    v0.1, 26/04/2019 - Initial version
*    v0.2, 02/05/2019 - Added dynamic resizing of form. Added possibility to save Log.

## Known issues
1.	Minimum requirement is PowerShell 3, but tested only with 5.1
2.	Does not Overwrite the existing permissions (if user has “FolderVisible”, and you grant “Editor” – the “FolderVisible” stays). Workaround is to Remove all existing access rights and re-add them with new rights.
3. You cannot remove legacy users, who no longer exist in AD
