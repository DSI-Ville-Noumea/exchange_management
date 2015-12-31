
<#
.SYNOPSIS
Add-CustomMailboxPerm.ps1

.DESCRIPTION 
A proof of concept script for adding mailbox permissions and mailbox folder
permissions to all folders in a mailbox.

.OUTPUTS
Console output for progress.

.PARAMETER Mailbox
The mailbox that the folder permissions will be added to.

.PARAMETER User
The user you are granting mailbox folder permissions to.

.PARAMETER AccessRights
The permissions to grant to the mailbox top level.

The AccessRights parameter specifies the rights needed to perform the operation. Valid values include:
FullAccess
ExternalAccount
DeleteItem
ReadPermission
ChangePermission
ChangeOwner

.PARAMETER FolderAccessRights
The permissions to grant for each folder.

The roles that are available, along with the permissions that they assign, are described in the following list:

Author                CreateItems, DeleteOwnedItems, EditOwnedItems, FolderVisible, ReadItems
Contributor           CreateItems, FolderVisible
Editor                CreateItems, DeleteAllItems, DeleteOwnedItems, EditAllItems, EditOwnedItems, FolderVisible, ReadItems
None                  FolderVisible
NonEditingAuthor      CreateItems, FolderVisible, ReadItems
Owner                 CreateItems, CreateSubfolders, DeleteAllItems, DeleteOwnedItems, EditAllItems, EditOwnedItems, FolderContact, FolderOwner, FolderVisible, ReadItems
PublishingEditor      CreateItems, CreateSubfolders, DeleteAllItems, DeleteOwnedItems, EditAllItems, EditOwnedItems, FolderVisible, ReadItems
PublishingAuthor      CreateItems, CreateSubfolders, DeleteOwnedItems, EditOwnedItems, FolderVisible, ReadItems
Reviewer              FolderVisible, ReadItems

.EXAMPLE
.\Add-CustomMailboxPerm.ps1 -Mailbox db.acsp -UserOrGroup DB-GG-RO-BAL-ACSP -AccessRights ReadPermission -FolderAccessRights Reviewer

This will grant DB-GG-RO-BAL-ACSP "ReadPermission" to the top levek of db.acsp's mailbox and "Reviewer" access to all folders in db.acsp's mailbox.

.NOTES
Written by: Romain Zanon

Change Log:
V1.0, 28/12/2015 - Clone version from https://github.com/cunninghamp/Powershell-Exchange/blob/master/MailboxFolderPermissions/Add-MailboxFolderPermissions.ps1
V2.0, 28/12/2015 - Add new param, tests, import-module, mailboxpermissions cmd-let, list of available access rights for mailbox and mailbox folders
v2.1, 29/12/2015 - Modify folder names for MS Exchange 2010 french version. Add applyFolderAccessRights function. Renamed to Add-CustomMailboxPerm.ps1.
#>

#requires -version 2

[CmdletBinding()]
param (
	[Parameter( Mandatory=$true)]
	[string]$Mailbox,
    
	[Parameter( Mandatory=$true)]
	[string]$UserOrGroup,
    
  	[Parameter( Mandatory=$true)]
	[string]$AccessRights,

    [Parameter( Mandatory=$true)]
	[string]$FolderAccessRights
)


#...................................
# Variables
#................................... 

$exclusions = @("/Problèmes de synchronisation",
                "/Problèmes de synchronisation/Conflits",
                "/Problèmes de synchronisation/Défaillances du serveur",
                "/Problèmes de synchronisation/Défaillances locales",
                "/Recoverable Items",
                "/Deletions",
                "/Purges",
                "/Versions",
                "/Calendrier",
                "/Tâches",
                "/Contacts",
                "/Contacts suggérés"
                )


#...................................
# Initialize
#...................................

if ("$AccessRights" -ne "ReadPermission" -and "$AccessRights" -ne "FullAccess" -and "$AccessRights" -ne "ExternalAccount" -and "$AccessRights" -ne "DeleteItem" -and "$AccessRights" -ne "ChangePermission" -and "$AccessRights" -ne "ChangeOwner") {
    Write-Host -BackgroundColor Black -ForegroundColor Red "Wrong permissions. Access rights should be ReadPermission, FullAccess, ExternalAccount, DeleteItem, ChangePermission or ChangeOwner."
    exit 10
}

if ("$FolderAccessRights" -ne "Author" -and "$FolderAccessRights" -ne "Contributor" -and "$FolderAccessRights" -ne "Editor" -and "$FolderAccessRights" -ne "None" -and "$FolderAccessRights" -ne "NonEditingAuthor" -and "$FolderAccessRights" -ne "Owner" -and "$FolderAccessRights" -ne "PublishingEditor" -and "$FolderAccessRights" -ne "PublishingAuthor" -and "$FolderAccessRights" -ne "Reviewer") {
    Write-Host -BackgroundColor Black -ForegroundColor Red "Wrong permissions. Folder access rights should be Author, Contributor, Editor, None, NonEditingAuthor, Owner, PublishingEditor, PublishingAuthor or Reviewer."
    exit 10
}

# Import Active Directory module
# Check AD module
write-host "[INFO] - Check if PowerShell AD module is loaded..." -ForegroundColor DarkCyan
if ( -not (Get-Module ActiveDirectory)) {
	write-host "[INFO] - ActiveDirectory module not loaded." -ForegroundColor DarkCyan
	if ( Get-Module -ListAvailable ActiveDirectory ) {
		write-host "[INFO] - AD module available. Loading..." -ForegroundColor DarkCyan
		Import-Module ActiveDirectory
		sleep 1
		if ( Get-Module ActiveDirectory ) {
			write-host "[OK] - ActiveDirectory module loaded successfully." -ForegroundColor Green
		} else {
			write-error "Failed to load ActiveDirectory module. Exiting.."
			exit 1
		}
	} else {
		write-error "You must install ActiveDirectory PowerShell module first before using this script. Exiting.."
		exit 2
	}
} else {
	write-host "[OK] - ActiveDirectory module loaded"  -ForegroundColor Green
}


#Add Exchange 2010 snapin if not already loaded in the PowerShell session
if (!(Get-PSSnapin | where {$_.Name -eq "Microsoft.Exchange.Management.PowerShell.E2010"}))
{
	try
	{
		Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction STOP
	}
	catch
	{
		#Snapin was not loaded
		Write-Warning $_.Exception.Message
		exit 1
	}
}

# Check AD object class (user or group)
Write-Host "[INFO] - Check valid AD user or group..." -ForegroundColor DarkCyan
Try { 
    $objectClass = (Get-ADUser $UserOrGroup).ObjectClass
    Write-Host "[OK] - $UserOrGroup user is valid." -ForegroundColor Green
} catch {
    Try { 
    $objectClass = (Get-ADGroup $UserOrGroup).ObjectClass
    Write-Host "[OK] - $UserOrGroup group is valid." -ForegroundColor Green
    } catch {
        $objectClass = "Wrong Identify"
        Write-Host -BackgroundColor Black -ForegroundColor Red "Invalid AD user or group. Please try again with a valid Active Directory user or group. Exiting.."
        exit 10
    }
}



#...................................
# Script
#...................................

Write-Host " "

Write-Warning "The following permissions will be applied: "
    Write-Host "$AccessRights rights will be granted on the top level $Mailbox mailbox to $UserOrGroup." -ForegroundColor DarkCyan
    Write-Host "$FolderAccessRights rights will be granted on each $Mailbox mailbox folder to $UserOrGroup." -ForegroundColor DarkCyan
    $answer = Read-Host "Confirm? [Yes|No]" -Verbose


Write-Host " "

# If answer is YES, perform the following tasks
if ( $answer -eq "Yes" ) {

#### Step 1 : Mailbox Access Rights
# Apply mailbox permissions
Write-Host "Adding $UserOrGroup to $Mailbox with $AccessRights permissions..."
    try
    {
        Add-MailboxPermission -Identity $Mailbox -User $UserOrGroup -AccessRights $AccessRights -ErrorAction Continue
    }
    catch
    {
        Write-Warning $_.Exception.Message
        exit 5
    }

$mailboxfolders = @(Get-MailboxFolderStatistics $Mailbox | Where {!($exclusions -icontains $_.FolderPath)} | Select FolderPath)


#### Step 2 : Mailbox Folder Access Rights
# Create applyFolderAccessRights function to apply access rights for each mailbox folder.
function applyFolderAccessRights() {

    [CmdletBinding()]
    param (
	    [Parameter( Mandatory=$true)]
	    [string]$ADMember
    )

    foreach ($mailboxfolder in $mailboxfolders)
    {
        $folder = $mailboxfolder.FolderPath.Replace("/","\")
	    $mailboxTopLevel = "Partie supérieure de la banque d'informations"
        if ($folder -match $mailboxTopLevel)
        {
            $folder = $folder.Replace(“\$mailboxTopLevel”,”\”)
        }
        $identity = "$($Mailbox):$folder"
        Write-Host "Adding $UserOrGroup to $identity with $FolderAccessRights folder permissions..."
        try
        {
            Add-MailboxFolderPermission -Identity $identity -User $ADMember -AccessRights $FolderAccessRights -ErrorAction Continue
        }
        catch
        {
            Write-Warning $_.Exception.Message
            exit 6
        }

    }
}
    
# Apply mailbox folder permissions
if ( $objectClass -eq "user" ) {

    applyFolderAccessRights -ADMember $UserOrGroup
    
} else {
    
    # Foreach each loop to get group members and grant access rights to each one.
    foreach ($adUser in (Get-ADGroupMember $UserOrGroup)) {

        applyFolderAccessRights -ADMember $adUser.SamAccountName

    }

}

#End If
}

exit $?

#...................................
# End
#...................................