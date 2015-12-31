
<#
.SYNOPSIS
Remove-CustomMailboxPerm.ps1

.DESCRIPTION 
A proof of concept script for removing mailbox permissions and mailbox folder
permissions to all folders in a mailbox.

.OUTPUTS
Console output for progress.

.PARAMETER Mailbox
The mailbox that the folder permissions will be removed to.

.PARAMETER User
The user you are removing/revoking mailbox folder permissions to.

.PARAMETER AccessRights
The permissions to remove to the mailbox top level.

The AccessRights parameter specifies the rights needed to perform the operation. Valid values include:
All                   All rights will be revoked. Default value.
FullAccess
ExternalAccount
DeleteItem
ReadPermission
ChangePermission
ChangeOwner
None                  Keep mailbox access rights as it. No access rights will be revoked.

.PARAMETER FolderAccessRights
The permissions to revoke for each folder.

All                   All rights revoked. Default value.
None                  Keep mailbox folder access rights as it. No access rights will be removed.

.EXAMPLE
.\Remove-CustomMailboxPerm.ps1 -Mailbox db.acsp -UserOrGroup DB-GG-RO-BAL-ACSP -AccessRights ReadPermission

This will remove DB-GG-RO-BAL-ACSP "ReadPermission" to the top level of db.acsp's mailbox and all access to all folders in db.acsp's mailbox.

.\Remove-CustomMailboxPerm.ps1 -Mailbox db.acsp -UserOrGroup DB-GG-RO-BAL-ACSP

This will remove all DB-GG-RO-BAL-ACSP permissions, both on db.acsp's mailbox and db.acsp's mailbox folders

.NOTES
Written by: Romain Zanon

Change Log:
V1.0, 29/12/2015 - Initial version
#>

#requires -version 2

[CmdletBinding()]
param (
	[Parameter( Mandatory=$true)]
	[string]$Mailbox,
    
	[Parameter( Mandatory=$true)]
	[string]$UserOrGroup,
    
  	[Parameter( Mandatory=$false)]
	[string]$AccessRights = "All",

    [Parameter( Mandatory=$false)]
	[string]$FolderAccessRights = "All"
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

if ("$AccessRights" -ne "All" -and "$AccessRights" -ne "ReadPermission" -and "$AccessRights" -ne "FullAccess" -and "$AccessRights" -ne "ExternalAccount" -and "$AccessRights" -ne "DeleteItem" -and "$AccessRights" -ne "ChangePermission" -and "$AccessRights" -ne "ChangeOwner") {
    Write-Host -BackgroundColor Black -ForegroundColor Red "Wrong permissions. Access rights should be All, ReadPermission, FullAccess, ExternalAccount, DeleteItem, ChangePermission or ChangeOwner. Default value is All"
    exit 10
}

if ("$FolderAccessRights" -ne "All" -and "$FolderAccessRights" -ne "None") {
    Write-Host -BackgroundColor Black -ForegroundColor Red "Wrong permissions. Folder access rights should be All or None. Default value is All."
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

Write-Warning "The following permissions will be removed: "
    Write-Host "$AccessRights rights will be revoked on the top level $Mailbox mailbox to $UserOrGroup." -ForegroundColor DarkCyan
    Write-Host "$FolderAccessRights rights will be revoked on each $Mailbox mailbox folder to $UserOrGroup." -ForegroundColor DarkCyan
    $answer = Read-Host "Confirm? [Yes|No]" -Verbose


Write-Host " "

# If answer is YES, perform the following tasks
if ( $answer -eq "Yes" ) {

#### Step 1 : Mailbox Access Rights
# Remove mailbox permissions
Write-Host "Removing $UserOrGroup to $Mailbox with $AccessRights permissions..."

if ( $AccessRights -ne "All" ) {

    try
    {
        Remove-MailboxPermission -Identity $Mailbox -User $UserOrGroup -AccessRights $AccessRights -Confirm:$false -ErrorAction Continue
    }
    catch
    {
        Write-Warning $_.Exception.Message
        exit 5
    }

} else {

    try
    {
        Get-MailboxPermission -Identity $Mailbox -User $UserOrGroup | Remove-MailboxPermission -Confirm:$false -ErrorAction Continue
    }
    catch
    {
        Write-Warning $_.Exception.Message
        exit 5
    }

}

Write-Host " "

#### Step 2 : Mailbox Folder Access Rights

# Get all mailbox folders except those excluded
$mailboxfolders = @(Get-MailboxFolderStatistics $Mailbox | Where {!($exclusions -icontains $_.FolderPath)} | Select FolderPath)

# Create applyFolderAccessRights function to apply access rights for each mailbox folder.
function removeFolderAccessRights() {

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
        Write-Host "Removing $UserOrGroup to $identity with $FolderAccessRights folder permissions..." 

        Try {

            $perm = Get-MailboxFolderPermission -Identity $identity -User $ADMember -ErrorAction SilentlyContinue
         
            if ($perm -ne $null) {
                Write-Host "The following access rights are currently applied and are going to be removed: " -ForegroundColor DarkCyan
                foreach ($obj in $(Get-MailboxFolderPermission -Identity $identity -User $ADMember -ErrorAction SilentlyContinue)) { $obj.AccessRights }

                if ( $FolderAccessRights -ne "All" ) {

                    try
                    {
                        Remove-MailboxFolderPermission -Identity $identity -User $ADMember -Confirm:$false -ErrorAction Continue
                    }
                    catch
                    {
                        Write-Warning $_.Exception.Message
                        exit 6
                    }

                } else {

                    try
                    {
                        Remove-MailboxFolderPermission -Identity $identity -User $ADMember -Confirm:$false -ErrorAction Continue
                    }
                    catch
                    {
                        Write-Warning $_.Exception.Message
                        exit 6
                    }

                }

            } else {

                Write-Host "No access right on folder $identity for $ADMember user." -ForegroundColor DarkCyan

            }

        } catch { Write-Host "No access right on folder $identity for $ADMember user." -ForegroundColor DarkCyan}
    }
}
    
# Apply mailbox folder permissions
if ($FolderAccessRights -eq "All") {

    if ( $objectClass -eq "user" ) {

        removeFolderAccessRights -ADMember $UserOrGroup
    
    } else {
    
        # Foreach each loop to get group members and grant access rights to each one.
        foreach ($adUser in (Get-ADGroupMember $UserOrGroup)) {

            removeFolderAccessRights -ADMember $adUser.SamAccountName

        }

    }

} else {

    Write-Host "[INFO] - FolderAccessRights has been set to None. Skipping folder access rights revocation process.." -ForegroundColor DarkCyan

}

#End If
}

exit $?

#...................................
# End
#...................................