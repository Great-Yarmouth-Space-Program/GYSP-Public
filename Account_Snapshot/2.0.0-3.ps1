<#
    Name:               2.0.0-3.PS1
    Version:            2.0.0
    Date:               27-09-2023
    Original Author:    Si Ansell
    Email:              graph@greatyarmouthspaceprogram.space

    Use:                Creates a Quick Account Excel report to assist with scoping a Microsoft 365 Tenant to Tenant Migration 

    Files:              Template Excel file MUST be in the working directory

    Modules:            Install-Module -Name ExchangeOnlineManagement
                        Install-Module -Name ImportExcel
                        Install-Module -Name Microsoft.Online.SharePoint.PowerShell
                        Install-Module Microsoft.Graph -Scope AllUsers

    Updates:            

#>

# Parameters

#region Functions
#region Check for Modules Function
function CheckForPowerShellModule([string]$ModuleName) {
    Write-Host "Checking for $ModuleName module..."
    if (Get-Module -ListAvailable -Name $ModuleName) {
        # Module already installed 
        Write-Host "$ModuleName installed - Continuing" -ForegroundColor Green
    }
    else {
        # Module not installed
        Write-Host "$ModuleName not found. Please install module and rerun script" -ForegroundColor Red
        #Exit
    }
}
#endregion

#region Open Report
Function Show-Report {
    Export-Excel $Output -show
}

#endregion
#endregion

#region Check for Modules
# Check for the presence of PowerShell modules
CheckForPowerShellModule("Microsoft.Graph")
CheckForPowerShellModule("ExchangeOnlineManagement")
CheckForPowerShellModule("ImportExcel")
CheckForPowerShellModule("Microsoft.Online.SharePoint.PowerShell")
#endregion

#region Connect to Microsoft 365
# Disconnects Microsfoft Graph to ensure connection to correct tenant
Write-Host "Disconnecting Microsoft Graph" -ForegroundColor Red
Disconnect-MgGraph
# Connect to Microsoft Graph with specified scopes
Write-Host "Connecting to Microsoft Graph" -ForegroundColor Green
Connect-MGGraph -Scopes User.Read.All, Group.Read.All, OrgContact.Read.All, Device.Read.All, Policy.Read.All, Application.Read.All, Organization.Read.All
# Connect to Exchange Online
Write-Host "Connecting to Exchange Online" -ForegroundColor Green
Connect-ExchangeOnline
# Retrieve and process accepted domains
$OnMicrosoftDomain = Get-AcceptedDomain | Where-Object { $_.DomainName -like "*.onmicrosoft.com" } | Select-Object DomainName -ExpandProperty DomainName
$OnMicrosoftPrefix = $OnMicrosoftDomain.split('.')[0]
# Construct the admin SharePoint URL and connect
$AdminURL = "https://" + $OnMicrosoftPrefix + "-admin.sharepoint.com"
Write-Host "Connecting to SharePoint" -ForegroundColor Green
Connect-SPOService -URL $AdminURL
#endregion

##### Script Start #####

#region Script
#region Manage Template and Paths
# Get organization information using the Microsoft Graph module
$Org = Get-MgOrganization

# Get the current date and format it
$Date = Get-Date -format 'yyyyMMdd_HHmmss'

# Create the output filename based on the date
$OrgDisplayName = $Org.Displayname
$OrgDisplayName = $OrgDisplayName -replace " ", "_"

$Output = $OrgDisplayName + "_Tenant_Assessment-" + $Date + ".xlsx"

# Define the template filename
$Template = "Quick-Snapshot-Template.xlsx"

if (-not (Test-Path $Template)) {
    Write-Host "Error: Template $Template does not exist." -ForegroundColor Red
    return
}
try {
    Write-Host "Copying Template to $Output"
    Copy-Item $Template $Output
}
catch {
    Write-Host "An error occurred: $_" -ForegroundColor Red
}

# Create the transcript filename based on the date
$Transcript = $OrgDisplayName + "_Tenant_Assessment-" + $Date + "-Transcript.Log"

# Start a transcript of the PowerShell session
Start-Transcript -Path $Transcript
#endregion

# Get Master Data and store in variables
$MGUsers = Get-MGUser -All | Sort-Object displayname
$OneDrives = Get-SPOSite -IncludePersonalSite $true -Limit all -Filter "Url -like '-my.sharepoint.com/personal/'" | Sort-Object Title
$M365Groups = Get-UnifiedGroup -ResultSize Unlimited | Sort-Object DisplayName
$DistributionGroups = Get-DistributionGroup -ResultSize Unlimited | Sort-Object DisplayName
$DynamicDistributionGroups = Get-DynamicDistributionGroup -ResultSize Unlimited | Sort-Object DisplayName
$MGContacts = Get-MGContact -All | Sort-Object Displayname
$MGGroups = Get-MgGroup

# Get and process MG user data
$MGUsersArray = @()

$i = 1
$ErrorActionPreference = “silentlycontinue”  
# Loop through each user in the list of MGUsers.
Foreach ($MGUser in $MGUsers) {

    # If there are multiple users, show a progress bar.
    If ($MGUsers.Count -gt "1") {
        Write-Progress -ID 1 -Activity "Processing All MG User Accounts" -Status "User Account $i of $($MGUsers.Count)" -PercentComplete (($i / $MGUsers.Count) * 100)
    }

    # Initialize an empty array for storing mailbox information.
    $UserMailbox = @()

    # Retrieve properties for the current user.
    $MGUserProperties = Get-MgUser -UserID $MGUser.ID -Property id, userprincipalname, displayname, mail, licenseAssignmentStates, usagelocation, usertype, accountenabled, OnPremisesSyncEnabled, AssignedLicenses, assignedplans, Activities | Select-Object id, userprincipalname, displayname, mail, licenseAssignmentStates, usagelocation, usertype, accountenabled, OnPremisesSyncEnabled., AssignedLicenses, assignedplans, Activities
    
    # Get the mailbox associated with the current user.
    $UserMailbox = Get-Mailbox -Identity $MGUser.UserPrincipalName
    
    # If a mailbox exists for the user, retrieve mailbox statistics.
    If ($UserMailbox -ne $Null) {
        $MailboxStats = Get-MailboxStatistics -Identity $UserMailbox.UserPrincipalName | Select-Object LastLogonTime, DisplayName, @{Name = "TotalItemSizeMB"; Expression = { [math]::Round(($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 0) } }, ItemCount, DeletedItemCount, @{Name = "TotalDeletedItemSizeMB"; Expression = { [math]::Round(($_.TotalDeletedItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 0) } }
        $HasMailbox = "Active"
    } 

    # If no mailbox exists for the user, set the HasMailbox variable to an empty string.
    If ($UserMailbox -eq $Null) {
        $HasMailbox = ""
    } 
    
    # If the user has an active archive mailbox, retrieve its statistics.
    If ($UserMailbox.ArchiveStatus -eq "Active") {
        $ArchiveStats = Get-MailboxStatistics -Identity $UserMailbox.UserPrincipalName -Archive | Select-Object DisplayName, @{Name = "TotalItemSizeMB"; Expression = { [math]::Round(($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 0) } }, ItemCount, DeletedItemCount, @{Name = "TotalDeletedItemSizeMB"; Expression = { [math]::Round(($_.TotalDeletedItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 0) } }
        $HasArchive = "Active"
    } 

    # If the user does not have an active archive mailbox, set the ArchiveStats and HasArchive variables to empty strings.
    If ($UserMailbox.ArchiveStatus -ne "Active") {
        $ArchiveStats = ""
        $HasArchive = ""
    }

    # Initialize an empty variable to store the OneDrive status for the user.
    $HasOneDrive = ""

    # Filter the OneDrives to get the OneDrive that belongs to the current MGUser.
    $UserOneDrive = $OneDrives | Where-Object { $_.Owner -eq $MGUser.UserPrincipalName }

    # Check if the user has an associated OneDrive.
    If ($UserOneDrive -ne $Null) {

        # Retrieve and sort the ServicePlans for the current MGUser.
        $ServicePlans = (Get-MgUserLicenseDetail -UserId $MGUser.Id -Property ServicePlans).ServicePlans | Sort-Object ServicePlanName

        # Check if the ServicePlans are not null, which implies the OneDrive is active.
        If ($ServicePlans -ne $Null) {
            $HasOneDrive = "Active"
        }

        # Check if the ServicePlans are null, which implies the OneDrive is inactive.
        If ($ServicePlans -eq $Null) {
            $HasOneDrive = "Inactive"
        }
    }

    # Check if the user does not have an associated OneDrive.
    If ($UserOneDrive -eq $Null) {
        $HasOneDrive = ""
    }

    # Check the RecipientTypeDetails property of the UserMailbox object to determine the type of mailbox or user.

    # Check if the mailbox type is a regular user mailbox.
    If ($UserMailbox.RecipientTypeDetails -eq "UserMailbox") {
        $Type = "Mail Enabled User"
    }

    # Check if the mailbox type is a shared mailbox.
    If ($UserMailbox.RecipientTypeDetails -eq "SharedMailbox") {
        $Type = "SharedMailbox"
    }

    # Check if the mailbox type is a room mailbox.
    If ($UserMailbox.RecipientTypeDetails -eq "RoomMailbox") {
        $Type = "RoomMailbox"
    }

    # Check if the mailbox type is an equipment mailbox.
    If ($UserMailbox.RecipientTypeDetails -eq "EquipmentMailbox") {
        $Type = "EquipmentMailbox"
    }

    # Check if the mailbox type is a guest.
    If ($UserMailbox.RecipientType -eq "Guest") {
        $Type = "Guest"
    }

    # Check if the mailbox type is null, which might indicate an unlicensed user.
    If ($UserMailbox.RecipientType -eq $Null) {
        $Type = "Unlicensed User?"
    }

    # Check the UserPrincipalName of the Mguser object to determine if the user is a guest .
    If ($Mguser.UserPrincipalName -like "*#EXT#@*") {
        $Type = "Guest"
    }



    # Get License SKUs
    $LicenseSKUs = Get-MgUserLicenseDetail -UserId $MGUser.ID 

    # Add user data to the array
    $MGUsersArray = $MGUsersArray + [PSCustomObject]@{
        DisplayName       = $MGUser.DisplayName ;
        UserPrincipalName = $MGUser.UserPrincipalName ;
        Mail              = $MGUserProperties.Mail ;
        Type              = $Type ;
        Mailbox           = $Hasmailbox ;
        Archive           = $HasArchive ;
        OneDrive          = $HasOneDrive ;
        MailboxSizeMB     = $MailboxStats.TotalItemSizeMB ;
        ArchiveSizeMB     = $ArchiveStats.TotalItemSizeMB ;
        OneDriveSize      = $UserOneDrive.StorageUsageCurrent ;
        UsageLocation     = $MGUserProperties.UsageLocation ;
        AccountEnabled    = $MGUserProperties.AccountEnabled ;
        IsDirSynced       = $MGUserProperties.OnPremisesSyncEnabled ;
        Members           = "" ;
        GuestMembers      = "" ;
        ID                = $MGUser.ID;
        LicenseSKUs       = $LicenseSKUs.SkuPartNumber -join ";" ;  

    }
    $i++
} 

$i = 1

foreach ($M365Group in $M365Groups) {
    If ($M365Groups.Count -gt "1") {
        Write-Progress -ID 1 -Activity "Processing Microsoft Groups and Teams" -Status "Group $i of $($M365Groups.Count)" -PercentComplete (($i / $M365Groups.Count) * 100)
    }
    
    # Get M365 mailbox statistics
    $M365MailboxStats = Get-MailboxStatistics -Identity $M365Group.PrimarySMTPAddress | Select-Object LastLogonTime, DisplayName, @{Name = "TotalItemSizeMB"; Expression = { [math]::Round(($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 0) } }, ItemCount, DeletedItemCount, @{Name = "TotalDeletedItemSizeMB"; Expression = { [math]::Round(($_.TotalDeletedItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 0) } }
    If ($M365Group.ResourceProvisioningOptions -like "*Team*") {
        $Type = "Microsoft Team" 
    }   
    Else {
        $Type = "Microsoft 365 Group" 
    }
    # Add M365 Group Mailbox data to the array
    $MGUsersArray = $MGUsersArray + [PSCustomObject]@{
        DisplayName       = $M365Group.DisplayName ;
        UserPrincipalName = "" ;
        Mail              = $M365Group.PrimarySmtpAddress ;
        Type              = $Type ;
        Mailbox           = "Active" ;
        Archive           = "" ;
        OneDrive          = "" ;
        MailboxSizeMB     = $M365MailboxStats.TotalItemSizeMB ;
        ArchiveSizeMB     = "" ;
        OneDriveSize      = "" ;
        UsageLocation     = "" ;
        AccountEnabled    = "" ;
        IsDirSynced       = "" ;
        Members           = $M365Group.GroupMemberCount ;
        GuestMembers      = $M365Group.GroupExternalMemberCount ;
        ID                = $M365Group.ID;
        LicenseSKUs       = "" ;  

    }    
    $i++
}

$i = 1
foreach ($DistributionGroup in $DistributionGroups) {
    If ($DistributionGroups.Count -gt "1") {
        Write-Progress -ID 1 -Activity "Processing Distribution Groups" -Status "Distribution Group $i of $($DistributionGroups.Count)" -PercentComplete (($i / $DistributionGroups.Count) * 100)
    }
    If ($DistributionGroup.RecipientTypeDetails -eq "MailUniversalDistributionGroup") {
        $Type = "Distribution List"
    }
    If ($DistributionGroup.RecipientTypeDetails -eq "RoomList") {
        $Type = "Room List"
    }


    $DistributionGroupMembers = Get-DistributionGroupMember -ResultSize Unlimited -Identity $DistributionGroup.Name | Measure-Object

    # Add Disytribution Group data to the array
    $MGUsersArray = $MGUsersArray + [PSCustomObject]@{
        DisplayName       = $DistributionGroup.DisplayName ;
        UserPrincipalName = "" ;
        Mail              = $DistributionGroup.WindowsEmailAddress ;
        Type              = $Type ;
        Mailbox           = "" ;
        Archive           = "" ;
        OneDrive          = "" ;
        MailboxSizeMB     = "" ;
        ArchiveSizeMB     = "" ;
        OneDriveSize      = "" ;
        UsageLocation     = "" ;
        AccountEnabled    = "" ;
        IsDirSynced       = $DistributionGroup.IsDirSynced ;
        Members           = $DistributionGroupMembers.Count ;
        GuestMembers      = "" ;
        ID                = $DistributionGroup.ID;
        LicenseSKUs       = "" ;  
    }    
    $i++
}

$i = 1
foreach ($DynamicDistributionGroup in $DynamicDistributionGroups) {
    If ($DynamicDistributionGroups.Count -gt "1") {
        Write-Progress -ID 1 -Activity "Processing Dynamic Distribution Groups" -Status "Dynamic Distribution Group $i of $($DynamicDistributionGroups.Count)" -PercentComplete (($i / $DynamicDistributionGroups.Count) * 100)
    }

    $DynamicDistributionGroupMembers = Get-DynamicDistributionGroupMember -ResultSize Unlimited -Identity $DynamicDistributionGroup.Name | Measure-Object

    # Add Disytribution Group data to the array
    $MGUsersArray = $MGUsersArray + [PSCustomObject]@{
        DisplayName       = $DynamicDistributionGroup.DisplayName ;
        UserPrincipalName = "" ;
        Mail              = $DynamicDistributionGroup.WindowsEmailAddress ;
        Type              = "Dynamic Distribution List" ;
        Mailbox           = "" ;
        Archive           = "" ;
        OneDrive          = "" ;
        MailboxSizeMB     = "" ;
        ArchiveSizeMB     = "" ;
        OneDriveSize      = "" ;
        UsageLocation     = "" ;
        AccountEnabled    = "" ;
        IsDirSynced       = $DynamicDistributionGroup.IsDirSynced ;
        Members           = $DynamicDistributionGroupMembers.Count ;
        GuestMembers      = "" ;
        ID                = $DynamicDistributionGroup.ID;
        LicenseSKUs       = "" ;  
    }    
    $i++
}

$i = 1
Foreach ($MGContact in $MGContacts) { 
    If ($MGContacts.Count -gt "1") { 
        Write-Progress -ID 1 -Activity "Processing Contacts" -Status "Contact $i of $($MGContacts.Count)" -PercentComplete (($i / $MGContacts.Count) * 100)  
    }
    
    # Add contact data to the array
    $MGUsersArray = $MGUsersArray + [PSCustomObject]@{
        DisplayName       = $MGContact.DisplayName ;
        UserPrincipalName = "" ;
        Mail              = $MGContact.Mail ;
        Type              = "Contact";
        Mailbox           = "" ;
        Archive           = "" ;
        OneDrive          = "" ;
        MailboxSizeMB     = "" ;
        ArchiveSizeMB     = "" ;
        OneDriveSize      = "" ;
        UsageLocation     = "" ;
        AccountEnabled    = "" ;
        IsDirSynced       = $MGContact.OnPremisesSyncEnabled ;
        Members           = "" ;
        GuestMembers      = "" ;
        ID                = $MGContact.ID;
        LicenseSKUs       = "" ;     
    }
    $i++
}

$SecurityGroups = $mggroups | Where-Object { $_.Mailenabled -eq $False -and $_.securityEnabled -eq $True }
$i = 1
Foreach ($SecurityGroup in $SecurityGroups) { 
    If ($SecurityGroups.Count -gt "1") { 
        Write-Progress -ID 1 -Activity "Processing Security Groups" -Status "SecurityGroup $i of $($SecurityGroups.Count)" -PercentComplete (($i / $SecurityGroups.Count) * 100)  
    }

    If ($SecurityGroup.MembershipRuleProcessingState -ne "On") {
        $Type = "Security Group"
    }
    If ($SecurityGroup.MembershipRuleProcessingState -eq "On") {
        $Type = "Dynamic Security Group"
    }
    #>  
    # Add contact data to the array
    $MGUsersArray = $MGUsersArray + [PSCustomObject]@{
        DisplayName       = $SecurityGroup.DisplayName ;
        UserPrincipalName = "" ;
        Mail              = "" ;
        Type              = $Type ;
        Mailbox           = "" ;
        Archive           = "" ;
        OneDrive          = "" ;
        MailboxSizeMB     = "" ;
        ArchiveSizeMB     = "" ;
        OneDriveSize      = "" ;
        UsageLocation     = "" ;
        AccountEnabled    = "" ;
        IsDirSynced       = $SecurityGroup.OnPremisesSyncEnabled ;
        Members           = "" ;
        GuestMembers      = "" ;
        ID                = $SecurityGroup.ID;
        LicenseSKUs       = "" ;     
    }
    $i++
}
# Mark progress as completed

# Wait for 5 seconds
Start-Sleep 5


# Export  data to Excel if array is not empty
If ($MGUsersArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $MGUsersArray | Export-Excel -Path $Output -AutoSize -TableName User_Overview -WorksheetName User_Overview
    Start-Sleep 5
    Show-Report  
}

Stop-Transcript