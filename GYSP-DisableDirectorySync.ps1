<#
    Name:               GYSP-DisableDirectorySync.PS1
    Version:            NOT TESTED
    Date:               11-08-2023
    Original Author:    Si Ansell
    Email:              graph@greatyarmouthspaceprogram.space

    Use:                Disables AADC directory Sync
                        ### WARNING - Directory Sync CANNOT be reestablished to the tenant for 48-72 hours ###

    Updates:        
#>

# Connect to Microsoft Graph API with required scopes
Connect-MgGraph -scopes Organization.ReadWrite.All

# Get the ID of the organization
$OrgID = (Get-MgOrganization).id

# Create a parameter hash table to update on-premises sync setting
$params = @{
    onPremisesSyncEnabled = $null  # Set to $null to disable on-premises sync
}

# Update the organization using the Beta version of the Microsoft Graph API
Update-MgBetaOrganization -OrganizationId $OrgID -BodyParameter $params
