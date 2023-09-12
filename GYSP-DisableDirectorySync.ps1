<#
DISCLAIMER:

    By running this PowerShell script, you acknowledge and accept that:
    
    - You execute this script at your own risk and discretion.
    - The author of this script is not responsible for any damage, data loss, or unintended consequences 
      that may result from executing this script.
    - It is recommended to test this script in a controlled environment before use in a production scenario.

    If you do not agree with these terms, please refrain from executing this script.


    Name:               GYSP-DisableDirectorySync.PS1
    Version:            NOT TESTED
    Date:               11-08-2023
    Original Author:    Si Ansell
    Email:              graph@greatyarmouthspaceprogram.space

    Use:                Disables AADC directory Sync
                        ### WARNING - Directory Sync CANNOT be reestablished to the tenant for 48-72 hours ###

    Updates:           12-09-2023 Disclaimer added - SRA
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
