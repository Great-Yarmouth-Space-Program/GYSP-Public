<#
    Name:               GYSP-SetDenyAddAndCustomizePages.ps1
    Version:            1.0
    Date:               11-08-2023
    Original Author:    Si Ansell
    Email:              graph@greatyarmouthspaceprogram.space

    Modules:            

    Use:                Gets a total of recipients for each custom accepted domain
                        Exports to CSV by default

    Updates:        
#>






Connect-SPOservice 
$URL = "https://.sharepoint.com/sites/sitename"
Get-SPOSite -Identity $URL | select DenyAddAndCustomizePages
Set-SPOSite -Identity $URL -DenyAddAndCustomizePages 0
