<#

DISCLAIMER:

    By running this PowerShell script, you acknowledge and accept that:
    
    - You execute this script at your own risk and discretion.
    - The author of this script is not responsible for any damage, data loss, or unintended consequences 
      that may result from executing this script.
    - It is recommended to test this script in a controlled environment before use in a production scenario.

    If you do not agree with these terms, please refrain from executing this script.


    Name:               GYSP-SetUPNs_SMTP.PS1
    Version:            1.0
    Date:               11-08-2023
    Original Author:    Si Ansell
    Email:              graph@greatyarmouthspaceprogram.space

    Use:                Sets the UPN and SMTP of users in a CSV file
                        CSV should be in the format: OLDUPN,NEWUPN

    Modules:            ExchangeOnlineManagement
    
    Connections:        Connect-ExchangeOnline
    
    Updates:
#>

#region Functions

#region Check for Modules Function
function CheckForPowerShellModule([string]$ModuleName)
{
  Write-Host "Checking for $ModuleName module..."
  if (Get-Module -ListAvailable -Name $ModuleName)
  {
    # Module already installed 
    Write-Host "$ModuleName installed - Continuing" -ForegroundColor Green
  }
  else
  {
    # Module not installed
    Write-Host "$ModuleName not found. Please install module and rerun script" -ForegroundColor Red
    #Exit
  }
}

#endregion Functions

#region Open Report
Function Show-Report {
    Export-Excel $Output -show
}

#endregion


#endregion

#region Check for Modules
CheckForPowerShellModule("Microsoft.Graph")
CheckForPowerShellModule("ExchangeOnlineManagement")
CheckForPowerShellModule("ImportExcel")
CheckForPowerShellModule("Microsoft.Online.SharePoint.PowerShell")
CheckForPowerShellModule("MicrosoftTeams")
#CheckForPowerShellModule("MicrosoftPowerBIMgmt")

#endregion

#region Connect to Microsoft 365
Connect-ExchangeOnline
$OnMicrosoftDomain = Get-AcceptedDomain | Where-Object{$_.DomainName -like "*.onmicrosoft.com"} | Select-Object DomainName -ExpandProperty DomainName
$OnMicrosoftPrefix = $OnMicrosoftDomain.split('.')[0] 
$AdminURL = "https://" + $OnMicrosoftPrefix + "-admin.sharepoint.com"
Connect-SPOService -URL $AdminURL
Connect-IPPSSession
Connect-MGGraph -Scopes User.Read.All,Group.Read.All,OrgContact.Read.All,Device.Read.All,Policy.Read.All,Application.Read.All,SecurityEvents.Read.All,RoleManagement.Read.Directory,AuditLog.Read.All
Connect-MSGraph
Connect-MicrosoftTeams
#endregion


#region Authentication
#region L1
# Ensure modern authentication for Exchange Online is enabled
$OrganizationConfig = Get-OrganizationConfig #| Format-Table -Auto Name, OAuth*
If($OrganizationConfig.OAuth2ClientProfileEnabled -eq $True) {
    $ExchangeModernAuthentication = "PASS"
}
Else {
    $ExchangeModernAuthentication = "FAIL"
}

# Ensure modern authentication for Skype for Business Online is enabled
###$TeamsModernAuthentication = Get-CsOAuthConfiguration |fl ClientAdalAuthOverride

# Ensure modern authentication for SharePoint applications is required
$SPOTenant = Get-SPOTenant 
If($SPOTenant.LegacyAuthProtocolsEnabled -eq $True) {
    $SharePointModernAuthentication = "PASS"
}
Else {
    $SharePointModernAuthentication = "FAIL"
}

# Ensure that Office 365 Passwords Are Not Set to Expire
$MGDomainsArray = @()
$MGDomains = Get-MGDomain
Foreach($MGDomain in $MGDomains) {
    If($MGDomain.PasswordValidityPeriodInDays -eq "2147483647") {
        $DomainPasswordsNotSetToExpire = "PASS"
    }
    Else {
        $DomainPasswordsNotSetToExpire = "FAIL"
    }
    $MGDomainsArray = $MGDomainsArray  + [PSCustomObject]@{
        ID = $MGDomain.ID;
        PasswordNotificationWindowInDays = $MGDomain.PasswordNotificationWindowInDays ;
        PasswordValidityPeriodInDays = $MGDomain.PasswordValidityPeriodInDays
        PasswordsNotSetToExpire = $DomainPasswordsNotSetToExpire
    }
}

$PasswordsNotSetToExpireArray = $MGDomainsArray.PasswordsNotSetToExpire
If($PasswordsNotSetToExpireArray -contains "Fail" -and $PasswordsNotSetToExpireArray -contains "Pass") {
    $PasswordsNotSetToExpire = "Partial"
}
If($PasswordsNotSetToExpireArray -notcontains "Fail") {
    $PasswordsNotSetToExpire = "PASS"
}
If($PasswordsNotSetToExpireArray -notcontains "Pass") {
    $PasswordsNotSetToExpire = "FAIL"
}


#endregion L1
#endregion Authentication

#region Azure Active Directory
#region L1
#Ensure multifactor authentication is enabled for all users in administrative roles

###Get-MgSecuritySecureScore -top 1 | FL

# Ensure that between two and four global admins are designated
$GlobalAdmins = @()
$RoleId = (Get-MgDirectoryRole -Filter "DisplayName eq 'Global Administrator'").Id
$GlobalAdminArray = Get-MgDirectoryRoleMember -DirectoryRoleId $roleId
foreach ($GlobalAdmin in $GlobalAdminArray) {
    $UPN = (Get-MgUser -UserId $GlobalAdmin.id).UserPrincipalName
    $GlobalAdmins += $UPN
}
If($GlobalAdmins.Count -lt "5") {
    $GlobalAdminCount = "PASS"
}
Else{
    $GlobalAdminCount = "FAIL"
}

### Ensure self-service password reset is enabled

###Import-Module Microsoft.Graph.Reports

####Get-MgReportAuthenticationMethodUserRegistrationDetail

###  Ensure that password protection is enabled for Active Directory

# Enable Conditional Access policies to block legacy authentication

$MGConditionalAccessPolicies = Get-MgIdentityConditionalAccessPolicy | Sort-Object Name

###  Ensure that password hash sync is enabled for resiliency and leaked credential detection

#https://learn.microsoft.com/en-us/azure/active-directory/hybrid/cloud-sync/how-to-inbound-synch-ms-graph

### Ensure Security Defaults is disabled on Azure Active Directory

#(Invoke-RestMethod -Uri "$baseuri/policies/identitySecurityDefaultsEnforcementPolicy" -Headers $Header -Method get -ContentType "application/json")

#endregion L1



#endregion Azure Active Directory


#region Data Management
#region L1

# Ensure DLP policies are enabled

$DLPPolicies = Get-DlpCompliancePolicy 
$CustomSensitivityLabels = Get-DlpSensitiveInformationType | Where-Object{$_.Publisher -notlike "Microsoft Corporation"} | Sort-Object Name 




#endregion L1
#endregion Data Management



#region Export Results
$ExchangeModernAuthentication
$SharePointModernAuthentication
#$TeamsModernAuthentication
$PasswordsNotSetToExpire
#
$GlobalAdminCount



#endregion Export Results

