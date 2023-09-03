<#
    Softcat Tenant 


  .SYNOPSIS
    Creates a Tenant Snapshot Report to assist with scoping a Microsoft 365 Tenant to Tenant Migration 

  .DESCRIPTION
  

  .PARAMETER 

        tenantId - The ID of the tenant to connect to

  .PARAMETER 

        ClientID - The ID of the client app

  .PARAMETER 
        
        Secret - The client app secret
   
  .INPUTS
  None. You cannot pipe objects to this script

  .OUTPUTS

  .EXAMPLE


    Modules:
            Install-Module -Name ExchangeOnlineManagement -noclobber
            Install-Module -Name ImportExcel
            Install-Module -Name Microsoft.Online.SharePoint.PowerShell
            Install-Module -Name MicrosoftTeams -allowclobber
            Install-Module Microsoft.Graph -Scope AllUsers
            Install-Module -Name Microsoft.Graph.Intune

 
#>

# Parameters
<#
Param(
    [parameter(Mandatory = $true)]
    $tenantId,
    [parameter(Mandatory = $true)]
    $ClientID,
    [parameter(Mandatory = $true)]
    $Secret
)
#>

#cd 

$ScriptVersion = "MG1.0.2"
$TemplateVersion = "MG1.0.2"

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
#endregion

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
Connect-MGGraph -Scopes User.Read.All,Group.Read.All,OrgContact.Read.All,Device.Read.All,Policy.Read.All,Application.Read.All
Connect-MSGraph
Connect-MicrosoftTeams
#endregion

##### Script Start #####

#region Script

#region Manage Template and Paths
$Org = Get-MgOrganization
$Date = Get-Date -Format 'yyyyMMdd_HHmmss'
$Output = "Softcat_Tenant_Assessment-" + $Date + ".xlsx"
$Template = "Softcat_Tenant_Assessment-Template-vMG1.0.2.xlsx"
Write-Host "Copying Template to $Output"
Copy-Item $Template $Output
$Transcript = "Softcat_Tenant_Assessment-" + $Date + "-Transcript.Log"
Start-Transcript -Path $Transcript
#endregion

#region Process Bars
$p=1
# Update this to the number of process so the count is correct
$TP = 34
# Sets the start time for the elapsed time counter
$StartTime = $(get-date)
#endregion

#region Users

#region All Accounts
$Process = "All Accounts"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
$MGUsers = @()
$MGUsersArray = @()
$MGUsers = Get-MGUser -All | Sort-Object displayname
$i = 1
Foreach($MGUser in $MGUsers) {
    If($MGUsers.Count -gt "1"){
        Write-Progress  -ID 1 -Activity "Processing All MG User Accounts" -Status "User Account $i of $($MGUsers.Count)" -PercentComplete (($i / $MGUsers.Count) * 100) 
    }
        $MGUserProperties = Get-MgUser -UserID $MGUser.ID -Property id,userprincipalname,displayname,mail,licenseAssignmentStates,usagelocation,usertype,accountenabled,OnPremisesSyncEnabled | select id,userprincipalname,displayname,mail,licenseAssignmentStates,usagelocation,usertype,accountenabled,OnPremisesSyncEnabled 
        $LicenseSKUs = Get-MgUserLicenseDetail -UserId $MGUser.ID 
        $MGUsersArray  = $MGUsersArray  + [PSCustomObject]@{
            ID = $MGUser.ID;
            DisplayName = $MGUser.DisplayName ;
            UserPrincipalName = $MGUser.UserPrincipalName ;
            Mail = $MGUserProperties.Mail ;
            UserType = $MGUserProperties.UserType
            UsageLocation = $MGUserProperties.UsageLocation ;
            AccountEnabled = $MGUserProperties.AccountEnabled ;
            LicenseSKUs = $LicenseSKUs.SkuPartNumber -join ";" ;
            IsDirSynced = $MGUserProperties.OnPremisesSyncEnabled
    }
$i++
}
Start-Sleep 5
If($MGUsersArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"    
    $MGUsersArray | Export-Excel -Path $Output -AutoSize -TableName Accounts_All -WorksheetName Accounts_All 
}
Write-Progress  -ID 1 -Activity "Processing All MG User Accounts" -Completed
$P++

#endregion

#region Contacts

$Process = "Contacts"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
$MGContactsArray = @()
$MGContacts = Get-MGContact -All | Sort-Object Displayname
$i = 1
Foreach ($MGContact in $MGContacts) { 
    If($MGContacts.Count -gt "1") { 
        Write-Progress  -ID 1 -Activity "Processing Contacts" -Status "MGContact $i of $($MGContacts.Count)" -PercentComplete (($i / $MGContacts.Count) * 100)  
    }
    $MGContactsArray = $MGContactsArray + [PSCustomObject]@{
        DisplayName = $MGContact.DisplayName ; 
        #RecipientTypeDetails = $MGContact.RecipientTypeDetails ;
        Company = $MGContact.Company ; 
        FirstName = $MGContact.GivenName ; 
        LastName = $MGContact.SurName ; 
        Email = $MGContact.Mail ; 
        IsDirSynced = $MGContact.OnPremisesSyncEnabled ; 
    }
$i++
}
Start-Sleep 5
If($MGContactsArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"    
    $MGContactsArray | Export-Excel -Path $Output -AutoSize -TableName Accounts_Contacts -WorksheetName Accounts_Contacts
}
Write-Progress  -ID 1 -Activity "Processing Contacts" -Completed
$P++

#endregion

#region Guests

$Process = "Guest Accounts"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)

$MGGuestAccounts = $MGUsersArray | Where{$_.UserType -eq "Guest"}
$MGGuestAccountsArray = @()
$i = 1
Foreach($MGGuestAccount in $MGGuestAccounts) {
    If($MGGuestAccounts.Count -gt "1"){
        Write-Progress  -ID 1 -Activity "Processing Guest Accounts" -Status "Guest Account $i of $($MGGuestAccounts.Count)" -PercentComplete (($i / $MGGuestAccounts.Count) * 100) 
    }
    $MGGuestAccountsArray = $MGGuestAccountsArray + [PSCustomObject]@{
            ID = $MGGuestAccount.ID;
            DisplayName = $MGGuestAccount.DisplayName ;
            UserPrincipalName = $MGGuestAccount.UserPrincipalName ;
            Mail = $MGGuestAccount.Mail ;
            UserType = $MGGuestAccount.UserType
            UsageLocation = $MGGuestAccount.UsageLocation ;
            AccountEnabled = $MGGuestAccount.AccountEnabled ;
            IsDirSynced = $MGGuestAccount.OnPremisesSyncEnabled
    }
    Start-Sleep 1
    $i++
}
Start-Sleep 5
If($MGGuestAccountsArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $MGGuestAccountsArray | Export-Excel -Path $Output -AutoSize -TableName Accounts_Guest -WorksheetName Accounts_Guest  
}
Write-Progress  -ID 1 -Activity "Processing Guest Accounts" -Completed
$P++

#endregion

#region Licensed Accounts
$Process = "Licensed Accounts"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
$MGLicensedAccounts = $MGUsersArray | Where{$_.LicenseSKUs -ne ""}
$MGLicensedAccountsArray = @()
$i = 1
Foreach($MGLicensedAccount in $MGLicensedAccounts) {
    If($MGLicensedAccounts.Count -gt "1"){
        Write-Progress  -ID 1 -Activity "Processing Licensed Accounts" -Status "Licensed Account $i of $($MGLicensedAccounts.Count)" -PercentComplete (($i / $MGLicensedAccounts.Count) * 100) 
    }
    $MGLicensedAccountsArray = $MGLicensedAccountsArray + [PSCustomObject]@{
            ID = $MGLicensedAccount.ID;
            DisplayName = $MGLicensedAccount.DisplayName ;
            UserPrincipalName = $MGLicensedAccount.UserPrincipalName ;
            Mail = $MGLicensedAccount.Mail ;
            UserType = $MGLicensedAccount.UserType
            UsageLocation = $MGLicensedAccount.UsageLocation ;
            AccountEnabled = $MGLicensedAccount.AccountEnabled ;
            LicenseSKUs = $MGLicensedAccount.LicenseSKUs -join ";" ;
            IsDirSynced = $MGLicensedAccount.OnPremisesSyncEnabled
    }
    $1++
}
#Start-Sleep 5
If($MGLicensedAccountsArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $MGLicensedAccountsArray | Export-Excel -Path $Output -AutoSize -TableName Accounts_Licensed -WorksheetName Accounts_Licensed  
}
Write-Progress  -ID 1 -Activity "Processing Licensed Accounts" -Completed
$P++

#endregion

#region UnLicensed Accounts
$Process = "UnLicensed Accounts"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
$MGUnLicensedAccounts = $MGUsersArray | Where{$_.LicenseSKUs -eq ""}
$MGUnLicensedAccountsArray = @()
$i = 1
Foreach($MGUnLicensedAccount in $MGUnLicensedAccounts) {
    If($MGLicensedAccounts.Count -gt "1"){
        Write-Progress  -ID 1 -Activity "Processing UnLicensed Accounts" -Status "UnLicensed Account $i of $($MGUnLicensedAccounts.Count)" -PercentComplete (($i / $MGUnLicensedAccounts.Count) * 100) 
    }
    $MGUnLicensedAccountsArray = $MGUnLicensedAccountsArray + [PSCustomObject]@{
            ID = $MGUnLicensedAccount.ID;
            DisplayName = $MGUnLicensedAccount.DisplayName ;
            UserPrincipalName = $MGUnLicensedAccount.UserPrincipalName ;
            Mail = $MGUnLicensedAccount.Mail ;
            UserType = $MGUnLicensedAccount.UserType
            UsageLocation = $MGUnLicensedAccount.UsageLocation ;
            AccountEnabled = $MGUnLicensedAccount.AccountEnabled ;
            LicenseSKUs = $MGUnLicensedAccount.LicenseSKUs -join ";" ;
            IsDirSynced = $MGUnLicensedAccount.OnPremisesSyncEnabled
    }
    $1++
}
#Start-Sleep 5
If($MGUnLicensedAccountsArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $MGUnLicensedAccountsArray | Export-Excel -Path $Output -AutoSize -TableName Accounts_UnLicensed -WorksheetName Accounts_UnLicensed  
}
Write-Progress  -ID 1 -Activity "Processing UnLicensed Accounts" -Completed
$P++

#endregion

#region Account SKUs

$Process = "License SKUs"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)

$MGAccountSkus = Get-MgSubscribedSku | Sort-Object SkuPartNumber 
$MGSKUArray = @()
$i = 1
Foreach ($MGAccountSku in $MGAccountSkus){
    If($MGAccountSkus.Count -gt "1"){
        Write-Progress  -ID 1 -Activity "Processing License SKUs" -Status "SKU $i of $($MGAccountSkus.Count)" -PercentComplete (($i / $MGAccountSkus.Count) * 100) 
    }
        $SKU = (get-MgSubscribedSku -SubscribedSkuId $MGAccountSku.id).SkuPartNumber
        $LicenseCount = (get-MgSubscribedSku -SubscribedSkuId $MGAccountSku.id -Property PrepaidUnits | select-object -expandproperty prepaidunits).enabled
        $ConsumedLicenses = (get-MgSubscribedSku -SubscribedSkuId $MGAccountSku.id).ConsumedUnits
        $MGSKUArray = $MGSKUArray + [PSCustomObject]@{
            SKU = $SKU;
            Purchased = $LicenseCount ;
            Consumed = $ConsumedLicenses ; 
            Available = $LicenseCount - $ConsumedLicenses
    }
#    Start-Sleep 1
    $i++
   }
If($MGSKUArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $MGSKUArray | Export-Excel -Path $Output -AutoSize -TableName Accounts_SKU -WorksheetName Accounts_SKU  
}
Write-Progress  -ID 1 -Activity "Processing License SKUs" -Completed
$P++

#endregion

#region Devices

$Process = "Devices"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
# Gets Device data
$MGDevices = Get-MGDevice -All  | Sort-Object Displayname
$MGDeviceArray = @()
$i = 1
Foreach($MGDevice in $MGDevices) {
    If($MGdevices.count -gt "1") {
        Write-Progress  -ID 1 -Activity "Processing Devices" -Status "Device $i of $($MGDevices.Count)" -PercentComplete (($i / $MGDevices.Count) * 100)  
    }
    $MGDeviceArray = $MGDeviceArray + [PSCustomObject]@{
        DisplayName = $MGDevice.DisplayName ;         
        DeviceOsType = $MGDevice.OperatingSystem ;
        DeviceOsVersion = $MGDevice.OperatingSystemVersion ; 
        DeviceTrustType = $MGDevice.TrustType ; 
        ApproximateLastLogonTimestamp = $MGDevice.ApproximateLastSignInDateTime ; 
    }
$i++
}
Start-Sleep 5
If($MGDeviceArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $MGDeviceArray | Export-Excel -Path $Output -AutoSize -TableName Devices_All -WorksheetName Devices_All
}
Start-Sleep 5
$MGDevicesUniqueArray = $MGDeviceArray | Sort-Object -Property DisplayName -Unique
If($MGDevicesUniqueArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $MGDevicesUniqueArray | Export-Excel -Path $Output -AutoSize -TableName Devices_Unique -WorksheetName Devices_Unique
}
Write-Progress  -ID 1 -Activity "Processing Devices" -Completed
$P++

#endregion

#region Domains

$Process = "Domains"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
$AllMGDomains = Get-MGDomain | Sort-Object Name
$AllMGDomainsArray = @()
$i = 1
Foreach($MGDomain in $AllMGDomains) {
    If($AllMGDomains.Count -gt "1") {
        Write-Progress -ID 1 -Activity "Processing MG Domains" -Status "Domain $i of $($AllMGDomains.Count)" -PercentComplete (($i / $AllMGDomains.Count) * 100)  
    }
    $AllMGDomainsArray  = $AllMGDomainsArray  + [PSCustomObject]@{
        Id = $MGDomain.Id ;
        IsDefault = $MGDomain.IsDefault ;
        IsInitial = $MGDomain.IsInitial ;
        AuthenticationType = $MGDomain.AuthenticationType ;
        IsVerified = $MGDomain.IsVerified ;
        SupportedServices = $MGDomain.SupportedServices -join ";" ; 
    }
#Start-Sleep 1
$i++
}
Start-Sleep 5
If($AllMGDomainsArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $AllMGDomainsArray | Export-Excel -Path $Output -AutoSize -TableName Domains_All -WorksheetName Domains_All  
}
Write-Progress -ID 1 -Activity "Processing MG Domains" -Completed
$P++


#endregion

#region MX Records

$Process = "MX Records"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
# Gets Accepted Domains and resolves MX records
$AcceptedDomains = Get-AcceptedDomain | Sort-Object Name
$AllMXRecordsArray = @()
$i = 1
Foreach($AcceptedDomain in $AcceptedDomains) {
    If($AcceptedDomains.Count -gt "1") {
        Write-Progress -ID 1 -Activity "Processing MX Records" -Status "Accepted Domain $i of $($AcceptedDomains.Count)" -PercentComplete (($i / $AcceptedDomains.Count) * 100)  
    }
    $MXRecords = $AcceptedDomain | resolve-dnsname -Type MX -Server 8.8.8.8 | Where-Object {$_.QueryType -eq "MX"}  | Select-Object Name,NameExchange,Preference,TTL | Sort-Object Preference
    Foreach($MXrecord in $MXRecords) {        
        $AllMXRecordsArray = $AllMXRecordsArray  + [PSCustomObject]@{
            Name = $MXRecord.Name ;
            NameExchange = $MXRecord.NameExchange ;
            Preference = $MXRecord.Preference ;
            TTL = $MXRecord.TTL ;
        }
    }
Start-Sleep 1
$i++
}
Start-Sleep 5
If($AllMXREcordsArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $AllMXRecordsArray | Export-Excel -Path $Output -AutoSize -TableName Domains_MXRecords -WorksheetName Domains_MXRecords 
} 
Write-Progress -ID 1 -Activity "Processing MX Records" -Completed
$P++

#endregion

#region SPF Records

$Process = "SPF Records"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
$AcceptedDomains = Get-AcceptedDomain
$AllSPFRecordsArray = @()
$i = 1
Foreach($AcceptedDomain in $AcceptedDomains) {
    If($AcceptedDomains.Count -gt "1") {
        Write-Progress -ID 1 -Activity "Processing SPF Records" -Status "Accepted Domain $i of $($AcceptedDomains.Count)" -PercentComplete (($i / $AcceptedDomains.Count) * 100)  
    }
$SPFRecords = $AcceptedDomain | resolve-dnsname -Type TXT -Server 8.8.8.8
    Foreach($SPFRecord in $SPFRecords) {  
        If($SPFRecord.Strings -like "V=*"){     
            $AllSPFRecordsArray = $AllSPFRecordsArray  + [PSCustomObject]@{
                Name = $SPFRecord.Name ;
                String = $SPFRecord.Strings -join "," ;
                TTL = $SPFRecord.TTL ;
            }
        }
    }
Start-Sleep 1
}
Start-Sleep 5
If($AllSPFRecordsArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $AllSPFRecordsArray | Export-Excel -Path $Output -AutoSize -TableName Domains_SPFRecords -WorksheetName Domains_SPFRecords 
} 
Write-Progress -ID 1 -Activity "Processing SPF Records" -Completed
$P++

#endregion

#region EOL Inbound Connectors

$Process = "Exchange InBound Connectors"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
# Gets Inbound Connectors
$InboundConnectors = Get-InboundConnector | Sort-Object Name
$InboundConnectorsArray = @()
$i = 1
Foreach($InboundConnector in $InboundConnectors) {
    If($InboundConnectors.Count -gt "1") {
        Write-Progress  -ID 1 -Activity "Processing Inbound Connectors" -Status "Connector $i of $($InboundConnectorsArray.Count)" -PercentComplete (($i / $InboundConnectorsArray.Count) * 100) 
    }
    $InboundConnectorsArray = $InboundConnectorsArray + [PSCustomObject]@{
        Name = $InboundConnector.Name; 
        Enabled = $InboundConnector.Enabled ;
        ConnectorType =$InboundConnector.ConnectorType ;
        SenderIPAddresses = $InboundConnector.SenderIPAddresses  -Join ","  ;
        SenderDomains = $InboundConnector.SenderDomains  -Join "," ;
    }
Start-Sleep 1
$1++
}
Start-Sleep 5
If($InboundConnectorsArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $InboundConnectorsArray | Export-Excel -Path $Output -AutoSize -TableName EOL_InboundConnectors -WorksheetName EOL_InboundConnectors
}
Write-Progress  -ID 1 -Activity "Processing Inbound Connectors" -Completed
$P++

#endregion

#region EOL Outbound Connectors

$Process = "Exchange Outbound Connectors"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
# Gets Outbound Connectors. Does not get TestModeConnectors
$OutboundConnectors = Get-OutboundConnector | Sort-Object Name
$OutboundConnectorsArray = @()
$i = 1
Foreach($OutboundConnector in $OutboundConnectors) {
    If($OutboundConnectors.Count -gt "1") {
        Write-Progress  -ID 1 -Activity "Processing Outbound Connectors" -Status "Connector $i of $($OutboundConnectorsArray.Count)" -PercentComplete (($i / $OutboundConnectorsArray.Count) * 100) 
    }
    $OutboundConnectorsArray = $OutboundConnectorsArray + [PSCustomObject]@{
        Name = $OutboundConnector.Name; 
        Enabled = $OutboundConnector.Enabled ;
        ConnectorType =$OutboundConnector.ConnectorType ;
        UseMXRecord = $OutboundConnector.UseMXRecord ;
        IsValidated = $OutboundConnector.IsValidated ;
        TlsSettings = $OutboundConnector.TlsSettings ; 
        SmartHosts = $OutboundConnector.SmartHosts  -Join "," ; 
        RecipientDomains = $OutboundConnector.RecipientDomains -Join "," ;
    }
Start-Sleep 1
$1++
}
Start-Sleep 5
If($OutboundConnectorsArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $OutboundConnectorsArray | Export-Excel -Path $Output -AutoSize -TableName EOL_OutboundConnectors -WorksheetName EOL_OutboundConnectors #  -Append
}
Write-Progress  -ID 1 -Activity "Processing Outbound Connectors" -Completed
$P++

#endregion

#region Mail Flow Rules

$Process = "Mail Flow Rules"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
# Gets Transport Rules
$TransportRules = Get-TransportRule | Sort-Object Priority
$TransportRulesArray =@()
$i = 1
ForEach($TransportRule in $TransportRules) {
    If($TransportRules.Count -gt "1") {
        Write-Progress  -ID 1 -Activity "Processing Transport Rules" -Status "Transport Rule $i of $($TransportRules.Count)" -PercentComplete (($i / $TransportRules.Count) * 100)  
    }
    $TransportRulesArray = $TransportRulesArray + [PSCustomObject]@{
        Name = $TransportRule.Name ;
        State = $TransportRule.State ; 
        Mode = $TransportRule.Mode ;
        Priority = $TransportRule.Priority ; 
        Comments = $TransportRule.Comments ;
    }
Start-Sleep 1
$i++
}
Start-Sleep 5
If($TransportRulesArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $TransportRulesArray | Export-Excel -Path $Output -AutoSize -TableName EOL_TransportRules -WorksheetName EOL_TransportRules
}
Write-Progress  -ID 1 -Activity "Processing Transport Rules" -Completed
$P++

#endregion

#region Distribution Groups

$Process = "Distribution Groups"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
# Gets Distribution Groups data 
$DistributionGroups = Get-DistributionGroup -ResultSize Unlimited | Sort-Object DisplayName
$DistributionGroupsArray = @()
$i = 1
Foreach($DistributionGroup in $DistributionGroups) {
    If($DistributionGroups.Count -gt "1"){
        Write-Progress  -ID 1 -Activity "Processing Distribution Groups" -Status "Distribution Group $i of $($DistributionGroups.Count)" -PercentComplete (($i / $DistributionGroups.Count) * 100) 
    }
    $DistributionGroupMembers = Get-DistributionGroupMember  -ResultSize Unlimited -Identity $DistributionGroup.Name
    $DistributionGroupsArray = $DistributionGroupsArray + [PSCustomObject]@{
        Name = $DistributionGroup.Name ;
        DisplayName = $DistributionGroup.DisplayName ;
        GroupType = $DistributionGroup.GroupType ;
        RecipientTypeDetails = $DistributionGroup.RecipientTypeDetails; 
        PrimarySmtpAddress = $DistributionGroup.PrimarySmtpAddress; 
        Members = $DistributionGroupMembers.Count ;
    }
$i++
}
Write-Progress  -ID 1 -Activity "Processing Distribution Groups" -Completed
Start-Sleep 5
If($DistributionGroupsArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $DistributionGroupsArray | Export-Excel -Path $Output -AutoSize -TableName Groups_Distribution -WorksheetName Groups_Distribution  
}
Start-Sleep 5
$P++

#endregion

#region Dynamic Distribution Groups

$Process = "Dynamic Distribution Groups"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
# Gets Dynamic Distribution Groups data 
$DynamicDistributionGroups = Get-DynamicDistributionGroup -ResultSize Unlimited | Sort-Object DisplayName
$DynamicDistributionGroupsArray = @()
$i = 1
Foreach($DynamicDistributionGroup in $DynamicDistributionGroups) {
    If($DynamicDistributionGroups.Count -gt "1"){
        Write-Progress  -ID 1 -Activity "Processing Dynamic Distribution Groups" -Status "Dynamic Distribution Group $i of $($DynamicDistributionGroups.Count)" -PercentComplete (($i / $DynamicDistributionGroups.Count) * 100) 
    }
    $DynamicDistributionGroupsArray = $DynamicDistributionGroupsArray + [PSCustomObject]@{
        Name = $DynamicDistributionGroup.Name ;
        DisplayName = $DynamicDistributionGroup.DisplayName ;
        RecipientFilterType = $DynamicDistributionGroup.RecipientFilterType ;
        RecipientTypeDetails = $DynamicDistributionGroup.RecipientTypeDetails; 
        PrimarySmtpAddress = $DynamicDistributionGroup.PrimarySmtpAddress;
        ManagedBy = $DynamicDistributionGroup.ManagedBy ;
    }
Start-Sleep 1
$i++
}
Start-Sleep 5
If($DynamicDistributionGroupsArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $DynamicDistributionGroupsArray | Export-Excel -Path $Output -AutoSize -TableName Groups_DynamicDistribution -WorksheetName Groups_DynamicDistribution  
}
Write-Progress  -ID 1 -Activity "Processing Dynamic Distribution Groups" -Completed
$P++    

#endregion

#region Microsoft 365 Groups

$Process = "Microsoft 365 Groups"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
# Gets Microsoft 365 Groups data
$M365Groups = Get-UnifiedGroup -ResultSize Unlimited | Sort-Object DisplayName
$M365GroupOnlyArray = @()
$TeamsOnlyArray = @()
Foreach($M365Group in $M365Groups){
    If($M365Group.ResourceProvisioningOptions -eq "Team"){
        $TeamsOnlyArray += $M365Group 
    }
    Else{
        $M365GroupOnlyArray += $M365Group
    }
}
$M365GroupArray = @()
$i = 1
Foreach($M365GroupOnly in $M365GroupOnlyArray) {
    If($M365GroupOnlyArray.Count -gt "1"){
        Write-Progress  -ID 1 -Activity "Processing M365 Groups" -Status "M365 Group $i of $($M365GroupOnlyArray.Count)" -PercentComplete (($i / $M365GroupOnlyArray.Count) * 100)  
    }
    
    $SharePointSite = Get-SPOSite $M365GroupOnly.SharePointSiteUrl
    $GroupOwners = Get-UnifiedGroupLinks -Identity $M365GroupOnly.Name -LinkType Owners
    $GroupOwnersCounts = $GroupOwners | Measure-Object
    $GroupMembers = Get-UnifiedGroupLinks -Identity $M365GroupOnly.Name -LinkType Members
    $GroupMembersCount = $GroupMembers | Measure-Object
    $TotalUsers = [int]$GroupOwnerCount.Count + [int]$GroupMemberCount.Count 
    $M365GroupArray = $M365GroupArray + [PSCustomObject]@{
        Name = $M365GroupOnly.Name ;
        DisplayName = $M365GroupOnly.DisplayName ;
        AccessType = $M365GroupOnly.AccessType ;
        PrimarySMTPAddress = $M365GroupOnly.PrimarySMTPAddress ;
        EmailAddressess = $M365GroupOnly.EmailAddressess -join ',' ;
        SharePointSiteURL = $M365GroupOnly.SharePointSiteUrl ;
        SharePointDocumentsUrl = $M365GroupOnly.SharePointDocumentsUrl ;
        SharePointNotebookUrl = $M365GroupOnly.SharePointNotebookUrl ; 
        StorageMB = $SharePointSite.StorageUsageCurrent ;
        OwnerCounts = $GroupOwnerCount.Count ; 
        MemberCount = $GroupMemberCount.count ;
        UserCount = $TotalUsers ; 
        GroupExternalMemberCount = $M365GroupOnly.GroupExternalMemberCount
        ManagedBy = $M365GroupOnly.ManagedBy -join ',' ;
        Members = $GroupMember -join ',' ;

    }
$i++
}  
Start-Sleep 5
If($M365GroupArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $M365GroupArray | Export-Excel -Path $Output -AutoSize -TableName Groups_M365 -WorksheetName Groups_M365  
}
Write-Progress  -ID 1 -Activity "Processing M365 Groups" -Completed
$P++

#endregion

#region Guest Accounts in Microsoft 365 Groups



#endregion

#region Mailboxes

$Process = "Mailboxes"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
$AllMailboxes = Get-Mailbox -ResultSize Unlimited | Where-Object{$_.PrimarySMTPAddress -notLike "DiscoverySearchMailbox*"} |Sort-Object UserPrincipalname
$AllMailboxesArray = @()
$i = 1
Foreach($Mailbox in $AllMailboxes) {
    If($AllMailboxes.Count -gt "1"){
        Write-Progress  -ID 1 -Activity "Processing Mailbox Data" -Status "Mailbox $i of $($AllMailboxes.Count)" -PercentComplete (($i / $AllMailboxes.Count) * 100)  
    }
    $MailboxStats = Get-MailboxStatistics -Identity $Mailbox.UserPrincipalName | Select-Object LastLogonTime, DisplayName, @{Name="TotalItemSizeMB"; Expression={[math]::Round(($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB),0)}}, ItemCount, DeletedItemCount, @{Name="TotalDeletedItemSizeMB"; Expression={[math]::Round(($_.TotalDeletedItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB),0)}}
    If($Mailbox.ArchiveStatus -eq "Active"){
            $ArchiveStats =  Get-MailboxStatistics -Identity $Mailbox.UserPrincipalName -Archive | Select-Object DisplayName, @{Name="TotalItemSizeMB"; Expression={[math]::Round(($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB),0)}}, ItemCount, DeletedItemCount, @{Name="TotalDeletedItemSizeMB"; Expression={[math]::Round(($_.TotalDeletedItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB),0)}}
        }
    Else{
            $ArchiveStats = ""
        }
    $AllMailboxesArray  = $AllMailboxesArray  + [PSCustomObject]@{
        UserPrincipalName = $Mailbox.UserPrincipalName ;
        DisplayName = $Mailbox.DisplayName ;
        Alias = $Mailbox.Identity ;
        PrimarySmtpAddress = $Mailbox.PrimarySmtpAddress ;
        RecipientTypeDetails = $Mailbox.RecipientTypeDetails ;
        ItemCount = $MailboxStats.ItemCount; 
        TotalItemSizeMB = $MailboxStats.TotalItemSizeMB; 
        DeletedItemCount = $MailboxStats.DeletedItemCount;  
        TotalDeletedItemSizeMB = $MailboxStats.TotalDeletedItemSizeMB;
        Archive = $Mailbox.ArchiveStatus ;
        ArchiveDisplayName = $ArchiveStats.DisplayName
        ArchiveItemCount = $ArchiveStats.ItemCount ; 
        ArchiveTotalItemSizeMB = $ArchiveStats.TotalItemSizeMB ; 
        ArchiveDeletedItemCount = $ArchiveStats.DeletedItemCount ; 
        ArchiveTotalDeletedItemSizeMB = $ArchiveStats.TotalDeletedItemSizeMB ; 
        WhenCreatedUTC = $Mailbox.WhenCreatedUTC ;
        WhenChangedUTC = $Mailbox.WhenChangedUTC ;
        LastLogonTime = $MailboxStats.LastLogonTime ;
        EmailAddresses = $Mailbox.EmailAddresses -join ',';
        LitigationHoldEnabled = $Mailbox.LitigationHoldEnabled  ;
        LitigationHoldDuration = $Mailbox.LitigationHoldDuration ;
        InPlaceHolds = $Mailbox.InPlaceHolds -join ',' ;
        RetentionPolicy = $Mailbox.RetentionPolicy ;
        RetentionHoldEnabled = $Mailbox.RetentionHoldEnabled ;
        StartDateForRetentionHold = $Mailbox.StartDateForRetentionHold ; 
        EndDateForRetentionHold = $Mailbox.EndDateForRetentionHold ;
        AccountSKU = $licenseString -join ',' ;
        Guid = $Mailbox.Guid ;
    }
$i++
}
Write-Progress -ID 1 -Activity "Gathering Mailbox Data" -Completed
$P++

If($AllMailboxesArray -ne $Null) {
    $AllMailboxesArray | Export-Excel -Path $Output -AutoSize -TableName Mailbox_All -WorksheetName Mailbox_All 
} 
Start-Sleep 5
$UserMailboxes = $AllMailboxesArray | Where-Object{$_.RecipientTypeDetails -eq "UserMailBox"}
If($UserMailboxes -ne $Null) {
    $UserMailboxes | Export-Excel -Path $Output -AutoSize -TableName Mailbox_User -WorksheetName Mailbox_User  
}
Start-Sleep 5
$NONUserMailboxes = $AllMailboxesArray | Where-Object{$_.RecipientTypeDetails -ne "UserMailBox"} 
If($NONUserMailboxes -ne $Null) {
    $NONUserMailboxes | Export-Excel -Path $Output -AutoSize -TableName Mailbox_Other -WorksheetName Mailbox_Other          
}
Start-Sleep 5
$SharedMailboxes = $AllMailboxesArray | Where-Object{$_.RecipientTypeDetails -eq "SharedMailBox"}
If($SharedMailboxes -ne $Null) {
    $SharedMailboxes | Export-Excel -Path $Output -AutoSize -TableName Mailbox_Shared -WorksheetName Mailbox_Shared          
}
Start-Sleep 5
$RoomMailboxes = $AllMailboxesArray | Where-Object{$_.RecipientTypeDetails -eq "RoomMailBox"}
If($RoomMailboxes -ne $Null) {
    $RoomMailboxes | Export-Excel -Path $Output -AutoSize -TableName Mailbox_Room -WorksheetName Mailbox_Room          
}
Start-Sleep 5
$EquipmentMailboxes = $AllMailboxesArray | Where-Object{$_.RecipientTypeDetails -eq "EquipmentMailBox"}
If($EquipmentMailboxes -ne $Null) {
    $EquipmentMailboxes | Export-Excel -Path $Output -AutoSize -TableName Mailbox_Equipment -WorksheetName Mailbox_Equipment  
}
Start-Sleep 5
$OnHoldMailboxes = $AllMailboxesArray | Where-Object{$_.LitigationHoldEnabled -eq $TRUE}
If($OnHoldMailboxes -ne $Null) {
    $OnHoldMailboxes | Export-Excel -Path $Output -AutoSize -TableName Mailbox_OnHold -WorksheetName Mailbox_OnHold  
}
Start-Sleep 5
$Inactivemailboxes = $UserMailboxes | Where-Object{$_.Lastlogontime -lt (Get-Date).AddDays(-90)}
If($Inactivemailboxes -ne $Null) {
    $Inactivemailboxes | Export-Excel -Path $Output -AutoSize -TableName Mailbox_Inactive -WorksheetName Mailbox_Inactive  
}
Start-Sleep 5
$LargetsMailboxes = $AllMailboxesArray | Sort-Object TotalItemSizeMB -Descending | Select-Object -First 10
If($LargetsMailboxes -ne $Null) {
    $LargetsMailboxes | Export-Excel -Path $Output -AutoSize -TableName Mailbox_TopTen -WorksheetName Mailbox_TopTen  
}
Write-Host "Writing $Process data to $Output"
$P++

#endregion

#region Archive Mailboxes

$Process = "Archive Mailboxes"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
# Gets Archive Mailbox data
$AllArchiveMailboxes = Get-Mailbox -ResultSize Unlimited -Archive | Where-Object{$_.PrimarySMTPAddress -notLike "DiscoverySearchMailbox*"} |Sort-Object UserPrincipalname
$AllArchiveMailboxesArray = @()
$i = 1
Foreach($ArchiveMailbox in $AllArchiveMailboxes) {
    If($AllArchiveMailboxes.Count -gt "1"){
        Write-Progress  -ID 1 -Activity "Processing Archive Mailboxes" -Status "Archive Mailbox $i of $($AllArchiveMailboxes.Count)" -PercentComplete (($i / $AllArchiveMailboxes.Count) * 100)  
    }
    $ArchiveStats =  Get-MailboxStatistics -Identity $ArchiveMailbox.UserPrincipalName -Archive | Select-Object DisplayName, @{Name="TotalItemSizeMB"; Expression={[math]::Round(($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB),0)}}, ItemCount, DeletedItemCount, @{Name="TotalDeletedItemSizeMB"; Expression={[math]::Round(($_.TotalDeletedItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB),0)}}
    $AllArchiveMailboxesArray  = $AllArchiveMailboxesArray  + [PSCustomObject]@{
        UserPrincipalName = $ArchiveMailbox.UserPrincipalName ;
        DisplayName = $ArchiveMailbox.DisplayName ;
        Alias = $ArchiveMailbox.Identity ;
        PrimarySmtpAddress = $ArchiveMailbox.PrimarySmtpAddress ;
        RecipientTypeDetails = $ArchiveMailbox.RecipientTypeDetails ;
        Archive = $ArchiveMailbox.ArchiveStatus ;
        ArchiveItemCount = $ArchiveStats.ItemCount ; 
        ArchiveTotalItemSizeMB = $ArchiveStats.TotalItemSizeMB ; 
        ArchiveDeletedItemCount = $ArchiveStats.DeletedItemCount ; 
        ArchiveTotalDeletedItemSizeMB = $ArchiveStats.TotalDeletedItemSizeMB ; 
    }
$i++
}
Start-Sleep 5
If($AllArchiveMailboxesArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $AllArchiveMailboxesArray | Export-Excel -Path $Output -AutoSize -TableName Mailbox_Archive -WorksheetName Mailbox_Archive  
}
Write-Progress  -ID 1 -Activity "Processing Archive Mailboxes" -Completed
$P++

#endregion

#region Group Mailboxes

$Process = "Group Mailboxes"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
# Gets all M365 Group Mailbox data
$M365Groups = Get-UnifiedGroup -ResultSize Unlimited | Sort-Object DisplayName
$M365GroupMailboxesArray = @()
$i = 1
foreach($M365Group in $M365Groups) {
    If($M365Groups.Count -gt "1"){
        Write-Progress  -ID 1  -Activity "Processing Group Mailboxes" -Status "Group Mailbox $i of $($M365Groups.Count)" -PercentComplete (($i / $M365Groups.Count) * 100)
    }
    $M365MailboxStats = Get-MailboxStatistics -Identity $M365Group.PrimarySMTPAddress | Select-Object LastLogonTime, DisplayName, @{Name="TotalItemSizeMB"; Expression={[math]::Round(($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB),0)}}, ItemCount, DeletedItemCount, @{Name="TotalDeletedItemSizeMB"; Expression={[math]::Round(($_.TotalDeletedItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB),0)}}
    $M365GroupMailboxesArray = $M365GroupMailboxesArray  + [PSCustomObject]@{
        DisplayName = $M365Group.DisplayName ;
        Alias = $M365Group.Identity ;
        PrimarySmtpAddress = $M365Group.PrimarySmtpAddress ;
        RecipientTypeDetails = $M365Group.RecipientTypeDetails ;
        ItemCount = $M365MailboxStats.ItemCount; 
        TotalItemSizeMB = $M365MailboxStats.TotalItemSizeMB; 
        DeletedItemCount = $M365MailboxStats.DeletedItemCount;  
        TotalDeletedItemSizeMB = $M365MailboxStats.TotalDeletedItemSizeMB;
        WhenCreatedUTC = $M365Group.WhenCreatedUTC ;
        WhenChangedUTC = $M365Group.WhenChangedUTC ;
        LastLogonTime = $M365MailboxStats.LastLogonTime ;
        EmailAddresses = $M365Group.EmailAddresses -join ',';
        Guid = $M365Group.Guid ;
    }
$i++
}
Start-Sleep 5
If($M365GroupMailboxesArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $M365GroupMailboxesArray | Export-Excel -Path $Output -AutoSize -TableName Mailbox_M365Group -WorksheetName Mailbox_M365Group  
}
Write-Progress  -ID 1  -Activity "Processing Group Mailboxes" -Completed    
$P++

#endregion

#region Public Folders

$Process = "Public Folders"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
# Gets Public Folder data
$PublicFolders = Get-PublicFolder -Recurse -ResultSize Unlimited | Sort-Object Parentpath
$PublicFolderArray = @()
$i = 1
Foreach($PublicFolder in $PublicFolders) {
    If($PublicFolders.Count -gt "1") {
        Write-Progress  -ID 1 -Activity "Processing Public Folders" -Status "Public Folder $i of $($PublicFolders.Count)" -PercentComplete (($i / $PublicFolders.Count) * 100)  
    }
    $PublicFolderStats = Get-PublicFolderStatistics -Identity $PublicFolder.Identity | Select-Object Name, @{Name="TotalItemSizeMB"; Expression={[math]::Round(($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB),0)}}, ItemCount, DeletedItemCount, @{Name="TotalDeletedItemSizeMB"; Expression={[math]::Round(($_.TotalDeletedItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB),0)}}
    $PublicFolderArray = $PublicFolderArray + [PSCustomObject]@{
        Name = $PublicFolder.Name ;
        Identity = $PublicFolder.Identity ;
        ParentPath = $PublicFolder.ParentPath ;
        ItemCount = $PublicFolderStats.ItemCount; 
        TotalItemSizeMB = $PublicFolderStats.TotalItemSizeMB; 
        FolderClass = $PublicFolder.FolderClass ;
        MailEnabled = $PublicFolder.MailEnabled ; 
    }
$i++
}
Start-Sleep 5
If($PublicFolderArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $PublicFolderArray | Export-Excel -Path $Output -AutoSize -TableName PublicFolders_All -WorksheetName PublicFolders_All
}
Write-Progress  -ID 1 -Activity "Processing Public Folders" -Completed
$P++

#endregion

#region Recipient Addresses

$Process = "Recipient Addresses"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
# Gets ALL Recipient Addresses
$AcceptedDomains = Get-AcceptedDomain
$RecipientSMTPAddressArray = @()
$RecipientPrimaryAddressArray = @()
$RecipientProxyAddressArray = @()
$RecipientUserMailboxAddressArray = @()
$RecipientMailUniversalDistributionGroupAddressArray = @()
$RecipientMailUniversalSecurityGroupAddressArray = @()
$MailboxSMTPAddressArray = @()
$i = 1
Foreach($AcceptedDomain in $AcceptedDomains) {
    $Recipients = @()       
    $CustomDomain = $AcceptedDomain.Name
    $Descriptor = "*@" + $CustomDomain
    $Recipients = Get-Recipient -ResultSize Unlimited| Where-Object{$_.EmailAddresses -match $CustomDomain} 
    Foreach($Recipient in $Recipients) {
        If($Recipients.Count -gt "1"){
            Write-Progress  -ID 1 -Activity "Processing Accepted Domain Recipient Addresses" -Status "Recipient Address $i of $($AllMailboxes.Count)" -PercentComplete (($i / $AllMailboxes.Count) * 100)  
        }
        $RecipientAddresses = $Recipient.EmailAddresses | Where-Object { $_ -match '^smtp:'}
        Foreach($RecipientAddress in $RecipientAddresses) {
            $Type = $RecipientAddress.SubString(0,4)
            $Address = $RecipientAddress.substring(5) 
            If($RecipientAddress -Like $Descriptor) {
                $RecipientSMTPAddressArray = $RecipientSMTPAddressArray + [PSCustomObject]@{
                    WindowsLiveID = $Recipient.WindowsLiveID
                    Identity = $Recipient.Identity
                    Type = $Type
                    SMTPAddress = $Address
                    RecipientType = $Recipient.RecipientType
                }
            }
        }
    }
}
Start-Sleep 5
If($RecipientSMTPAddressArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $RecipientSMTPAddressArray | Sort WindowsliveID | Export-Excel -Path $Output -AutoSize -TableName Recipients_All -WorksheetName Recipients_All  
}
Write-Progress  -ID 1 -Activity "Processing Accepted Domain Recipient Addresses" -Completed
$P++  

#endregion

#region OneDrives

$Process = "OneDrives"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
# Gets all OneDrive data
$OneDrives = Get-SPOSite -IncludePersonalSite $true -Limit all -Filter "Url -like '-my.sharepoint.com/personal/'" | Sort-Object Title
$ActiveOneDriveArray = @()
$InactiveOneDriveArray = @()
$i = 1
Foreach($OneDrive in $OneDrives) {
    If($OneDrives.Count -gt "1"){
        Write-Progress  -ID 1 -Activity "Processing OneDrives" -Status "OneDrive $i of $($OneDrives.Count)" -PercentComplete (($i / $OneDrives.Count) * 100)  
    }
    $UserLicense = Get-MgUserLicenseDetail -UserId $OneDrive.Owner
    If($UserLicense -ne $Null) {        
        $ActiveOneDriveArray = $ActiveOneDriveArray + [PSCustomObject]@{
            Title = $OneDrive.Title ; 
            Owner = $OneDrive.Owner ;
            Status = $OneDrive.Status ; 
            LastContentModifiedDate = $OneDrive.LastContentModifiedDate ; 
            StorageUsageCurrentMB = $OneDrive.StorageUsageCurrent ; 
            Url = $OneDrive.Url ;
            SharingCapability = $OneDrive.SharingCapability ; 
            SiteDefinedSharingCapability = $OneDrive.SiteDefinedSharingCapability ; 
            ConditionalAccessPolicy = $OneDrive.ConditionalAccessPolicy ; 
        }
    }
    ElseIf($UserLicense -eq $Null) {
        $InactiveOneDriveArray = $InactiveOneDriveArray + [PSCustomObject]@{
            Title = $OneDrive.Title ; 
            Owner = $OneDrive.Owner ; 
            Status = $OneDrive.Status ;
            LastContentModifiedDate = $OneDrive.LastContentModifiedDate ; 
            StorageUsageCurrentMB = $OneDrive.StorageUsageCurrent ; 
            Url = $OneDrive.Url ; 
        }
    }
$i++
}
Start-Sleep 5
If($ActiveOneDriveArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $ActiveOneDriveArray | Export-Excel -Path $Output -AutoSize -TableName OneDrives -WorksheetName OneDrives          
}
Start-Sleep 5
If($InactiveOneDriveArray -ne $Null) {
    $InactiveOneDriveArray | Export-Excel -Path $Output -AutoSize -TableName OneDrives_Inactive -WorksheetName OneDrives_Inactive
}
Start-Sleep 5
$LargestOneDrives = $ActiveOneDriveArray | Sort-Object StorageUsageCurrentMB -Descending | Select-Object -First 10
Start-Sleep 5
If($LargestOneDrives -ne $Null) {
    $LargestOneDrives | Export-Excel -Path $Output -AutoSize -TableName OneDrives_TopTen -WorksheetName OneDrives_TopTen
}
Write-Progress  -ID 1 -Activity "Processing OneDrives" -Completed
$P++

#endregion

#region SharePoint Sites

$Process = "SharePoint Sites"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
# Gets SharePoint site data
$SPOSites = Get-SPOSite -IncludePersonalSite $False -Limit All | Sort-Object Title 
$SPOSitesArray = @()
$i = 1
ForEach ($SPOSite in $SPOSites) {
    If($SPOSites.Count -gt "1"){
        Write-Progress  -ID 1 -Activity "Processing SharePoint Sites" -Status "Site $i of $($SPOSites.Count)" -PercentComplete (($i / $SPOSites.Count) * 100) 
    }
    $M365Group =""
# Checks to see if the site is associated with a Microsoft 365 Group. Throws a lot of red
    $M365Group = Get-UnifiedGroup -Identity $SPOSite.Title
        If($M365Group -ne $Null) {
            $IsM365Group = "Group"
        }
        Else{
            $IsM365Group = "No"
        }
    $IsTeam = $M365Group | Select-Object ResourceProvisioningOptions -ExpandProperty ResourceProvisioningOptions
    $SPOSitesArray  = $SPOSitesArray  + [PSCustomObject]@{
        Title = $SPOSite.Title ; 
        LocaleId = $SPOSite.LocaleId ; 
        Url = $SPOSite.Url ; 
        Status = $SPOSite.Status ;
        M365Group = $IsM365Group ; 
        MicrosoftTeam = $IsTeam ; 
        LastContentModifiedDate = $SPOSite.LastContentModifiedDate ; 
        Owner = $SPOSite.Owner ; 
        Template = $SPOSite.Template ; 
        ResourceUsageCurrent = $SPOSite.ResourceUsageCurrent ; 
        ResourceUsageAverage = $SPOSite.ResourceUsageAverage ; 
        StorageUsageCurrentMB = $SPOSite.StorageUsageCurrent ; 
        ConditionalAccessPolicy = $SPOSite.ConditionalAccessPolicy ;
        SensitivityLabel = $SPOSite.SensitivityLabel ;
        AllowSelfServiceUpgrade = $SPOSite.AllowSelfServiceUpgrade ;
        AllowEditing = $SPOSite.AllowEditing ; 
        SharingAllowedDomainList = $SPOSite.SharingAllowedDomainList ; 
        SharingBlockedDomainList = $SPOSite.SharingBlockedDomainList ; 
        DenyAddAndCustomizePages = $SPOSite.DenyAddAndCustomizePages ;
        BlockDownloadLinksFileType = $SPOSite.BlockDownloadLinksFileType ;
        DefaultLinkPermission = $SPOSite.DefaultLinkPermission ;
        DefaultSharingLinkType = $SPOSite.DefaultSharingLinkType ; 
        DisableAppViews = $SPOSite.DisableAppViews ; 
        DisableCompanyWideSharingLinks = $SPOSite.DisableCompanyWideSharingLinks ; 
        DisableFlows = $SPOSite.DisableFlows ; 
        LimitedAccessFileType = $SPOSite.LimitedAccessFileType ; 
        LockState = $SPOSite.LockState ; 
        SandboxedCodeActivationCapability = $SPOSite.SandboxedCodeActivationCapability ; 
        SharingCapability = $SPOSite.SharingCapability ; 
        ShowPeoplePickerSuggestionsForGuestUsers = $SPOSite.ShowPeoplePickerSuggestionsForGuestUsers ; 
        SharingDomainRestrictionMode = $SPOSite.SharingDomainRestrictionMode ; 
        LockIssue = $SPOSite.LockIssue ; 
        WebsCount = $SPOSite.WebsCount ; 
        CompatibilityLevel = $SPOSite.CompatibilityLevel ; 
        DisableSharingForNonOwnersStatus = $SPOSite.DisableSharingForNonOwnersStatus ; 
        HubSiteId = $SPOSite.HubSiteId ; 
        IsHubSite = $SPOSite.IsHubSite ; 
        RelatedGroupId = $SPOSite.RelatedGroupId ; 
        GroupId = $SPOSite.GroupId ; 
    }  
$i++
}
$SPOTemplates = $SPOSites.Template
$SPOTemplatesGroup = $SPOTemplates | Group-Object
Start-Sleep 5
If($SPOSitesArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    #All Sites
    $SPOSitesArray | Export-Excel -Path $Output -AutoSize -TableName SPOSites_All -WorksheetName SPOSites_All  
    Start-Sleep 5
    #Unique SPO Sites
    $SPOSitesArray | Where-Object{$_.M365Group -ne "Group"-and $_.Template -NotLike "TEAMCHANNEL*" } | Export-Excel -Path $Output -AutoSize -TableName SPOSites_Unique -WorksheetName SPOSites_Unique
    Start-Sleep 5
    #Microsoft 365 Groups
    $SPOSitesArray | Where-Object{$_.M365Group -eq "Group" -and $_.MicrosoftTeam -Ne "Team"} | Export-Excel -Path $Output -AutoSize -TableName SPOSites_M365Group -WorksheetName SPOSites_M365Group  
    Start-Sleep 5
    #Microsoft Teams
    $SPOSitesArray | Where-Object{$_.MicrosoftTeam -eq "Team"} | Export-Excel -Path $Output -AutoSize -TableName SPOSites_Teams -WorksheetName SPOSites_Teams
    Start-Sleep 5
    #Largest Sites
    $LargestSites = $SPOSitesArray | Sort-Object StorageUsageCurrentMB -Descending | Select-Object -First 10
    $LargestSites | Export-Excel -Path $Output -AutoSize -TableName SPOSites_TopTen -WorksheetName SPOSites_TopTen
    Start-Sleep 5
    #Microsoft Teams Channels
    $SPOSitesArray | Where-Object{$_.M365Group -ne "Yes"-and $_.Template -Like "TEAMCHANNEL*" } | Export-Excel -Path $Output -AutoSize -TableName SPOSites_TeamsChannels -WorksheetName SPOSites_TeamsChannels
}
Write-Progress  -ID 1 -Activity "Processing SharePoint Sites" -Completed   
$P++

#endregion

#region Custom Sensitivity Labels

$Process = "Custom Sensitivity Labels"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
# Gets Custom Sensitivity Label data
$CustomSensitivityLabels = Get-DlpSensitiveInformationType | Where-Object{$_.Publisher -notlike "Microsoft Corporation"} | Sort-Object Name 
$CustomSensitivityLabelsArray = @()
$i = 1
Foreach($CustomSensitivityLabel in $CustomSensitivityLabels){
    If($CustomSensitivityLabels.Count -gt "1") {
        Write-Progress  -ID 1 -Activity "Processing Custom Sensitivity Labels" -Status "Custom Sensitivity Label $i of $($DLPPolicies.Count)" -PercentComplete (($i / $DLPPolicies.Count) * 100)  
    }
    $CustomSensitivityLabelsArray = $CustomSensitivityLabelsArray + [PSCustomObject]@{
        Name = $CustomSensitivityLabel.Name ; 
        Description = $CustomSensitivityLabel.Description ; 
        RecommendedConfidence = $CustomSensitivityLabel.RecommendedConfidence ; 
        Publisher = $CustomSensitivityLabel.Publisher ; 
        Type = $CustomSensitivityLabel.Type ; 
        RulePackId = $CustomSensitivityLabel.RulePackId ; 
    }    
Start-Sleep 1
$i++
}    
Start-Sleep 5
If($CustomSensitivityLabelsArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"    
    $CustomSensitivityLabelsArray | Export-Excel -Path $Output -AutoSize -TableName Policies_CustomLabels -WorksheetName Policies_CustomLabels
}
Write-Progress  -ID 1 -Activity "Processing Custom Sensitivity Labels" -Completed
$P++


#endregion

#region DLP Policies

$Process = "DLP Policies"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
# Gets DLP Policiy data
$DLPPolicies = Get-DlpCompliancePolicy 
$DLPPoliciesArray = @()
$i = 1
Foreach($DLPPolicy in $DLPPolicies) {
    If($DLPPolicies.Count -gt "1") {
        Write-Progress  -ID 1 -Activity "Processing DLP Policies" -Status "DLP Policy $i of $($DLPPolicies.Count)" -PercentComplete (($i / $DLPPolicies.Count) * 100)  
    }
    $DLPPoliciesArray = $DLPPoliciesArray + [PSCustomObject]@{
        Name = $DLPPolicy.Name ; 
        Mode = $DLPPolicy.Mode ; 
        Type = $DLPPolicy.Type ; 
        Workload = $DLPPolicy.Workload ; 
        Priority = $DLPPolicy.Priority ; 
        CreatedBy = $DLPPolicy.CreatedBy ; 
        Enabled = $DLPPolicy.Enabled ; 
        Comment = $DLPPolicy.Comment ; 
    }    
Start-Sleep 1
$i++
}    
Start-Sleep 5
If($DLPPoliciesArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"    
    $DLPPoliciesArray | Export-Excel -Path $Output -AutoSize -TableName Policies_DLP -WorksheetName Policies_DLP
}
Write-Progress  -ID 1 -Activity "Processing DLP Policies" -Completed
$P++

#endregion

#region Retention Policies

$Process = "Retention Policies"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
# Gets Retention Policy data
$RetentionPolicies = Get-RetentionCompliancePolicy | Sort-Object Name
$RetentionPoliciesArray = @()
$i = 1
Foreach($RetentionPolicy in $RetentionPolicies) {
    If($RetentionPolicies.Count -gt "1") {
        Write-Progress  -ID 1 -Activity "Processing Retention Policies" -Status "Retention Policy $i of $($RetentionPolicies.Count)" -PercentComplete (($i / $RetentionPolicies.Count) * 100)  
    }
    $RetentionPoliciesArray = $RetentionPoliciesArray + [PSCustomObject]@{
        Name =$retentionPolicy.Name ; 
        Mode =$retentionPolicy.Mode ; 
        Workload =$retentionPolicy.Workload ; 
        CreatedBy =$retentionPolicy.CreatedBy ; 
        Enabled =$retentionPolicy.Enabled ; 
        Comment =$retentionPolicy.Comment ;
    }
Start-Sleep 1
$i++
}    
Start-Sleep 5
If($RetentionPoliciesArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $RetentionPoliciesArray | Export-Excel -Path $Output -AutoSize -TableName Policies_Retention -WorksheetName Policies_Retention
}
Write-Progress  -ID 1 -Activity "Processing Retention Policies" -Completed
$P++

#endregion

#region Conditional Access Policies

$Process = "Conditional Access Policies"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
$MGConditionalAccessPolicies = Get-MgIdentityConditionalAccessPolicy | Sort-Object Name
$MGConditionalAccessPoliciesArray = @()
$i = 1
Foreach($ConditionalAccessPolicy in $MGConditionalAccessPolicies) {
    If($ConditionalAccessPolicies.Count -gt "1") {
        Write-Progress  -ID 1 -Activity "Processing Conditional Access Policies" -Status "Device Compliance Policy $i of $($MGConditionalAccessPolicies.Count)" -PercentComplete (($i / $MGConditionalAccessPolicies.Count) * 100)  
    }
    $MGConditionalAccessPoliciesArray = $MGConditionalAccessPoliciesArray + [PSCustomObject]@{
        DisplayName =$ConditionalAccessPolicy.DisplayName ; 
        CreatedDate =$ConditionalAccessPolicy.CreatedDateTime ; 
        State =$ConditionalAccessPolicy.State
        }
Start-Sleep 1
$i++
}    
Start-Sleep 5
If($MGConditionalAccessPoliciesArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $MGConditionalAccessPoliciesArray | Export-Excel -Path $Output -AutoSize -TableName Policies_ConditionalAccess -WorksheetName Policies_ConditionalAccess
}
Write-Progress  -ID 1 -Activity "Processing Conditional Access Policies" -Completed
$P++

#endregion

#region Intune Policies

Connect-MSGraph
$Process = "Intune Policies"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
# Gets Intune Policies
$IntuneAppProtectionPolicies = Get-IntuneAppProtectionPolicy | Sort-Object Displayname
$IntuneDeviceCompliancePolicies = Get-IntuneDeviceCompliancePolicy | Sort-Object Displayname
$IntuneDeviceConfigurationPolicies = Get-IntuneDeviceConfigurationPolicy | Sort-Object Displayname
$IntuneMdmWindowsInformationProtectionPolicies = Get-IntuneMdmWindowsInformationProtectionPolicy | Sort-Object Displayname
$IntuneMobileAppConfigurationPolicies = Get-IntuneMobileAppConfigurationPolicy | Sort-Object Displayname
$IntuneWindowsInformationProtectionPolicies = Get-IntuneWindowsInformationProtectionPolicy | Sort-Object Displayname
$PoliciesCount = $IntuneAppProtectionPolicies.Count + $IntuneDeviceCompliancePolicies.Count + $IntuneDeviceConfigurationPolicies.Count + $IntuneMdmWindowsInformationProtectionPolicies.Count + $IntuneMobileAppConfigurationPolicies.Count + $IntuneWindowsInformationProtectionPolicies.count
$IntunePoliciesArray = @()
$i = 1
Foreach($IntuneAppProtectionPolicy in $IntuneAppProtectionPolicies) {
    If($IntuneAppProtectionPolicies.Count -gt "1") {
        Write-Progress  -ID 1 -Activity "Processing Endpoint Policies" -Status "Policy $i of $($PoliciesCount)" -PercentComplete (($i / $PoliciesCount) * 100) 
    }
    $IntunePoliciesArray = $IntunePoliciesArray + [PSCustomObject]@{
        DisplayNameName = $IntuneAppProtectionPolicy.DisplayName; 
        Description = $IntuneAppProtectionPolicy.description ;
        Type = "App Protection Policy" ;
    }
Start-Sleep 1
$1++
}
Foreach($IntuneDeviceCompliancePolicy in $IntuneDeviceCompliancePolicies) {
    If($IntuneDeviceCompliancePolicies.Count -gt "1") {
        Write-Progress  -ID 1 -Activity "Processing Endpoint Policies" -Status "Policy $i of $($PoliciesCount)" -PercentComplete (($i / $PoliciesCount) * 100) 
    }
    $IntunePoliciesArray = $IntunePoliciesArray + [PSCustomObject]@{
        DisplayNameName = $IntuneDeviceCompliancePolicy.DisplayName; 
        Description = $IntuneDeviceCompliancePolicy.description ;
        Type = "Device Compliance Policy" ;
}
Start-Sleep 1
$1++
}
Foreach($IntuneDeviceConfigurationPolicy in $IntuneDeviceConfigurationPolicies){
    If($IntuneDeviceConfigurationPolicies.Count -gt "1") {
        Write-Progress  -ID 1 -Activity "Processing Endpoint Policies" -Status "Policy $i of $($PoliciesCount)" -PercentComplete (($i / $PoliciesCount) * 100) 
    }
    $IntunePoliciesArray = $IntunePoliciesArray + [PSCustomObject]@{
        DisplayNameName = $IntuneDeviceConfigurationPolicy.DisplayName; 
        Description = $IntuneDeviceConfigurationPolicy.description ;
        Type = "Intune Device Configuration Policy" ;
}
Start-Sleep 1
$1++
}
Foreach($IntuneMdmWindowsInformationProtectionPolicy in $IntuneMdmWindowsInformationProtectionPolicies){
    If($IntuneMdmWindowsInformationProtectionPolicies.Count -gt "1") {
        Write-Progress  -ID 1 -Activity "Processing Endpoint Policies" -Status "Policy $i of $($PoliciesCount)" -PercentComplete (($i / $PoliciesCount) * 100) 
    }
    $IntunePoliciesArray = $IntunePoliciesArray + [PSCustomObject]@{
        DisplayNameName = $IntuneMdmWindowsInformationProtectionPolicy.DisplayName; 
        Description = $IntuneMdmWindowsInformationProtectionPolicy.description ;
        Type = "Intune MDM Windows Information Protection Policy" ;
}
Start-Sleep 1
$1++
}
Foreach($IntuneMobileAppConfigurationPolicy in $IntuneMobileAppConfigurationPolicies) {
    If($IntuneMobileAppConfigurationPolicies.Count -gt "1") {
        Write-Progress  -ID 1 -Activity "Processing Endpoint Policies" -Status "Policy $i of $($PoliciesCount)" -PercentComplete (($i / $PoliciesCount) * 100) 
    }
    $IntunePoliciesArray = $IntunePoliciesArray + [PSCustomObject]@{
        DisplayNameName = $IntuneMobileAppConfigurationPolicy.DisplayName; 
        Description = $IntuneMobileAppConfigurationPolicy.description ;
        Type = "Intune Mobile App Configuration Policy" ;
}
Start-Sleep 1
$1++
}
Foreach($IntuneWindowsInformationProtectionPolicy in $IntuneWindowsInformationProtectionPolicies) {
    If($IntuneWindowsInformationProtectionPolicies.Count -gt "1"){
        Write-Progress  -ID 1 -Activity "Processing Endpoint Policies" -Status "Policy $i of $($PoliciesCount)" -PercentComplete (($i / $PoliciesCount) * 100) 
    }
    $IntunePoliciesArray = $IntunePoliciesArray + [PSCustomObject]@{
        DisplayNameName = $IntuneWindowsInformationProtectionPolicy.DisplayName; 
        Description = $IntuneWindowsInformationProtectionPolicy.description ;
        Type = "Intune Windows Information Protection" ;
}
Start-Sleep 1
$1++
}
Start-Sleep 5
If($IntunePoliciesArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $IntunePoliciesArray | Export-Excel -Path $Output -AutoSize -TableName Policies_EndPoint -WorksheetName Policies_EndPoint
}
Write-Progress  -ID 1 -Activity "Processing Endpoint Policies" -Completed
$P++

#endregion

#region Power Environments

$Process = "Power Environments"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)

$PowerEnvironments = Get-AdminPowerAppEnvironment | Sort DisplayName
$PowerEnvironmentsArray = @()
$i = 1
Foreach($PowerEnvironment in $PowerEnvironments) {
    If($PowerEnvironments.Count -gt "1"){
        Write-Progress  -ID 1 -Activity "Processing Power Apps" -Status "Environment $i of $($PowerEnvironments.Count)" -PercentComplete (($i / $PowerEnvironments.Count) * 100) 
    }
    $PowerEnvironmentsArray = $PowerEnvironmentsArray + [PSCustomObject]@{
            DisplayName = $PowerEnvironment.DisplayName ;
            IsDefault = $PowerEnvironment.IsDefault ;
            Location = $PowerEnvironment.IsDefault ;
            Created = $PowerEnvironment.CreatedTime ;
            CreatedBy = $PowerEnvironment.CreatedBy.DisplayName ;
            CreationType = $PowerEnvironment.CreationType ;
    }
    Start-Sleep 1
    $i++
}
Start-Sleep 5
If($PowerEnvironmentsArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $PowerEnvironmentsArray | Export-Excel -Path $Output -AutoSize -TableName Power_Environments -WorksheetName Power_Environments  
}
Write-Progress  -ID 1 -Activity "Processing Power Apps" -Completed
$P++

#endregion

#region Power Apps

$Process = "Power Apps"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)

$PowerApps = Get-AdminPowerApp | Sort DisplayName
$PowerAppsArray = @()
$i = 1
Foreach($PowerApp in $PowerApps) {
    If($PowerApps.Count -gt "1"){
        Write-Progress  -ID 1 -Activity "Processing Power Apps" -Status "Power App $i of $($PowerApps.Count)" -PercentComplete (($i / $PowerApps.Count) * 100) 
    }
    $PowerAppsArray = $PowerAppsArray + [PSCustomObject]@{
            DisplayName = $PowerApp.DisplayName ;
            AppType = $PowerApp.Internal.appType ;
            Created = $PowerApp.CreatedTime ;
            EnvironmentName = $PowerApp.EnvironmentName ;
            Owner = $Powerapp.owner.displayName ;
            OwnerEmail = $Powerapp.owner.email
    }
    Start-Sleep 1
    $i++
}
Start-Sleep 5
If($PowerAppsArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $PowerAppsArray | Export-Excel -Path $Output -AutoSize -TableName Power_Apps -WorksheetName Power_Apps  
}
Write-Progress  -ID 1 -Activity "Processing Power Apps" -Completed
$P++

#endregion

#region Power Flows

$Process = "Power Flows"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)

$PowerFlows = Get-AdminFlow | Sort DisplayName
$PowerFlowsArray = @()
$i = 1
Foreach($PowerFlow in $PowerFlows) {
    If($PowerFlows.Count -gt "1"){
        Write-Progress  -ID 1 -Activity "Processing Power Flows" -Status "Flow $i of $($PowerFlows.Count)" -PercentComplete (($i / $PowerFlows.Count) * 100) 
    }
    $PowerFlowsArray = $PowerFlowsArray + [PSCustomObject]@{
            DisplayName = $PowerFlow.DisplayName ;
            Enabled = $PowerFlow.Enabled ;
            UserType = $PowerFlow.UserType ;
            CreatedTime = $PowerFlow.CreatedTime ;
            CreatedBy = $PowerFlow.CreatedBy.UserID ;
            EnvironmentName = $PowerFlow.EnvironmentName
    }
    Start-Sleep 1
    $i++
}
Start-Sleep 5
If($PowerFlowsArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $PowerFlowsArray | Export-Excel -Path $Output -AutoSize -TableName Power_Flows -WorksheetName Power_Flows  
}
Write-Progress  -ID 1 -Activity "Processing Power Flows" -Completed
$P++

#endregion

#region Registered Applications

$Process = "Registered Applications"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)

$MGRegisteredApplications = Get-MgApplication
$MGRegisteredApplicationsArray = @()
$i = 1
Foreach($MGRegisteredApplication in $MGRegisteredApplications) {
    If($MGRegisteredApplications.Count -gt "1"){
        Write-Progress  -ID 1 -Activity "Processing Registered Applications" -Status "Environment $i of $($MGRegisteredApplications.Count)" -PercentComplete (($i / $MGRegisteredApplications.Count) * 100) 
    }
    $MGRegisteredApplicationsArray = $MGRegisteredApplicationsArray + [PSCustomObject]@{
            DisplayName = $MGRegisteredApplication.DisplayName ;
            AppId = $MGRegisteredApplication.AppId ;
            PublisherDomain = $MGRegisteredApplication.PublisherDomain ;
            CreatedDateTime = $MGRegisteredApplication.CreatedDateTime ;
    }
    Start-Sleep 1
    $i++
}
Start-Sleep 5
If($MGRegisteredApplicationsArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $MGRegisteredApplicationsArray | Export-Excel -Path $Output -AutoSize -TableName Registered_Applications -WorksheetName Registered_Applications  
}
Write-Progress  -ID 1 -Activity "Processing Registered Applications" -Completed
$P++

#endregion

#region Microsoft Teams

$Process = "Microsoft Teams"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
# Gets Microsoft Teams data    
$Teams = Get-Team | Sort-Object Displayname
$TeamArray = @()
$i = 1
Foreach($Team in $Teams) {  
    If($Teams.Count -gt "1") {
        Write-Progress  -ID 1 -Activity "Processing Microsoft Teams" -Status "Microsoft Team $i of $($Teams.Count)" -PercentComplete (($i / $Teams.Count) * 100) 
    }
    $TeamGroup = Get-UnifiedGroup -identity $Team.GroupId
    $SharePointSite = Get-SPOSite $TeamGroup.SharePointSiteUrl
    $Channels = Get-TeamChannel -GroupId $Team.GroupId
    $ChannelsCount = $Channels | Measure-Object
    $TeamOwner = (Get-TeamUser -GroupId $Team.GroupId | Where-Object{$_.Role -eq 'Owner'}).User
    $TeamOwnerCount = $TeamOwner | Measure-Object
    $TeamMember = (Get-TeamUser -GroupId $Team.GroupId | Where-Object{$_.Role -eq 'Member'}).User
    $TeamMemberCount = $TeamMember | Measure-Object
    $TeamGuest = (Get-TeamUser -GroupId $Team.GroupId | Where-Object{$_.Role -eq 'Guest'}).User
    $TeamGuestCount = $TeamGuest | Measure-Object
    $TotalUsers = [int]$TeamOwnerCount.Count + [int]$TeamMemberCount.Count + [int]$TeamGuestCount.Count
    $TeamArray = $TeamArray + [PSCustomObject]@{
        GroupId = $Team.GroupId ; 
        DisplayName = $Team.DisplayName ; 
        Description = $Team.Description ; 
        Visibility = $Team.Visibility ; 
        MailNickName = $Team.MailNickName ; 
        Classification = $Team.Classification ; 
        Archived = $Team.Archived ; 
        StorageMB = $SharePointSite.StorageUsageCurrent ;
        Channels = $ChannelsCount.Count ; 
        TeamOwners = $TeamOwnerCount.Count ; 
        TeamMembers = $TeamMemberCount.count ;
        TotalUsers = $TotalUsers ; 
        TeamGuests =  $TeamGuestCount.Count ;
    }
$i++
}
Start-Sleep 5
If($TeamArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"    
    $TeamArray | Export-Excel -Path $Output -AutoSize -TableName Teams_All -WorksheetName Teams_All  
}
Write-Progress  -ID 1 -Activity "Processing Microsoft Teams" -Completed
$P++

#endregion

#region Finishing Up

$Process = "Miscellaneous Data"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
# Adds data to the Dashboard or cells
$CompanyName = $Org.Displayname 
$ReportDate = (Get-Date).ToString()
If($Org.OnPremisesSyncEnabled -eq $Null) {
    $AADCStatus = "Disabled"
}
If($Org.OnPremisesSyncEnabled -ne $Null) {
    $AADCStatus = "Enabled"
}
$EndTime = $(get-date) - $StartTime
$TotalTime = "{0:HH:mm:ss}" -f ([datetime]$EndTime.Ticks)
$DefaultDomainName = $AcceptedDomains | Where{$_.Default -eq $True}
Start-Sleep 5
Write-Host "Writing $Process data to $Output"
$Excel = Open-ExcelPackage -Path $Output
$worksheet = $excel.Workbook.Worksheets['Data']
$worksheet.Cells['H4'].value = $ReportDate
$worksheet.Cells['H5'].value = $CompanyName
$worksheet.Cells['H6'].value = $Org.ID
$worksheet.Cells['H7'].value = $AdminURL
$worksheet.Cells['H8'].value = $Output
$worksheet.Cells['H9'].value = $TotalTime
$worksheet.Cells['H10'].value = $ScriptVersion
$worksheet.Cells['H11'].value = $TemplateVersion
$worksheet.Cells['C29'].value = $AADCStatus
$worksheet.Cells['C71'].value = $SPOTemplatesGroup.Count 
$worksheet.Cells['H12'].value = $DefaultDomainName.Name

Close-ExcelPackage $Excel
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -Completed
$P++

#####
$Process = "Cleaning Up"
$CurrentTime = $(get-date) - $StartTime
$ElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$CurrentTime.Ticks)
Write-Progress -Activity "Running Check $p of $TP - $Process - Elapsed Time - $ElapsedTime" -PercentComplete (($p / $TP) * 100)
Write-Host "Data written to $Output"
Write-Progress -ID 3 -Activity "Having a rest before opening the report in Excel"
Start-Sleep -Seconds 10  
Stop-Transcript
Show-Report

#endregion
#endregion

#endregion Script


Disconnect-MgGraph