<#

DISCLAIMER:

    By running this PowerShell script, you acknowledge and accept that:
    
    - You execute this script at your own risk and discretion.
    - The author of this script is not responsible for any damage, data loss, or unintended consequences 
      that may result from executing this script.
    - It is recommended to test this script in a controlled environment before use in a production scenario.

    If you do not agree with these terms, please refrain from executing this script.


    Name:               GYSP-GetRecipientProxies.PS1
    Version:            1.0
    Date:               01-09-2023
    Original Author:    Si Ansell
    Email:              graph@greatyarmouthspaceprogram.space

    Use:                Gets ALL SMTP/smtp addressess

    Modules:            ExchangeOnlineManagement
    
    Connections:        Connect-ExchangeOnline
    
    Updates:
#>

# Gets ALL Recipient Addresses
$AcceptedDomains = Get-AcceptedDomain
$RecipientSMTPAddressArray = @()
$RecipientPrimaryAddressArray = @()
$RecipientProxyAddressArray = @()
$RecipientUserMailboxAddressArray = @()
$RecipientMailUniversalDistributionGroupAddressArray = @()
$RecipientMailUniversalSecurityGroupAddressArray = @()
$MailboxSMTPAddressArray = @()
Foreach($AcceptedDomain in $AcceptedDomains) {
    $Recipients = @()       
    $CustomDomain = $AcceptedDomain.Name
    $Descriptor = "*@" + $CustomDomain
    $Recipients = Get-Recipient -ResultSize Unlimited| Where-Object{$_.EmailAddresses -match $CustomDomain} 
    Foreach($Recipient in $Recipients) {
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
                    AcceptedDomain = $AcceptedDomain.Domainname ;
                    RecipientType = $Recipient.RecipientType;
                    RecipientTypeDetails = $Recipient.RecipientTypeDetails
                }
            }
        }
    }
}
Start-Sleep 5
If($RecipientSMTPAddressArray -ne $Null) {
    Write-Host "Writing $Process data to $Output"
    $RecipientSMTPAddressArray | Sort WindowsliveID | Export-Excel -Path .\RecipientProxiesMASTER.xlsx -AutoSize -TableName Recipient_Proxies -WorksheetName Recipient_Proxies  
}

#endregion