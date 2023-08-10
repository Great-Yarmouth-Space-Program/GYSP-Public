$AcceptedDomains = Get-AcceptedDomain | Where-Object { $_.DomainName -notlike "*onmicrosoft.com" }
$RecipientTotalsArray = @()

foreach ($AcceptedDomain in $AcceptedDomains) {
    $CustomDomain = $AcceptedDomain.Name
    $Recipients = Get-Recipient -ResultSize Unlimited | Where-Object { $_.EmailAddresses -match $CustomDomain }
    
    $UserMailboxArray = $Recipients | Where-Object { $_.RecipientType -eq "UserMailbox" }
    $MailUserArray = $Recipients | Where-Object { $_.RecipientType -eq "MailUser" }
    $MailContactArray = $Recipients | Where-Object { $_.RecipientType -eq "MailContact" }
    $MailUniversalDistributionGroupArray = $Recipients | Where-Object { $_.RecipientType -eq "MailUniversalDistributionGroup" }
    $MailUniversalSecurityGroupArray = $Recipients | Where-Object { $_.RecipientType -eq "MailUniversalSecurityGroup" }
    
    $RecipientsCount =$Recipients.Count
    $UserMailboxCount = $UserMailboxArray.Count
    $MailUserCount = $MailUserArray.Count
    $MailContactCount = $MailContactArray.Count
    $MailUniversalDistributionGroupCount = $MailUniversalDistributionGroupArray.Count
    $MailUniversalSecurityGroupCount = $MailUniversalSecurityGroupArray.Count

    $RecipientTotalsArray += [PSCustomObject]@{
        CustomDomain = $CustomDomain
        UserMailbox = $UserMailboxCount
        MailUser = $MailUserCount
        MailContact = $MailContactCount
        MailUniversalDistributionGroup = $MailUniversalDistributionGroupCount
        MailUniversalSecurityGroup = $MailUniversalSecurityGroupCount
        Total = $RecipientsCount
    }
}

$RecipientTotalsArray | Export-Csv .\UBRecipients.csv -NoTypeInformation