$AcceptedDomains = Get-AcceptedDomain | Where-Object { $_.DomainName -notlike "*onmicrosoft.com" }

$UserMailboxArray = @()
$MailUserArray = @()
$MailContactArray = @()
$MailUniversalDistributionGroupArray = @()
$MailUniversalSecurityGroupArray = @()
$RecipientTotalsArray = @()

foreach ($AcceptedDomain in $AcceptedDomains) {
    $CustomDomain = $AcceptedDomain.Name
    $Recipients = Get-Recipient -ResultSize Unlimited | Where-Object { $_.EmailAddresses -match $CustomDomain }
    
    $UserMailboxArray = $Recipients | Where-Object { $_.RecipientType -eq "UserMailbox" }
    $MailUserArray = $Recipients | Where-Object { $_.RecipientType -eq "MailUser" }
    $MailContactArray = $Recipients | Where-Object { $_.RecipientType -eq "MailContact" }
    $MailUniversalDistributionGroupArray = $Recipients | Where-Object { $_.RecipientType -eq "MailUniversalDistributionGroup" }
    $MailUniversalSecurityGroupArray = $Recipients | Where-Object { $_.RecipientType -eq "MailUniversalSecurityGroup" }
    
    $RecipientsTotal =$Recipients | Measure
    $UserMailboxArrayTotal = $UserMailboxArray | Measure
    $MailUserArrayTotal = $MailUserArray | Measure
    $MailContsctArrayTotal = $MailContsctArray | Measure
    $MailUniversalDistributionGroupArrayTotal = $MailUniversalDistributionGroupArray | Measure
    $MailUniversalSecurityGroupArrayTotal = $MailUniversalSecurityGroupArray | Measure

    $RecipientTotalsArray += [PSCustomObject]@{
        CustomDomain = $CustomDomain
            UserMailbox = $UserMailboxArrayTotal.Count
            MailUser = $MailUserArrayTotal.Count
            MailContact = $MailContactArrayTotal.Count
            MailUniversalDistributionGroup = $MailUniversalDistributionGroupArrayTotal.Count
            MailUniversalSecurityGroup = $MailUniversalSecurityGroupArrayTotal.Count
            Total = $RecipientsTotal.Count
    }
}

$RecipientTotalsArray | Export-Csv .\Recipients.csv -NoTypeInformation