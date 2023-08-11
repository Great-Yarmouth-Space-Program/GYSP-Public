<#
    Name:               GYSP-CustomDomainsRecipientCounts.PS1
    Version:            1.0
    Date:               11-08-2023
    Original Author:    Si Ansell
    Email:              graph@greatyarmouthspaceprogram.space

    Use:                Gets a total of recipients for each custom accepted domain

    Updates:        
#>

# Get accepted domains excluding the default "onmicrosoft.com" domain
$AcceptedDomains = Get-AcceptedDomain | Where-Object { $_.DomainName -notlike "*onmicrosoft.com" }

# Initialize arrays for different recipient types and total recipient counts
$UserMailboxArray = @()
$MailUserArray = @()
$MailContactArray = @()
$MailUniversalDistributionGroupArray = @()
$MailUniversalSecurityGroupArray = @()
$RecipientTotalsArray = @()

# Iterate through each accepted domain
foreach ($AcceptedDomain in $AcceptedDomains) {
    $CustomDomain = $AcceptedDomain.Name
    
    # Retrieve recipients matching the current custom domain
    $Recipients = Get-Recipient -ResultSize Unlimited | Where-Object { $_.EmailAddresses -match $CustomDomain }
    
    # Separate recipients into different arrays based on their recipient types
    $UserMailboxArray = $Recipients | Where-Object { $_.RecipientType -eq "UserMailbox" }
    $MailUserArray = $Recipients | Where-Object { $_.RecipientType -eq "MailUser" }
    $MailContactArray = $Recipients | Where-Object { $_.RecipientType -eq "MailContact" }
    $MailUniversalDistributionGroupArray = $Recipients | Where-Object { $_.RecipientType -eq "MailUniversalDistributionGroup" }
    $MailUniversalSecurityGroupArray = $Recipients | Where-Object { $_.RecipientType -eq "MailUniversalSecurityGroup" }
    
    # Calculate total counts for each recipient type and the overall total
    $RecipientsTotal = $Recipients | Measure-Object
    $UserMailboxArrayTotal = $UserMailboxArray | Measure-Object
    $MailUserArrayTotal = $MailUserArray | Measure-Object
    $MailContactArrayTotal = $MailContactArray | Measure-Object
    $MailUniversalDistributionGroupArrayTotal = $MailUniversalDistributionGroupArray | Measure-Object
    $MailUniversalSecurityGroupArrayTotal = $MailUniversalSecurityGroupArray | Measure-Object

    # Create a custom object for the current domain with recipient counts
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

# Export recipient counts to a CSV file
$RecipientTotalsArray | Export-Csv .\Recipients.csv -NoTypeInformation
