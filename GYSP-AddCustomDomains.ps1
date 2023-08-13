<#
    Name:               GYSP-AddCustomDomains.PS1
    Version:            1.0
    Date:               11-08-2023
    Original Author:    Si Ansell
    Email:              graph@greatyarmouthspaceprogram.space

    Use:                Create CustomDomains.csv with a single column headed id. 
                        Populate with custom domains
                        Run script
                        Verification DNS records will be exported to .\DomainVerification.csv

    Updates:        
#>

# Connect to Microsoft Graph API with required scopes
Connect-MGGraph -Scopes Domain.ReadWrite.All

# Initialize arrays for storing custom domains and DNS verification records
$CustomDomains = @()
$DNSVerification = @()

# Import custom domains from CSV file
$CustomDomains = Import-Csv CustomDomains.csv
$File = "DomainVerification.csv"

# Iterate through each custom domain
foreach ($CustomDomain in $CustomDomains) {
    $params = @{
        id = $CustomDomain.id
    }

    # Create a new domain using Microsoft Graph API
    New-MgDomain -BodyParameter $params

    # Retrieve and filter DNS verification TXT record
    $MgVerificationCode = (Get-MgDomainVerificationDnsRecord -DomainId $CustomDomain.Id | Where-Object { $_.RecordType -eq "Txt" }).AdditionalProperties.text

    # Store domain information and verification record in an object
    $DNSVerification += [PSCustomObject]@{
        Name = $CustomDomain.Id;
        TXTRecord = $MgVerificationCode;
    }
}

# Export DNS verification records to CSV file
$DNSVerification | Export-Csv $File -NoTypeInformation