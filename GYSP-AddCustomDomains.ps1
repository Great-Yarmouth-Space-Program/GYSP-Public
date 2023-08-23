<#
DISCLAIMER:

    By running this PowerShell script, you acknowledge and accept that:
    
    - You execute this script at your own risk and discretion.
    - The author of this script is not responsible for any damage, data loss, or unintended consequences 
      that may result from executing this script.
    - It is recommended to test this script in a controlled environment before use in a production scenario.

    If you do not agree with these terms, please refrain from executing this script.


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

# Disconnect from the Microsoft Graph API
Disconnect-MgGraph

# Remove the cached Graph API token
Remove-Item "$env:USERPROFILE\.graph" -Recurse -Force


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
$DNSVerification | Export-Csv .\domains.csv -NoTypeInformation