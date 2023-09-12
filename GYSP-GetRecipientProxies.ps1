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
                        Export-Excel
    
    Connections:        Connect-ExchangeOnline
    
    Updates:
#>
# Retrieve all accepted domains
$AcceptedDomains = Get-AcceptedDomain

# Initialize an array to store SMTP addresses of recipients
$RecipientSMTPAddressArray = @()

# Loop through each accepted domain
Foreach($AcceptedDomain in $AcceptedDomains) {
    # Initialize a temporary recipient array
    $Recipients = @()

    # Construct a descriptor to match recipients' email addresses
    $CustomDomain = $AcceptedDomain.Name
    $Descriptor = "*@" + $CustomDomain
    
    # Fetch recipients with email addresses that match the current domain
    $Recipients = Get-Recipient -ResultSize Unlimited | Where-Object{$_.EmailAddresses -match $CustomDomain} 

    # Loop through the list of retrieved recipients
    Foreach($Recipient in $Recipients) {
        # Extract SMTP addresses from the recipient
        $RecipientAddresses = $Recipient.EmailAddresses | Where-Object { $_ -match '^smtp:'}

        # Loop through each SMTP address
        Foreach($RecipientAddress in $RecipientAddresses) {
            # Extract the address type (e.g., 'smtp')
            $Type = $RecipientAddress.SubString(0,4)
            
            # Extract the actual email address without the type
            $Address = $RecipientAddress.substring(5) 

            # If the SMTP address matches the descriptor for the current domain
            If($RecipientAddress -Like $Descriptor) {
                # Add the address and other details to the main array
                $RecipientSMTPAddressArray = $RecipientSMTPAddressArray + [PSCustomObject]@{
                    WindowsLiveID = $Recipient.WindowsLiveID
                    Identity = $Recipient.Identity
                    Type = $Type
                    SMTPAddress = $Address
                    AcceptedDomain = $AcceptedDomain.Domainname
                    RecipientType = $Recipient.RecipientType
                    RecipientTypeDetails = $Recipient.RecipientTypeDetails
                }
            }
        }
    }
}



# Check if there's any data in the array
If($RecipientSMTPAddressArray -ne $Null) {
    # Export the collected addresses to an Excel file after sorting by WindowsLiveID
    $RecipientSMTPAddressArray | Sort-Object WindowsliveID | Export-Excel -Path .\RecipientProxiesMASTER.xlsx -AutoSize -TableName Recipient_Proxies -WorksheetName Recipient_Proxies 

    # Export the collected addresses to a CSV file after sorting by WindowsLiveID
    $RecipientSMTPAddressArray | Sort-Object WindowsliveID | CSV -Path .\RecipientProxiesMASTER.csv -NoTypeInformation  
}
```