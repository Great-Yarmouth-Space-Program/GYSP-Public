<#
DISCLAIMER:

    By running this PowerShell script, you acknowledge and accept that:
    
    - You execute this script at your own risk and discretion.
    - The author of this script is not responsible for any damage, data loss, or unintended consequences 
      that may result from executing this script.
    - It is recommended to test this script in a controlled environment before use in a production scenario.

    If you do not agree with these terms, please refrain from executing this script.


    Name:               GYSP-GetMailboxForwarders.PS1
    Version:            NOT TESTED
    Date:               12-09-2023
    Original Author:    Si Ansell
    Email:              graph@greatyarmouthspaceprogram.space

    Use:                Creates a CSV with all mailbox ForwardingAddress and ForwardingSmtpAddress
                        ForwardingAddress - Forwards to INTERNAL users
                        ForwardingSmtpAddress - Forwards to EXTERNAL users
                        

    Updates:            12-09-2023 Disclaimer added and script tidied up - SRA
#>

# Initialize an empty array to store mailboxes
$Mailboxes - @()

# Retrieve all mailboxes without a limit on results
$Mailboxes = Get-mailbox -ResultSize unlimited

# Initialize an empty array to store mailboxes with forwarding settings
$ForwardingMailboxes = @()

# Loop through each mailbox
Foreach ($Mailbox in $Mailboxes) {
    
    # Check if the mailbox has an internal forwarding address set
    If ($Mailbox.ForwardingAddress -ne $Null) {
        # Display the information with green foreground color
        Write-Host $Mailbox.UserPrincipalName "is forwarding to the INTERNAL mailbox" $Mailbox.ForwardingAddress -ForegroundColor Green
        
        # Add the mailbox details to the array
        $ForwardingMailboxes = $ForwardingMailboxes + [PSCustomObject]@{
            Mailbox           = $Mailbox.UserPrincipalName
            InternalForwarding = $Mailbox.ForwardingAddress
            ExternalForwarding = $Mailbox.ForwardingSMTPAddress
        }
    }

    # Check if the mailbox has an external forwarding SMTP address set
    If ($Mailbox.ForwardingSMTPAddress -ne $Null) {
        # Display the information with yellow foreground color
        Write-Host $Mailbox.UserPrincipalName "is forwarding to the EXTENAL mailbox" $Mailbox.ForwardingSMTPAddress -ForegroundColor Yellow
        
        # Add the mailbox details to the array
        $ForwardingMailboxes = $ForwardingMailboxes + [PSCustomObject]@{
            Mailbox           = $Mailbox.UserPrincipalName
            InternalForwarding = $Mailbox.ForwardingAddress
            ExternalForwarding = $Mailbox.ForwardingSMTPAddress 
        }
    }
}

# Check if the variable $Forwarders is not null or empty
If($Forwarders -ne $Null) {
    # Export the contents of the $Forwarders variable to an Excel file
    $Forwarders | Export-Excel -Path .\MailboxForwarders.xlsx -AutoSize -TableName Mailbox_Forwarders -WorksheetName Mailbox_Forwarders 
    
    # Export the contents of the $Forwarders variable to a CSV file
    $Forwarders | Export-CSV -Path .\MailboxForwarders.csv -Notypeinformation
}
