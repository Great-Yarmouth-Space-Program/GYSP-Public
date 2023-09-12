<#
    Name:               GYSP-SetMessageLimits.PS1
    Version:            1.0
    Date:               13-08-2023
    Original Author:    Si Ansell
    Email:              graph@greatyarmouthspaceprogram.space

   Modules:             ExchangeOnlineManagement
                        ImportExcel
   
    Use:                Configure ALL mailboxes 
                        OR
                        Create Mailboxes.csv with a single column headed UserPrincipalName. 
                        Populate with required UPNs
                        Set $MaxSendSize and $MaxReceiveSize to required limits. Default is 35/36, Max is 150
                        Run script

    Updates:        
#>

# Get ALL mailboxes or Import mailboxes from CSV
#$Mailboxes = Get-Mailbox -ResultSize Unlimited
#OR
$Mailboxes = Import-Csv .\Mailboxes.csv

# Set max send and receive sizes
$MaxSendSize = "35mb"
$MaxReceiveSize = "36mb"

# Array to store message limits data
$MessageLimits = @()

# Loop through each mailbox and update limits
ForEach ($Mailbox in $Mailboxes) {
    $LegacyLimits = Get-Mailbox -Identity $Mailbox.UserPrincipalName
    Set-Mailbox -Identity $Mailbox.UserPrincipalName -MaxSendSize $MaxSendSize -MaxReceiveSize $MaxReceiveSize
    $NewLimits = Get-Mailbox -Identity $Mailbox.UserPrincipalName

    # Create custom object to store limits data
    $MessageLimitObject = [PSCustomObject]@{
        Mailbox = $Mailbox.UserPrincipalName
        LegacyMaxSendSize = $LegacyLimits.MaxSendSize
        LegacyMaxReceiveSize = $LegacyLimits.MaxReceiveSize
        NewMaxSendSize = $NewLimits.MaxSendSize
        NewMaxReceiveSize = $NewLimits.MaxReceiveSize
    }

    # Add the custom object to the array
    $MessageLimits += $MessageLimitObject
}

# Export data to Excel
$MessageLimits | Export-Excel -Path .\MessageLimits.xlsx -AutoSize -TableName Message_Limits -WorksheetName Message_Limits