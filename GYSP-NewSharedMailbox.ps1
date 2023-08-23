<#DISCLAIMER:

    By running this PowerShell script, you acknowledge and accept that:
    
    - You execute this script at your own risk and discretion.
    - The author of this script is not responsible for any damage, data loss, or unintended consequences 
      that may result from executing this script.
    - It is recommended to test this script in a controlled environment before use in a production scenario.

    If you do not agree with these terms, please refrain from executing this script.




<#

# Export Mailboxes from source

$Mailboxes = Get-mailbox -ResultSize unlimited
$sharedmailbox = $Mailboxes | Where{$_.RecipientTypeDetails -eq "SharedMailbox"}
$sharedmailbox | Select name,alias,Primary* | Export-Csv SharedMailboxes.csv -NoTypeInformation

#>

# Import the list of new shared mailboxes from the CSV file
$NewSharedMailboxes = Import-Csv SharedMailboxes.csv

# Loop through each new shared mailbox in the list
foreach ($NewSharedMailbox in $NewSharedMailboxes) {
    try {
        # Create a new shared mailbox using New-Mailbox cmdlet
        New-Mailbox -Shared -Name $NewSharedMailbox.Name -DisplayName $NewSharedMailbox.Name -Alias $NewSharedMailbox.Alias
        # Output a success message for each created mailbox
        Write-Host "Created shared mailbox: $($NewSharedMailbox.Name)"
    } catch {
        # If an error occurs during mailbox creation, catch the error and display an error message
        Write-Host "Error creating shared mailbox: $($NewSharedMailbox.Name)"
        Write-Host "Error message: $($_.Exception.Message)"
    }
}
