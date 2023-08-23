<#DISCLAIMER:

    By running this PowerShell script, you acknowledge and accept that:
    
    - You execute this script at your own risk and discretion.
    - The author of this script is not responsible for any damage, data loss, or unintended consequences 
      that may result from executing this script.
    - It is recommended to test this script in a controlled environment before use in a production scenario.

    If you do not agree with these terms, please refrain from executing this script.




<#

Export Mailboxeds from source

$mailboxes = Get-mailbox -ResultSize unlimited
$sharedmailbox = $mailboxes | Where{$_.RecipientTypeDetails -eq "SharedMailbox"}
$sharedmailbox | Select name,alias,Primary* | Export-Csv SharedMailboxes.csv -NoTypeInformation

#>

$NewSharedmailboxes = Import-Csv SharedMailboxes.csv

Foreach($NewSharedmailbox in $NewSharedmailboxes) {

New-Mailbox -Shared -Name $NewSharedmailbox.Name -DisplayName $NewSharedmailbox.Name -Alias $NewSharedmailbox.Alias 

}