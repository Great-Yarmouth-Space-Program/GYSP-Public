




<#

Export Mailboxeds from source

$mailboxes = Get-mailbox -ResultSize unlimited
$sharedmailbox = $mailboxes | Where{$_.RecipientTypeDetails -eq "SharedMailbox"}
$sharedmailbox | Select name,alias,Primary* | Export-Csv SharedMailboxes.csv -NoTypeInformation

#>

$NewSharedmailboxes = Import-Csv SharedMailboxes.csv

Foreach($NewSharedmailbox in $NewSharedmailboxes) {

#New-Mailbox -Shared -Name $NewSharedmailbox.Name -DisplayName $NewSharedmailbox.Name -Alias $NewSharedmailbox.Alias 

}