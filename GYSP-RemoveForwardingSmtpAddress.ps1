<#
    Name:               GYSP-RemoveForwardingSMTPAddress.PS1
    Version:            1.0
    Date:               21-08-2023
    Original Author:    Si Ansell
    Email:              graph@greatyarmouthspaceprogram.space

    Use:                Removes the ForwardingSMTPAddress from mailboxes post a QODM Mailbox migration
                        Makesure to enter the $Forwarder forwarding domain variable 

    Updates:        
#>


$Forwarder = "*FORWARDINGDOMIN" # Enter the forwarding domain here

$Mailboxes = Get-Mailbox -ResultSize unlimited | Where{$_.ForwardingSmtpAddress -like $Forwarder}

Foreach($mailbox in $mailboxes) {

    Set-Mailbox $Mailbox -ForwardingSmtpAddress $Null

}