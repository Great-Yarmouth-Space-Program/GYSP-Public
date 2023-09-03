# Output will be added to C:\temp folder. Open the Remove-SMTP-Address.log with a text editor. For example, Notepad.
Start-Transcript -Path C:\temp\Remove-SMTP-Address.log -Append

# Get all mailboxes
$Mailboxes = Get-Mailbox -ResultSize Unlimited

# Loop through each mailbox
foreach ($Mailbox in $Mailboxes) {

    # Change @contoso.com to the domain that you want to remove
    $Mailbox.EmailAddresses | Where-Object { ($_ -clike "smtp*") -and ($_ -like "*@Urnerbarry.com") } | 

    # Perform operation on each item
    ForEach-Object {

        # Remove the -WhatIf parameter after you tested and are sure to remove the secondary email addresses
        Set-Mailbox $Mailbox.DistinguishedName -EmailAddresses @{remove = $_ } #-WhatIf

        # Write output
        Write-Host "Removing $_ from $Mailbox Mailbox" -ForegroundColor Green
    }
}

Stop-Transcript