function Remove-Automapping 
    {
        Param(
          [Parameter(Mandatory = $true)]
          [string] $Mailbox
        )

        foreach($mailbox in $mailboxes)
            {
                $mailboxPermissions = Get-Mailbox $mailbox | Get-MailboxPermission | ? {$_.AccessRights -eq "FullAccess" -and $_.User -ne "NT AUTHORITY\SELF" `
                -and $_.IsInherited -eq $false}
                foreach($mailboxPermission in $mailboxPermissions)
                    {
                        Get-Mailbox $mailbox | Remove-MailboxPermission -User $mailboxPermission.user -AccessRights $mailboxPermission.AccessRights -Confirm:$false
                        Get-Mailbox $mailbox | Add-MailboxPermission -User $mailboxPermission.user -AccessRights $mailboxPermission.AccessRights -AutoMapping:$false | Out-Null
                    }
            }

    }