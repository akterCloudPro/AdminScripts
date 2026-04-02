# Get all user mailboxes
$mailboxes = Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited

# Disable OOO for all
foreach ($mb in $mailboxes) {
    Write-Host "Disabling OOO for:" $mb.UserPrincipalName -ForegroundColor Yellow

    Set-MailboxAutoReplyConfiguration -Identity $mb.UserPrincipalName -AutoReplyState Disabled
}

Write-Host "✅ OOO disabled for all mailboxes." -ForegroundColor Cyan