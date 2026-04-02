# Connect
Connect-ExchangeOnline

# Message
$Message = @"
Thank you for your email.

Our office is currently closed in observance of Passover on April 2nd and April 3rd. We will return on Monday and respond to your message as soon as possible.

Wishing you a happy and meaningful holiday.

Best regards,
"@

# Define time in UTC (converted from Eastern Time)
$Start = Get-Date "2026-04-02T12:00:00Z"  # 8 AM EDT
$End   = Get-Date "2026-04-04T02:00:00Z"  # 10 PM EDT (April 3)

# Get mailboxes
$mailboxes = Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited

# Apply OOO
foreach ($mb in $mailboxes) {
    Write-Host "Setting OOO for:" $mb.UserPrincipalName -ForegroundColor Green

    Set-MailboxAutoReplyConfiguration -Identity $mb.UserPrincipalName `
        -AutoReplyState Scheduled `
        -StartTime $Start `
        -EndTime $End `
        -InternalMessage $Message `
        -ExternalMessage $Message `
        -ExternalAudience All
}

Write-Host "✅ OOO applied correctly (Eastern Time aligned)." -ForegroundColor Cyan