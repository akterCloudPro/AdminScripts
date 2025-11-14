<#
Simple: Export INBOUND message trace for a mailbox for the last N days.
Save as: C:\Export-Inbound-MessageTrace.ps1
#>

# ----- Configuration -----
$Mailbox  = "akter@clickforbaby.com"   # mailbox to check (recipient)
$DaysBack = 7                         # set 7, 3, or any number
$Output   = "C:\Inbound_MessageTrace.csv"
# --------------------------

# Connect to Exchange Online (interactive)
Connect-ExchangeOnline -ShowBanner:$false

# Choose cmdlet
$useV2 = [bool](Get-Command -Name Get-MessageTraceV2 -ErrorAction SilentlyContinue)

Write-Host "Querying inbound traces for $Mailbox for last $DaysBack days using $(if ($useV2){'Get-MessageTraceV2'} else {'Get-MessageTrace'})..."

$all = @()
$endUtc   = (Get-Date).ToUniversalTime()
$startUtc = $endUtc.AddDays(-$DaysBack)

for ($d = 0; $d -lt $DaysBack; $d++) {
    $s = $startUtc.AddDays($d)
    $e = $s.AddDays(1)
    Write-Host " Window: $($s.ToString('u')) -> $($e.ToString('u'))"

    try {
        if ($useV2) {
            $batch = Get-MessageTraceV2 -RecipientAddress $Mailbox -StartDate $s -EndDate $e -ErrorAction Stop
        } else {
            $batch = Get-MessageTrace -RecipientAddress $Mailbox -StartDate $s -EndDate $e -ErrorAction Stop
        }
    } catch {
        Write-Warning "Query failed for window starting $s : $($_.Exception.Message)"
        continue
    }

    if ($batch) { $all += $batch }
}

if ($all.Count -gt 0) {
    $all |
        Sort-Object Received |
        Select-Object Received, SenderAddress, RecipientAddress, Subject, Status, Size, MessageId, MessageTraceId, Direction, DeliveryTime |
        Export-Csv -Path $Output -NoTypeInformation -Encoding UTF8

    Write-Host "Inbound CSV saved to: $Output" -ForegroundColor Green
} else {
    Write-Warning "No inbound message trace records found for $Mailbox in the last $DaysBack days."
}

# Disconnect
Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
