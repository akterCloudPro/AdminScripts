<#
.SYNOPSIS
    Export User mailbox and archive sizes for all Exchange Online mailboxes to CSV.
.DESCRIPTION
    This script connects to Exchange Online using the Exchange Online PowerShell module,
    retrieves user mailbox and archive sizes, and exports them to a CSV file.
#>

# Connect to Exchange Online
Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
Connect-ExchangeOnline -UserPrincipalName (Read-Host "Enter your admin UPN")

# Get all user mailboxes
Write-Host "Retrieving mailbox information..." -ForegroundColor Cyan
$mailboxes = Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited

# Initialize results array
$results = @()

foreach ($mbx in $mailboxes) {
    Write-Host "Processing mailbox: $($mbx.DisplayName)" -ForegroundColor Yellow

    # Get primary mailbox stats
    $mbxStats = Get-MailboxStatistics -Identity $mbx.UserPrincipalName

    # Parse the TotalItemSize string (e.g. "12.34 GB (13,250,000 bytes)")
    $primarySizeText = $mbxStats.TotalItemSize.ToString().Split("(")[0].Trim()
    $primarySizeMB = 0
    if ($primarySizeText -match "([\d\.]+)\s*(\w+)") {
        $value = [double]$matches[1]
        $unit = $matches[2].ToUpper()
        switch ($unit) {
            "B"  { $primarySizeMB = $value / 1MB }
            "KB" { $primarySizeMB = $value / 1KB }
            "MB" { $primarySizeMB = $value }
            "GB" { $primarySizeMB = $value * 1024 }
            "TB" { $primarySizeMB = $value * 1024 * 1024 }
        }
    }

    # Get archive stats (if active)
    $archiveSizeMB = 0
    if ($mbx.ArchiveStatus -eq "Active") {
        $archiveStats = Get-MailboxStatistics -Identity $mbx.UserPrincipalName -Archive
        $archiveSizeText = $archiveStats.TotalItemSize.ToString().Split("(")[0].Trim()
        if ($archiveSizeText -match "([\d\.]+)\s*(\w+)") {
            $value = [double]$matches[1]
            $unit = $matches[2].ToUpper()
            switch ($unit) {
                "B"  { $archiveSizeMB = $value / 1MB }
                "KB" { $archiveSizeMB = $value / 1KB }
                "MB" { $archiveSizeMB = $value }
                "GB" { $archiveSizeMB = $value * 1024 }
                "TB" { $archiveSizeMB = $value * 1024 * 1024 }
            }
        }
    }

    # Create output object
    $results += [PSCustomObject]@{
        DisplayName            = $mbx.DisplayName
        UserPrincipalName      = $mbx.UserPrincipalName
        PrimaryMailboxSize_MB  = [math]::Round($primarySizeMB, 2)
        ArchiveMailboxSize_MB  = [math]::Round($archiveSizeMB, 2)
        TotalSize_MB           = [math]::Round(($primarySizeMB + $archiveSizeMB), 2)
    }
}

# Export to CSV
$outputFile = "MailboxSizes_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$results | Sort-Object TotalSize_MB -Descending | Export-Csv -Path $outputFile -NoTypeInformation

Write-Host "`nExport complete!" -ForegroundColor Green
Write-Host "File saved as: $outputFile" -ForegroundColor Cyan

# Disconnect
Disconnect-ExchangeOnline -Confirm:$false
