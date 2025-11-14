<#
.SYNOPSIS
    Export daily SharePoint and OneDrive file activity counts from the Unified Audit Log.

.DESCRIPTION
    This script connects to Exchange Online (requires Audit Log Search / Compliance permissions),
    retrieves file-related audit events (accessed, modified, downloaded, uploaded, copied, deleted, moved)
    for a specified user and date range, summarizes the data daily, and exports both detailed and summary
    CSV reports to a local folder.
#>

##########################################################################
##  Unified Audit Log: SharePoint / OneDrive daily file activity counts ##
##########################################################################

# Install/Import ExchangeOnlineManagement
Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force

# Connect (must have Audit Log Search / Compliance permissions)
Connect-ExchangeOnline -UserPrincipalName admin@domain.com

# === VARIABLES ===
$User  = "user@domain.com"
$Start = (Get-Date "2025-11-10T00:00:00Z")
$End   = (Get-Date "2025-11-13T23:59:59Z")

# Common SharePoint/OneDrive operations to capture
$ops = @(
    "FileAccessed", "FileModified", "FileDownloaded",
    "FileUploaded", "FileCopied", "FileDeleted", "FileMoved"
)

Write-Host "Pulling Unified Audit Logs for $User from $Start to $End..." -ForegroundColor Cyan

# === GET AUDIT LOGS ===
$results = Search-UnifiedAuditLog -StartDate $Start -EndDate $End -UserIds $User -Operations $ops -ResultSize 5000

if (-not $results) {
    Write-Host "No audit records found for this period." -ForegroundColor Yellow
    Disconnect-ExchangeOnline -Confirm:$false
    return
}

# === PARSE EVENTS ===
$events = $results | ForEach-Object {
    $j = $_.AuditData | ConvertFrom-Json

    # Extract file location/path if available
    $fileLocation = $null
    if ($j.SourceRelativeUrl) {
        $fileLocation = $j.SiteUrl.TrimEnd('/') + "/" + $j.SourceRelativeUrl.TrimStart('/')
    }
    elseif ($j.ObjectId -match "^https?://") {
        $fileLocation = $j.ObjectId
    }
    elseif ($j.SiteUrl) {
        $fileLocation = $j.SiteUrl
    }

    [PSCustomObject]@{
        Time        = [datetime]$j.CreationTime
        Operation   = $j.Operation
        User        = $j.UserId
        FileName    = $j.ObjectId
        FilePath    = $fileLocation
        SiteUrl     = $j.SiteUrl
        Detail      = ($j | Select-Object -Property * | ConvertTo-Json -Compress)
    }
}

# === DAILY SUMMARY ===
$dailyCounts = $events |
    Group-Object @{Expression={($_.Time).ToString('yyyy-MM-dd')}} ,
                 @{Expression={$_.Operation}} |
    Select-Object @{Name='Date';Expression={$_.Name.Split(',')[0].Trim()}},
                  @{Name='Operation';Expression={$_.Name.Split(',')[1].Trim()}},
                  @{Name='Count';Expression={$_.Count}} |
    Sort-Object Date, Operation

Write-Host "`n=== Daily Summary ===" -ForegroundColor Green
$dailyCounts | Format-Table -AutoSize

# === EXPORT RESULTS ===
$timestamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
$outFolder = "C:\Forensics"
if (!(Test-Path $outFolder)) { New-Item -ItemType Directory -Path $outFolder | Out-Null }

$detailCsv = Join-Path $outFolder "SPOD_Activity_Detail_$($User -replace '@','_')_$timestamp.csv"
$summaryCsv = Join-Path $outFolder "SPOD_Activity_Summary_$($User -replace '@','_')_$timestamp.csv"

$events | Export-Csv $detailCsv -NoTypeInformation -Encoding UTF8
$dailyCounts | Export-Csv $summaryCsv -NoTypeInformation -Encoding UTF8

Write-Host "`nExport complete!"
Write-Host "Detailed log: $detailCsv"
Write-Host "Summary log:  $summaryCsv" -ForegroundColor Cyan

Disconnect-ExchangeOnline -Confirm:$false
