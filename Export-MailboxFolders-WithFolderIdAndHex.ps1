####################################################################
# Export-AronMailboxFolders-WithFolderIdAndHex.ps1
#
# Purpose:
#   Lists all folders for a given mailbox
#   and outputs the following columns:
#     - Name
#     - FolderPath
#     - FolderType
#     - FolderId     (Base64)
#     - FolderHexId  (48-character hex derived from FolderId)
#
# Notes:
#   - You are already connected to Exchange Online 
#   - The script will NOT try to reconnect.
#   - Optionally export results to CSV (full untruncated fields).
#
# Usage:
#   .\Export-AronMailboxFolders-WithFolderIdAndHex.ps1
#   .\Export-AronMailboxFolders-WithFolderIdAndHex.ps1 -OutputCsv "C:\Reports\UserFolders.csv"
#   .\Export-AronMailboxFolders-WithFolderIdAndHex.ps1 -Mailbox "someone@tdomain.com"
####################################################################

param(
    [string] $Mailbox = '',
    [string] $OutputCsv = ''   # optional path to export CSV (leave blank to skip)
)

# ------------------------------------------------------------------
# Function: Convert-Base64FolderIdToHex
# Converts the FolderId (base64) to a 48-character hex value
# ------------------------------------------------------------------
function Convert-Base64FolderIdToHex {
    param([Parameter(Mandatory=$true)][string] $Base64FolderId)

    if ([string]::IsNullOrEmpty($Base64FolderId)) { return $null }

    try {
        $folderIdBytes = [Convert]::FromBase64String($Base64FolderId)
    } catch {
        Write-Verbose "Invalid Base64 for FolderId: $Base64FolderId"
        return $null
    }

    # Algorithm expects at least 47 bytes after decode
    if ($folderIdBytes.Length -lt 47) {
        Write-Verbose "Decoded FolderId too short (Length = $($folderIdBytes.Length))"
        return $null
    }

    $encoding = [System.Text.Encoding]::GetEncoding("us-ascii")
    $nibbler  = $encoding.GetBytes("0123456789ABCDEF")

    $indexIdBytes = New-Object byte[] 48
    $indexIdIdx = 0
    $slice = $folderIdBytes[23..46]  # 24 bytes

    foreach ($b in $slice) {
        $indexIdBytes[$indexIdIdx++] = $nibbler[($b -shr 4)]
        $indexIdBytes[$indexIdIdx++] = $nibbler[($b -band 0xF)]
    }

    return $encoding.GetString($indexIdBytes)
}

# ------------------------------------------------------------------
# MAIN
# ------------------------------------------------------------------
try {
    Write-Host "Retrieving folder statistics for mailbox:" $Mailbox -ForegroundColor Cyan

    # Pull folder stats
    $folders = Get-MailboxFolderStatistics -Identity $Mailbox -ErrorAction Stop

    if (-not $folders) {
        Write-Warning "No folders returned for mailbox $Mailbox."
        return
    }

    # Build results including both FolderId (base64) and FolderHexId (48-char hex)
    $results = foreach ($f in $folders) {
        $base64 = $f.FolderId
        $hex    = Convert-Base64FolderIdToHex -Base64FolderId $base64

        [PSCustomObject]@{
            Name        = $f.Name
            FolderPath  = $f.FolderPath
            FolderType  = ($f.PSObject.Properties.Match('FolderType') | ForEach-Object { $_.Value }) -join ','
            FolderId    = $base64
            FolderHexId = $hex
        }
    }

    # Console-friendly display (truncate long values for readability)
    $results |
        Select-Object `
            @{n='Name';e={$_.Name}},
            @{n='FolderPath';e={ if ($_.FolderPath -and $_.FolderPath.Length -gt 60) { $_.FolderPath.Substring(0,57) + '...' } else { $_.FolderPath } }},
            FolderType,
            @{n='FolderId (base64)';e={ if ($_.FolderId -and $_.FolderId.Length -gt 60) { $_.FolderId.Substring(0,57) + '...' } else { $_.FolderId } }},
            @{n='FolderHexId';e={ if ($_.FolderHexId) { $_.FolderHexId } else { '<n/a>' } }} |
        Format-Table -AutoSize

    # Optional: export full untruncated CSV with exact column names
    if ($OutputCsv -and $OutputCsv.Trim() -ne '') {
        $dir = Split-Path -Path $OutputCsv -Parent
        if ($dir -and -not (Test-Path $dir)) {
            New-Item -ItemType Directory -Path $dir | Out-Null
        }

        Write-Host "Saving full results to CSV:" $OutputCsv -ForegroundColor Cyan

        # Ensure CSV columns order and names are exactly as requested
        $results |
            Select-Object Name, FolderPath, FolderType, FolderId, FolderHexId |
            Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding ASCII
    }

} catch {
    Write-Error ("Error while retrieving folders for {0}: {1}" -f $Mailbox, $_.Exception.Message)
    Write-Host "Ensure that your admin account (tech@thefwdgroup.com) has permission to view mailbox folder statistics for this user."
}
