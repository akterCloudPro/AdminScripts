# ===============================================================
# Copy STRICTLY INACTIVE user profiles to NAS
# Then delete profile AFTER successful copy
# FUNCTION BASED – NTUSER.DAT SAFE
# ===============================================================

# ---------------- CONFIG ----------------
$DaysToCheck       = 30
$NASServer         = "192.168.4.223"
$NASShare          = "User_Profile_Archive"
$NASUser           = "NASadmin"
$NASPassword       = "P@ssword4N#"
$UseComputerFolder = $true
$RobocopyTimeout   = 1800000
# ---------------------------------------

# ---- Build NAS paths ----
$ComputerFolder = if ($UseComputerFolder) { $env:COMPUTERNAME } else { "" }
$NASSharePath   = "\\$NASServer\$NASShare"
$NASPath        = Join-Path $NASSharePath $ComputerFolder

# ---- Credential ----
$SecurePass = ConvertTo-SecureString $NASPassword -AsPlainText -Force
$Credential = New-Object System.Management.Automation.PSCredential ($NASUser, $SecurePass)

# ===============================================================
# FUNCTION: Copy profile data
# ===============================================================
function Copy-InactiveUserProfile {
    param (
        [CimInstance]$Profile,
        [string]$DestinationRoot
    )

    $Folders = @("Desktop","Documents","Downloads")
    $Failed  = $false

    foreach ($folder in $Folders) {

        $Source = Join-Path $Profile.LocalPath $folder
        $Dest   = Join-Path $DestinationRoot $folder

        if (Test-Path $Source) {

            $Proc = Start-Process robocopy.exe -ArgumentList @(
                "`"$Source`"",
                "`"$Dest`"",
                "/E","/Z","/R:2","/W:2",
                "/COPY:DAT","/XJ","/NFL","/NDL"
            ) -NoNewWindow -PassThru

            if (-not $Proc.WaitForExit($RobocopyTimeout) -or $Proc.ExitCode -ge 8) {
                Write-Warning "Copy failed: $($Profile.LocalPath)\$folder"
                $Failed = $true
            }
        }
    }

    return (-not $Failed)
}

# ===============================================================
# FUNCTION: Delete profile safely
# ===============================================================
function Remove-InactiveUserProfile {
    param ([CimInstance]$Profile)

    try {
        if (-not $Profile.Loaded) {
            Remove-CimInstance $Profile -ErrorAction Stop
            Write-Host "  Profile deleted successfully" -ForegroundColor Green
        }
    }
    catch {
        Write-Warning "  FAILED to delete profile: $_"
    }
}

# ---- Connect NAS ----
try {
    New-PSDrive -Name NASBackup -PSProvider FileSystem -Root $NASSharePath `
        -Credential $Credential -ErrorAction Stop | Out-Null
}
catch {
    Write-Error "Failed to connect to NAS"
    exit 1
}

if (!(Test-Path $NASPath)) {
    New-Item $NASPath -ItemType Directory -Force | Out-Null
}

# ---- Cutoff ----
$CutoffDate = (Get-Date).AddDays(-$DaysToCheck)
Write-Host "Evaluating inactivity before $CutoffDate" -ForegroundColor Yellow

# ---- Profiles ----
$AllProfiles = Get-CimInstance Win32_UserProfile | Where-Object {
    -not $_.Special -and
    $_.LocalPath -and
    $_.LocalPath -notmatch "ServiceProfiles|SystemProfile|DefaultAppPool|defaultuser0"
}

# ---- Login activity ----
$Events = Get-WinEvent -FilterHashtable @{
    LogName   = 'Security'
    Id        = 4624
    StartTime = $CutoffDate
} -ErrorAction Stop

$ActiveUsers = $Events | ForEach-Object {
    if ($_.Properties.Count -ge 6) {
        ($_.Properties[5].Value -split '\\')[-1]
    }
} | Sort-Object -Unique

# ---- Detect inactive profiles ----
$FinalProfiles = @()
$FoldersToCheck = @("Desktop","Documents","Downloads")

foreach ($profile in $AllProfiles) {

    if ($profile.Loaded) { continue }

    $UserName = Split-Path $profile.LocalPath -Leaf
    if ($ActiveUsers -contains $UserName) { continue }

    # ---- NTUSER.DAT SAFE CHECK ----
    $NtUser = Join-Path $profile.LocalPath "NTUSER.DAT"
    try {
        $NtItem = Get-Item $NtUser -ErrorAction Stop
        if ($NtItem.LastWriteTime -gt $CutoffDate) { continue }
    }
    catch {
        # Missing NTUSER.DAT = treat as inactive
    }

    # ---- Folder activity ----
    $Changed = $false
    foreach ($f in $FoldersToCheck) {
        $Path = Join-Path $profile.LocalPath $f
        if (Test-Path $Path) {
            if ((Get-Item $Path).LastWriteTime -gt $CutoffDate) {
                $Changed = $true
                break
            }
        }
    }

    if ($Changed) { continue }

    $FinalProfiles += $profile
}

if (-not $FinalProfiles) {
    Write-Host "No inactive profiles detected."
    Remove-PSDrive NASBackup
    exit
}

Write-Host "Profiles approved: $($FinalProfiles.Count)" -ForegroundColor Green

# ---- COPY → DELETE ----
foreach ($profile in $FinalProfiles) {

    $UserName = Split-Path $profile.LocalPath -Leaf
    Write-Host "`nProcessing $UserName" -ForegroundColor Cyan

    $DestRoot = Join-Path $NASPath $UserName
    if (!(Test-Path $DestRoot)) {
        New-Item $DestRoot -ItemType Directory -Force | Out-Null
    }

    if (Copy-InactiveUserProfile -Profile $profile -DestinationRoot $DestRoot) {
        Remove-InactiveUserProfile -Profile $profile
    }
    else {
        Write-Warning "Skipping delete for $UserName (copy incomplete)"
    }
}

# ---- Cleanup ----
Remove-PSDrive NASBackup -ErrorAction SilentlyContinue
Write-Host "`nBackup completed successfully!" -ForegroundColor Green
Write-Host "Data location: $NASPath" -ForegroundColor Green
exit
