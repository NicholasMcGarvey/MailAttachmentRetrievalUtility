#Requires -Version 5.1
<#
.SYNOPSIS
     updater. Launched by MARU_UI.ps1 when an update is available.
    Copies new files from the share version folder to the local script directory,
    backs up the previous versions, then relaunches MARU_UI.ps1.

.PARAMETER LocalDir
    The local directory where the MARU scripts are installed.

.PARAMETER VersionSourceDir
    The UNC path to the version subfolder on the share containing the new files.

.PARAMETER FilesToUpdate
    Comma-separated list of filenames to copy (from the version manifest).
#>
param(
    [Parameter(Mandatory)][string]$LocalDir,
    [Parameter(Mandatory)][string]$VersionSourceDir,
    [Parameter(Mandatory)][string]$FilesToUpdate
)

Add-Type -AssemblyName PresentationFramework

$ErrorActionPreference = "Stop"

# ---------------------------------------------------------------------------
# Brief pause — gives the calling UI process time to fully exit before we
# attempt to replace any files it may have had open.
# ---------------------------------------------------------------------------
Start-Sleep -Seconds 2

$backupDir = Join-Path $LocalDir "backup"
New-Item -ItemType Directory -Path $backupDir -Force | Out-Null

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$fileList  = $FilesToUpdate -split ','

# ---------------------------------------------------------------------------
# Copy helper with retry — handles the rare case where the UI process hasn't
# fully released file locks despite the initial sleep. Retries up to 3 times
# with 500 ms backoff before giving up.
# ---------------------------------------------------------------------------
function Copy-WithRetry {
    param(
        [string]$Source,
        [string]$Destination,
        [int]$MaxRetries = 3,
        [int]$RetryDelayMs = 500
    )
    $attempt = 0
    while ($true) {
        try {
            Copy-Item -Path $Source -Destination $Destination -Force
            return
        } catch {
            $attempt++
            if ($attempt -ge $MaxRetries) { throw }
            Write-Warning "Copy failed (attempt $attempt of $MaxRetries): $_. Retrying in ${RetryDelayMs}ms..."
            Start-Sleep -Milliseconds $RetryDelayMs
        }
    }
}

$errors = @()

foreach ($fileName in $fileList) {
    $fileName   = $fileName.Trim()
    $sourcePath = Join-Path $VersionSourceDir $fileName
    $destPath   = Join-Path $LocalDir $fileName
    $backupPath = Join-Path $backupDir "$([System.IO.Path]::GetFileNameWithoutExtension($fileName))_$timestamp$([System.IO.Path]::GetExtension($fileName))"

    try {
        # Back up the existing file if present
        if (Test-Path $destPath) {
            Copy-WithRetry -Source $destPath -Destination $backupPath
        }

        # Copy new version into place with retry
        Copy-WithRetry -Source $sourcePath -Destination $destPath

    } catch {
        $errors += "Failed to update '$fileName': $_"
    }
}

if ($errors.Count -gt 0) {
    $msg = "The following errors occurred during update:`n`n" + ($errors -join "`n") +
           "`n`nPrevious versions have been preserved in:`n$backupDir"
    [System.Windows.MessageBox]::Show(
        $msg,
        "Update Errors",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Error) | Out-Null
} else {
    # Relaunch the UI and exit immediately — don't wait for it
    $uiPath = Join-Path $LocalDir "MARU_UI.ps1"
    Start-Process -FilePath "powershell.exe" `
                  -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$uiPath`"" `
                  -WindowStyle Normal
    exit 0
}
