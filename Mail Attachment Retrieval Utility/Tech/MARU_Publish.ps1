#Requires -Version 5.1
<#
.SYNOPSIS
    Generates MARU_Version.json by calculating SHA256 checksums for all
    managed script files in the current directory.

.DESCRIPTION
    Run this script from the folder containing the MARU scripts whenever
    you publish a new release. It will prompt for the version number and release
    notes, calculate checksums, and write MARU_Version.json to the same
    directory.

    Copy the resulting JSON file (along with the updated scripts) to the version
    subfolder on your network share to complete the release.

.EXAMPLE
    .\MARU_Publish.ps1
    .\MARU_Publish.ps1 -Version "1.2.0" -ReleaseNotes "Fixed subfolder recursion."
    .\MARU_Publish.ps1 -OutputPath "\\server\share\MARU\version"
#>

param(
    # Version string — prompted interactively if not supplied
    [string]$Version,

    # Release notes — prompted interactively if not supplied
    [string]$ReleaseNotes,

    # Where to write the JSON (and optionally copy the scripts).
    # Defaults to the directory containing this script.
    [string]$OutputPath,

    # If set, also copies the managed script files to OutputPath.
    [switch]$CopyScripts
)

$ErrorActionPreference = "Stop"

# ---------------------------------------------------------------------------
# Managed files — must be present in the same directory as this script
# ---------------------------------------------------------------------------

$ManagedFiles = @(
    "MARU.ps1",
    "MARU_UI.ps1",
    "MARU_Update.ps1"
)

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# ---------------------------------------------------------------------------
# Validate managed files exist
# ---------------------------------------------------------------------------

$missing = $ManagedFiles | Where-Object { -not (Test-Path (Join-Path $ScriptDir $_)) }
if ($missing.Count -gt 0) {
    Write-Error "The following managed files were not found in '$ScriptDir':`n  $($missing -join "`n  ")"
    exit 1
}

# ---------------------------------------------------------------------------
# Read existing version.json (if present) so we can show the previous version
# ---------------------------------------------------------------------------

$existingVersion = $null
$existingJsonPath = Join-Path $ScriptDir "MARU_Version.json"
if (Test-Path $existingJsonPath) {
    try {
        $existing = Get-Content $existingJsonPath -Raw | ConvertFrom-Json
        $existingVersion = $existing.Version
    } catch {}
}

# ---------------------------------------------------------------------------
# Prompt for version and release notes if not supplied
# ---------------------------------------------------------------------------

if (-not $Version) {
    $prompt = if ($existingVersion) { "Version number (current: $existingVersion)" } else { "Version number (e.g. 1.0.0)" }
    $Version = Read-Host $prompt
    if ([string]::IsNullOrWhiteSpace($Version)) {
        Write-Error "Version is required."
        exit 1
    }
}

if (-not $ReleaseNotes) {
    $ReleaseNotes = Read-Host "Release notes (optional — press Enter to skip)"
}

# ---------------------------------------------------------------------------
# Calculate checksums
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host "Calculating checksums..." -ForegroundColor Cyan

$fileEntries = foreach ($fileName in $ManagedFiles) {
    $filePath = Join-Path $ScriptDir $fileName
    $hash     = (Get-FileHash -Path $filePath -Algorithm SHA256).Hash
    Write-Host "  $hash  $fileName"
    [PSCustomObject]@{
        Name     = $fileName
        Checksum = $hash
    }
}

# ---------------------------------------------------------------------------
# Build manifest object
# ---------------------------------------------------------------------------

$manifest = [ordered]@{
    Version      = $Version
    ReleasedDate = (Get-Date -Format "yyyy-MM-dd")
    ReleaseNotes = $ReleaseNotes
    Files        = $fileEntries
}

# ---------------------------------------------------------------------------
# Resolve output path and write JSON
# ---------------------------------------------------------------------------

$resolvedOutput = if ($OutputPath) {
    if (-not (Test-Path $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
        Write-Host ""
        Write-Host "Created output directory: $OutputPath" -ForegroundColor Yellow
    }
    $OutputPath
} else {
    $ScriptDir
}

$jsonPath = Join-Path $resolvedOutput "MARU_Version.json"
$manifest | ConvertTo-Json -Depth 5 | Out-File -FilePath $jsonPath -Encoding utf8 -Force

Write-Host ""
Write-Host "Version manifest written to:" -ForegroundColor Green
Write-Host "  $jsonPath"

# ---------------------------------------------------------------------------
# Optionally copy scripts to output path
# ---------------------------------------------------------------------------

if ($CopyScripts -and $resolvedOutput -ne $ScriptDir) {
    Write-Host ""
    Write-Host "Copying scripts to output directory..." -ForegroundColor Cyan
    foreach ($fileName in $ManagedFiles) {
        $src  = Join-Path $ScriptDir $fileName
        $dest = Join-Path $resolvedOutput $fileName
        Copy-Item -Path $src -Destination $dest -Force
        Write-Host "  Copied: $fileName"
    }
}

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host "===============================" -ForegroundColor Cyan
Write-Host "  Version      : $Version"
Write-Host "  Released     : $(Get-Date -Format 'yyyy-MM-dd')"
if ($ReleaseNotes) {
    Write-Host "  Notes        : $ReleaseNotes"
}
Write-Host "  Files hashed : $($ManagedFiles.Count)"
Write-Host "  Output       : $resolvedOutput"
Write-Host "===============================" -ForegroundColor Cyan
Write-Host ""

if (-not $CopyScripts -and $resolvedOutput -eq $ScriptDir) {
    Write-Host "Next step: copy MARU_Version.json and the updated scripts" -ForegroundColor Yellow
    Write-Host "to your network share version folder, or re-run with -OutputPath and -CopyScripts." -ForegroundColor Yellow
    Write-Host ""
}