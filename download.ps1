#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Pirum Consulting LLC - Build Script Downloader
.DESCRIPTION
    Downloads the latest PC-Build-Script package from GitHub and extracts it
    to C:\Pirum\PC-Build-Script-master, ready for use with PCSetup_v2.ps1.
.NOTES
    Pirum Consulting LLC | 330-597-0550 | pirumllc.com
#>

Set-ExecutionPolicy RemoteSigned -Scope Process -Force

$RepoUrl    = "https://github.com/jokerjim/PC-Build-Script/archive/master.zip"
$PirumDir   = "C:\Pirum"
$ZipPath    = "$PirumDir\PCBuild.zip"
$ExtractDir = $PirumDir

function Write-Status {
    param([string]$Message, [string]$Level = "INFO")
    $colors = @{ INFO = "Cyan"; OK = "Green"; WARN = "Yellow"; ERROR = "Red" }
    $prefix = @{ INFO = "   "; OK = " OK"; WARN = "WARN"; ERROR = " ERR" }
    Write-Host "[$($prefix[$Level])] $Message" -ForegroundColor $colors[$Level]
}

# ── Re-launch as admin if needed ──
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]"Administrator")) {
    Write-Status "Relaunching as Administrator..." "WARN"
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    Exit
}

Write-Host ""
Write-Host "  Pirum Consulting LLC - Build Script Downloader" -ForegroundColor DarkMagenta
Write-Host "  ------------------------------------------------" -ForegroundColor DarkMagenta
Write-Host ""

# ── Ensure C:\Pirum exists ──
if (-not (Test-Path $PirumDir)) {
    Write-Status "Creating $PirumDir..."
    New-Item -Path $PirumDir -ItemType Directory -Force | Out-Null
    Write-Status "$PirumDir created." "OK"
} else {
    Write-Status "$PirumDir already exists."
}

# ── Enforce TLS 1.2 for GitHub ──
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

# ── Download ──
Write-Status "Downloading build script package from GitHub..."
Write-Status "  $RepoUrl"
try {
    Invoke-WebRequest -Uri $RepoUrl -OutFile $ZipPath -UseBasicParsing -ErrorAction Stop
    $zipSize = (Get-Item $ZipPath).Length / 1KB
    Write-Status "Download complete. ($([math]::Round($zipSize, 1)) KB)" "OK"
} catch {
    Write-Status "Download failed: $_" "ERROR"
    Write-Host ""
    Read-Host "Press Enter to exit"
    Exit 1
}

# ── Extract ──
Write-Status "Extracting to $ExtractDir..."
try {
    # Remove existing extracted folder if present so we get a clean copy
    $existingFolder = "$ExtractDir\PC-Build-Script-master"
    if (Test-Path $existingFolder) {
        Write-Status "Removing existing $existingFolder..."
        Remove-Item $existingFolder -Recurse -Force -ErrorAction Stop
    }
    Expand-Archive -Path $ZipPath -DestinationPath $ExtractDir -Force -ErrorAction Stop
    Write-Status "Extraction complete." "OK"
} catch {
    Write-Status "Extraction failed: $_" "ERROR"
    Write-Host ""
    Read-Host "Press Enter to exit"
    Exit 1
}

# ── Clean up zip ──
Write-Status "Removing downloaded zip file..."
Remove-Item $ZipPath -Force -ErrorAction SilentlyContinue
Write-Status "Cleanup complete." "OK"

# ── Confirm result ──
Write-Host ""
if (Test-Path "$ExtractDir\PC-Build-Script-master") {
    Write-Status "Build script ready at: $ExtractDir\PC-Build-Script-master" "OK"
    Write-Status "You can now run PCSetup_v2.ps1." "OK"
} else {
    Write-Status "Expected folder not found after extraction. Check $ExtractDir manually." "WARN"
}

Write-Host ""
Read-Host "Press Enter to exit"
