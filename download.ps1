#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Pirum Consulting LLC - Build Script Downloader
.DESCRIPTION
    Downloads the latest PC-Build-Script package from GitHub and extracts it
    to C:\Pirum\PC-Build-Script-master, ready for use with PCSetup.ps1.
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

function Wait-WithCountdown {
    param([int]$Seconds = 30, [string]$Prompt = "Press Enter to exit")
    Write-Host ""
    for ($i = $Seconds; $i -gt 0; $i--) {
        Write-Host "`r  $Prompt  (closing in $i seconds)...  " -NoNewline
        $start = [DateTime]::Now
        while (([DateTime]::Now - $start).TotalMilliseconds -lt 1000) {
            if ([Console]::KeyAvailable) {
                [Console]::ReadKey($true) | Out-Null
                Write-Host ""
                return
            }
            Start-Sleep -Milliseconds 50
        }
    }
    Write-Host "`r  Closing automatically.                              "
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
    Wait-WithCountdown -Seconds 30 -Prompt "Press Enter to exit"
    Exit 1
}

# ── Extract ──
# Clear contents of the existing folder without deleting it.
# This removes stale files while keeping the folder intact so any running
# instance of PCSetup.ps1 keeps its file handle and continues executing.
$existingFolder = "$ExtractDir\PC-Build-Script-master"
if (Test-Path $existingFolder) {
    Write-Status "Clearing existing contents of $existingFolder..."
    Get-ChildItem -Path $existingFolder -Recurse | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
}
Write-Status "Extracting to $ExtractDir..."
try {
    Expand-Archive -Path $ZipPath -DestinationPath $ExtractDir -Force -ErrorAction Stop
    Write-Status "Extraction complete." "OK"
} catch {
    Write-Status "Extraction failed: $_" "ERROR"
    Write-Host ""
    Wait-WithCountdown -Seconds 30 -Prompt "Press Enter to exit"
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
    Write-Status "You can now run PCSetup.ps1." "OK"
} else {
    Write-Status "Expected folder not found after extraction. Check $ExtractDir manually." "WARN"
}

Write-Host ""
Wait-WithCountdown -Seconds 30 -Prompt "Press Enter to exit"
