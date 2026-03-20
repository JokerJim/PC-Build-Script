#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Pirum Consulting LLC - PC Setup & Configuration Tool v2.0
.DESCRIPTION
    GUI-driven workstation deployment tool for Windows 10 and Windows 11.
    Each major function is individually selectable via checkbox UI before execution.
    Supports app installation, system hardening, personalization, and management agent deployment.
.NOTES
    Pirum Consulting LLC | 330-597-0550 | pirumllc.com
    Run as Administrator. All steps are optional and individually controlled.
#>

Set-ExecutionPolicy RemoteSigned -Scope Process -Force

# ============================================================
# BASELINE APP LIST
# To add or remove baseline apps, edit this list only.
# Each entry: DisplayName, ChocoID, DefaultChecked ($true/$false)
# ============================================================
$script:BaselineApps = @(
    [PSCustomObject]@{ Name = "Google Chrome";          ChocoID = "googlechrome";                Default = $true  }
    [PSCustomObject]@{ Name = "Mozilla Firefox";        ChocoID = "firefox";                     Default = $true  }
    [PSCustomObject]@{ Name = "Adobe Acrobat Reader";   ChocoID = "adobereader";                 Default = $true  }
    [PSCustomObject]@{ Name = "Zoom";                   ChocoID = "zoom";                        Default = $true  }
    [PSCustomObject]@{ Name = "7-Zip";                  ChocoID = "7zip";                        Default = $true  }
    [PSCustomObject]@{ Name = "Notepad++";              ChocoID = "notepadplusplus";             Default = $true  }
    [PSCustomObject]@{ Name = "ShareX";                 ChocoID = "sharex";                      Default = $true  }
    [PSCustomObject]@{ Name = "Everything";             ChocoID = "everything";                  Default = $true  }
    [PSCustomObject]@{ Name = "PowerToys";              ChocoID = "powertoys";                   Default = $false }
    [PSCustomObject]@{ Name = "One Commander";          ChocoID = "onecommander";                Default = $false }
    [PSCustomObject]@{ Name = "ImageGlass";             ChocoID = "imageglass";                  Default = $false }
    [PSCustomObject]@{ Name = "Teams (standalone)";     ChocoID = "microsoft-teams";             Default = $false }
)

# ============================================================
# HELPER: Detect Windows version
# ============================================================
function Get-WinVersion {
    $build = [System.Environment]::OSVersion.Version.Build
    if ($build -ge 22000) { return "Win11" } else { return "Win10" }
}

# ============================================================
# HELPER: Append text to the log RichTextBox, scroll to end
# ============================================================
function Write-Log {
    param([string]$Text, [System.Drawing.Color]$Color = [System.Drawing.Color]::LightGreen)
    if ($script:LogBox) {
        $script:LogBox.SelectionStart  = $script:LogBox.TextLength
        $script:LogBox.SelectionLength = 0
        $script:LogBox.SelectionColor  = $Color
        $script:LogBox.AppendText("$Text`n")
        $script:LogBox.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
    }
}

# ============================================================
# HELPER: Write Chocolatey output - progress lines overwrite,
#         all other lines append normally
# ============================================================
function Write-ChocoLog {
    param([string]$Text)
    if (-not $script:LogBox) { return }

    # Detect Chocolatey progress lines:
    # - Download progress:  "XX% ..."  or  "Progress: XX%"
    # - Download bar lines: lines containing sequences of # or = characters
    # - Extracting lines that repeat rapidly
    $isProgress = ($Text -match '^\s*\d+%') -or
                  ($Text -match 'Progress:') -or
                  ($Text -match 'Downloading\s+\d') -or
                  ($Text -match '[#=]{5}')

    if ($isProgress) {
        # Find the start of the last line and replace it in-place
        $script:LogBox.SelectionStart  = $script:LogBox.TextLength
        $script:LogBox.SelectionLength = 0

        $txt   = $script:LogBox.Text
        $lastN = $txt.LastIndexOf("`n")
        if ($lastN -ge 0) {
            # Select from after the last newline to end and replace
            $script:LogBox.SelectionStart  = $lastN + 1
            $script:LogBox.SelectionLength = $script:LogBox.TextLength - ($lastN + 1)
        } else {
            $script:LogBox.SelectionStart  = 0
            $script:LogBox.SelectionLength = $script:LogBox.TextLength
        }
        $script:LogBox.SelectionColor = [System.Drawing.ColorTranslator]::FromHtml("#6a9a50")
        $script:LogBox.SelectedText   = "      $($Text.Trim())"
        $script:LogBox.ScrollToCaret()
    } else {
        # Normal line - append with newline
        $script:LogBox.SelectionStart  = $script:LogBox.TextLength
        $script:LogBox.SelectionLength = 0
        $script:LogBox.SelectionColor  = [System.Drawing.Color]::LightGreen
        $script:LogBox.AppendText("      $Text`n")
        $script:LogBox.ScrollToCaret()
    }
    [System.Windows.Forms.Application]::DoEvents()
}

# ============================================================
# HELPER: Ensure a registry path exists
# ============================================================
function Ensure-RegPath([string]$Path) {
    if (!(Test-Path $Path)) { New-Item -Path $Path -Force | Out-Null }
}

# ============================================================
# SECTION: Set PC Name  (values supplied by the PC Naming tab)
# ============================================================
function Invoke-SetPCName {
    param([string]$DeviceType, [string]$Company, [string]$Location, [string]$AssetID)
    Write-Log ">>> Setting PC name..."
    if ($DeviceType -and $Company -and $Location -and $AssetID) {
        $newName = "$Company-$Location-$DeviceType-$AssetID"
        Write-Log "    Renaming computer to: $newName"
        Rename-Computer -NewName $newName -Force -ErrorAction SilentlyContinue
        Write-Log "    Done. Restart required for name change to apply."
    } else {
        Write-Log "    PC rename skipped - fill in all fields on the PC Naming tab first." ([System.Drawing.Color]::Yellow)
    }
}

# ============================================================
# SECTION: Set Time Zone
# ============================================================
function Invoke-SetTimeZone {
    param([string]$TimeZoneId = "Eastern Standard Time")
    Write-Log ">>> Setting time zone to: $TimeZoneId"
    Set-TimeZone -Id $TimeZoneId
    Write-Log "    Done."
}

# ============================================================
# SECTION: Install Chocolatey
# ============================================================
function Invoke-InstallChoco {
    Write-Log ">>> Checking Chocolatey..."
    if (Get-Command choco -ErrorAction SilentlyContinue) {
        Write-Log "    Chocolatey already installed."
        return
    }
    Write-Log "    Installing Chocolatey..."
    Set-ExecutionPolicy Bypass -Scope Process -Force
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
    Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
    Write-Log "    Chocolatey installed."
}

# ============================================================
# SECTION: Install Baseline Apps (called with selected IDs)
# ============================================================
function Invoke-InstallApps {
    param([string[]]$SelectedIDs)
    if ($SelectedIDs.Count -eq 0) {
        Write-Log "    No apps selected, skipping." ([System.Drawing.Color]::Yellow)
        return
    }
    Invoke-InstallChoco
    foreach ($id in $SelectedIDs) {
        Write-Log "    Installing: $id"
        if ($id -eq "microsoft-office-deployment") {
            $params = "'/Channel:Monthly /Language:en-us /RemoveMSI /Product:O365BusinessRetail /Exclude:Lync,Groove'"
            & choco install $id --params=$params -y 2>&1 | ForEach-Object { Write-ChocoLog "$_" }
        } else {
            & choco install $id -y 2>&1 | ForEach-Object { Write-ChocoLog "$_" }
        }
    }
    Write-Log "    App installation complete."
}

# ============================================================
# SECTION: Install Microsoft 365 (separate choco command)
# ============================================================
function Invoke-InstallM365 {
    Write-Log ">>> Installing Microsoft 365 / Office..."
    Invoke-InstallChoco
    & choco install microsoft-office-deployment --params="'/Channel:Monthly /Language:en-us /Product:O365BusinessRetail /Exclude:Lync,Groove'" -y 2>&1 | ForEach-Object { Write-ChocoLog "$_" }
    Write-Log "    Microsoft 365 install complete."
}

# ============================================================
# SECTION: Install custom Choco app by ID
# ============================================================
function Invoke-InstallCustomApp {
    param([string]$ChocoID)
    if ([string]::IsNullOrWhiteSpace($ChocoID)) { return }
    Invoke-InstallChoco
    Write-Log ">>> Installing custom package: $ChocoID"
    & choco install $ChocoID -y 2>&1 | ForEach-Object { Write-ChocoLog "$_" }
    Write-Log "    Done."
}

# ============================================================
# SECTION: Defender Exclusion (Utilities USB drive)
# ============================================================
function Invoke-AddDefenderExclusion {
    Write-Log ">>> Adding Windows Defender exclusion for tech tools drive..."
    $utilitiesVol = Get-Volume | Where-Object { $_.FileSystemLabel -eq "Utilities" } | Select-Object -First 1
    if (-not $utilitiesVol) {
        Write-Log "    Utilities drive not found (not connected?). Skipping exclusion." ([System.Drawing.Color]::Yellow)
        return
    }
    $excludePath = "$($utilitiesVol.DriveLetter):\_Collections\_TechToolStore"
    if (-not (Test-Path $excludePath)) {
        Write-Log "    Path not found on Utilities drive: $excludePath. Skipping." ([System.Drawing.Color]::Yellow)
        return
    }
    try {
        Add-MpPreference -ExclusionPath $excludePath -ErrorAction Stop
        Write-Log "    Defender exclusion added: $excludePath"
    } catch {
        Write-Log "    ERROR adding Defender exclusion: $_" ([System.Drawing.Color]::Red)
    }
}

# ============================================================
# SECTION: System Restore Point
# ============================================================
function Invoke-CreateRestorePoint {
    Write-Log ">>> Enabling System Restore and creating Initial Provisioning restore point..."
    try {
        # Enable System Restore on C: if not already enabled
        Enable-ComputerRestore -Drive "C:\" -ErrorAction Stop
        Write-Log "    System Restore enabled on C:."
    } catch {
        Write-Log "    Note: $($_.Exception.Message)" ([System.Drawing.Color]::Yellow)
    }
    try {
        # Remove the 24-hour frequency limit so we can create a point immediately
        $srPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore"
        $origFreq = (Get-ItemProperty -Path $srPath -Name "SystemRestorePointCreationFrequency" -ErrorAction SilentlyContinue).SystemRestorePointCreationFrequency
        Set-ItemProperty -Path $srPath -Name "SystemRestorePointCreationFrequency" -Value 0 -Type DWord -Force
        Checkpoint-Computer -Description "Initial Provisioning" -RestorePointType "MODIFY_SETTINGS" -ErrorAction Stop
        # Restore original frequency setting (or remove if it did not exist)
        if ($null -ne $origFreq) {
            Set-ItemProperty -Path $srPath -Name "SystemRestorePointCreationFrequency" -Value $origFreq -Type DWord -Force
        } else {
            Remove-ItemProperty -Path $srPath -Name "SystemRestorePointCreationFrequency" -ErrorAction SilentlyContinue
        }
        Write-Log "    Restore point created: Initial Provisioning"
    } catch {
        Write-Log "    ERROR creating restore point: $_" ([System.Drawing.Color]::Red)
    }
}

# ============================================================
# SECTION: TPM and Secure Boot status log
# ============================================================
function Invoke-LogSecurityHardwareStatus {
    Write-Log ">>> Documenting TPM and Secure Boot status..."
    # TPM
    try {
        $tpm = Get-Tpm -ErrorAction Stop
        Write-Log "    TPM Present:     $($tpm.TpmPresent)"
        Write-Log "    TPM Ready:       $($tpm.TpmReady)"
        Write-Log "    TPM Enabled:     $($tpm.TpmEnabled)"
        Write-Log "    TPM Activated:   $($tpm.TpmActivated)"
        Write-Log "    TPM Owned:       $($tpm.TpmOwned)"
        $tpmVer = (Get-WmiObject -Namespace "root\cimv2\security\microsofttpm" -Class "Win32_Tpm" -ErrorAction SilentlyContinue).SpecVersion
        if ($tpmVer) { Write-Log "    TPM Spec Version: $tpmVer" }
    } catch {
        Write-Log "    TPM: Could not retrieve status - $_" ([System.Drawing.Color]::Yellow)
    }
    # Secure Boot
    try {
        $sb = Confirm-SecureBootUEFI -ErrorAction Stop
        Write-Log "    Secure Boot:     $sb"
    } catch {
        Write-Log "    Secure Boot: Not supported or could not be confirmed - $_" ([System.Drawing.Color]::Yellow)
    }
    # UEFI vs BIOS
    try {
        $fwType = (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control" -Name "PEFirmwareType" -ErrorAction Stop).PEFirmwareType
        Write-Log "    Firmware Type:   $(if ($fwType -eq 2) { UEFI } else { Legacy BIOS })"
    } catch {
        Write-Log "    Firmware Type: Could not determine" ([System.Drawing.Color]::Yellow)
    }
}

# ============================================================
# SECTION: BitLocker Encryption
# ============================================================
function Invoke-EnableBitLocker {
    Write-Log ">>> Enabling BitLocker on C: drive..."

    # Find the CustData drive by volume label
    $custVol = Get-Volume | Where-Object { $_.FileSystemLabel -eq "CustData" } | Select-Object -First 1
    if (-not $custVol) {
        Write-Log "    CustData drive not found. Connect the USB drive and try again." ([System.Drawing.Color]::Red)
        Write-Log "    BitLocker will NOT be enabled without a confirmed key backup destination." ([System.Drawing.Color]::Red)
        return
    }
    $keyDrive  = "$($custVol.DriveLetter):"
    $keyFolder = "$keyDrive\BitLocker"
    Write-Log "    CustData drive found at $keyDrive. Keys will be saved to $keyFolder"

    # Ensure key backup folder exists
    if (-not (Test-Path $keyFolder)) {
        New-Item -Path $keyFolder -ItemType Directory -Force | Out-Null
    }

    $driveLetter = "C:"

    # Check if BitLocker is already on
    $blStatus = Get-BitLockerVolume -MountPoint $driveLetter -ErrorAction SilentlyContinue
    if ($blStatus -and $blStatus.ProtectionStatus -eq "On") {
        Write-Log "    BitLocker is already enabled on $driveLetter." ([System.Drawing.Color]::Yellow)
        Write-Log "    Backing up existing recovery key..."
    } else {
        try {
            # Enable BitLocker with TPM only (no PIN - appropriate for managed business endpoints)
            Enable-BitLocker -MountPoint $driveLetter -TpmProtector -ErrorAction Stop | Out-Null
            Write-Log "    BitLocker enabled on $driveLetter (TPM protector)."
        } catch {
            Write-Log "    ERROR enabling BitLocker: $_" ([System.Drawing.Color]::Red)
            Write-Log "    Verify TPM is present, enabled, and Secure Boot is active." ([System.Drawing.Color]::Yellow)
            return
        }
    }

    # Add a recovery password protector so we have a key to back up
    try {
        $existingRecovery = (Get-BitLockerVolume -MountPoint $driveLetter).KeyProtector |
            Where-Object { $_.KeyProtectorType -eq "RecoveryPassword" } | Select-Object -First 1
        if (-not $existingRecovery) {
            Add-BitLockerKeyProtector -MountPoint $driveLetter -RecoveryPasswordProtector -ErrorAction Stop | Out-Null
            $existingRecovery = (Get-BitLockerVolume -MountPoint $driveLetter).KeyProtector |
                Where-Object { $_.KeyProtectorType -eq "RecoveryPassword" } | Select-Object -First 1
        }

        if ($existingRecovery) {
            $recoveryPassword = $existingRecovery.RecoveryPassword
            $keyID            = $existingRecovery.KeyProtectorId.Trim("{}")
            $hostname         = $env:COMPUTERNAME

            # Filename matches the Windows default format, prefixed with hostname
            $keyFileName = "$hostname-BitLocker Recovery Key $keyID.TXT"
            $keyFilePath = Join-Path $keyFolder $keyFileName

            $keyContent = @"
BitLocker Recovery Key
======================
Computer Name  : $hostname
Drive          : $driveLetter
Date Generated : $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Key ID         : $keyID

Recovery Key:
$recoveryPassword
"@
            $keyContent | Out-File -FilePath $keyFilePath -Encoding UTF8 -Force
            Write-Log "    Recovery key saved: $keyFilePath"
            Write-Log "    Key ID: $keyID"
        } else {
            Write-Log "    WARNING: Could not retrieve recovery password protector." ([System.Drawing.Color]::Yellow)
        }
    } catch {
        Write-Log "    ERROR saving recovery key: $_" ([System.Drawing.Color]::Red)
    }

    Write-Log "    NOTE: BitLocker encryption runs in the background. Full encryption" ([System.Drawing.Color]::Cyan)
    Write-Log "    may take some time to complete after this script finishes." ([System.Drawing.Color]::Cyan)
}

# ============================================================
# SECTION: Power Profile
# ============================================================
function Invoke-SetPowerProfile {
    Write-Log ">>> Applying Pirum Power Management profile..."
    $schemeGUID  = "381b4222-f694-41f0-9685-ff5bb260aaaa"
    # Processor performance GUIDs
    $procSubGUID = "54533251-82be-4824-96c1-47b60b740d00"
    $minCpuGUID  = "893dee8e-2bef-41e0-89c6-b55d0929964c"  # min processor state
    $maxCpuGUID  = "bc5038f7-23e0-4960-96da-33abaf5935ec"  # max processor state
    $boostGUID   = "be337238-0d82-4146-a960-4f3749d470c7"  # processor perf boost

    POWERCFG -DUPLICATESCHEME 381b4222-f694-41f0-9685-ff5bb260df2e $schemeGUID 2>$null
    POWERCFG -CHANGENAME $schemeGUID "Pirum Power Management" 2>$null
    POWERCFG -SETACTIVE $schemeGUID

    # Sleep / timeout settings
    POWERCFG -Change -monitor-timeout-ac 30
    POWERCFG -Change -monitor-timeout-dc 10
    POWERCFG -Change -disk-timeout-ac 30
    POWERCFG -Change -disk-timeout-dc 5
    POWERCFG -Change -standby-timeout-ac 0
    POWERCFG -Change -standby-timeout-dc 30
    POWERCFG -Change -hibernate-timeout-ac 0
    POWERCFG -Change -hibernate-timeout-dc 0

    # High Performance CPU behavior:
    # Min processor state 100% on AC (never throttle down), 50% on DC
    # Max processor state 100% on both
    # Processor performance boost: Aggressive on AC
    POWERCFG -SETACVALUEINDEX $schemeGUID $procSubGUID $minCpuGUID 100
    POWERCFG -SETDCVALUEINDEX $schemeGUID $procSubGUID $minCpuGUID 50
    POWERCFG -SETACVALUEINDEX $schemeGUID $procSubGUID $maxCpuGUID 100
    POWERCFG -SETDCVALUEINDEX $schemeGUID $procSubGUID $maxCpuGUID 100
    POWERCFG -SETACVALUEINDEX $schemeGUID $procSubGUID $boostGUID 2   # 2 = Aggressive
    POWERCFG -SETDCVALUEINDEX $schemeGUID $procSubGUID $boostGUID 1   # 1 = Enabled

    Write-Log "    Power profile applied (AC: no sleep/hibernate, high-perf CPU; DC: sleep at 30 min)."
}

# ============================================================
# SECTION: Layout Design (new user profiles)
# ============================================================
function Invoke-LayoutDesign {
    # Delegates to the dedicated XML functions which use C:\Pirum\xml\
    Invoke-ApplyAppAssociations   -XmlPath "C:\Pirum\xml\AppAssociations.xml"
    Invoke-ApplyLayoutModification -XmlBasePath "C:\Pirum\xml"
}

# ============================================================
# SECTION: Apply App Associations XML
# ============================================================
function Invoke-ApplyAppAssociations {
    param([string]$XmlPath)
    Write-Log ">>> Applying default app associations..."
    if ([string]::IsNullOrWhiteSpace($XmlPath) -or -not (Test-Path $XmlPath)) {
        Write-Log "    AppAssociations.xml not found at: $XmlPath" ([System.Drawing.Color]::Yellow)
        return
    }
    try {
        dism /online /Import-DefaultAppAssociations:"$XmlPath" 2>&1 | Out-Null
        Write-Log "    App associations applied from: $XmlPath"
    } catch {
        Write-Log "    ERROR applying app associations: $_" ([System.Drawing.Color]::Red)
    }
}

# ============================================================
# SECTION: Apply Layout Modification XML
# ============================================================
function Invoke-ApplyLayoutModification {
    param([string]$XmlBasePath)
    Write-Log ">>> Applying layout modification (taskbar / Start menu)..."

    $osVer = Get-WinVersion

    # Select the OS-specific file, fall back to generic name if not found
    if ($osVer -eq "Win11") {
        $xmlFile = Join-Path $XmlBasePath "LayoutModification_Win11.xml"
    } else {
        $xmlFile = Join-Path $XmlBasePath "LayoutModification_Win10.xml"
    }
    if (-not (Test-Path $xmlFile)) {
        $xmlFile = Join-Path $XmlBasePath "LayoutModification.xml"
        Write-Log "    OS-specific file not found, falling back to LayoutModification.xml" ([System.Drawing.Color]::Yellow)
    }
    if (-not (Test-Path $xmlFile)) {
        Write-Log "    No LayoutModification XML found in: $XmlBasePath" ([System.Drawing.Color]::Yellow)
        return
    }

    Write-Log "    Using: $xmlFile"

    try {
        if ($osVer -eq "Win11") {
            # Win11: Import-StartLayout is not supported.
            # Copy to Default user profile - applies to every new user created on this machine.
            $destDir  = "C:\Users\Default\AppData\Local\Microsoft\Windows\Shell"
            $destFile = Join-Path $destDir "LayoutModification.xml"
            if (-not (Test-Path $destDir)) { New-Item -Path $destDir -ItemType Directory -Force | Out-Null }
            Copy-Item -Path $xmlFile -Destination $destFile -Force -ErrorAction Stop
            Write-Log "    Win11: XML copied to Default user profile."
            Write-Log "    Taskbar layout applies the next time a new user logs in." ([System.Drawing.Color]::Cyan)
        } else {
            # Win10: use Import-StartLayout
            $boot   = Get-Partition | Where-Object { $_.IsBoot -eq $true }
            $OSDISK = $boot.DriveLetter + ":"
            Import-StartLayout -LayoutPath $xmlFile -MountPath "$OSDISK\" -ErrorAction Stop
            Write-Log "    Win10: layout applied via Import-StartLayout."
            Write-Log "    Changes take effect for new user profiles only." ([System.Drawing.Color]::Cyan)
        }
    } catch {
        Write-Log "    ERROR applying layout: $_" ([System.Drawing.Color]::Red)
    }
}

# ============================================================
# SECTION: Join Domain
# ============================================================
function Invoke-JoinDomain {
    Write-Log ">>> Domain join..."
    Add-Type -AssemblyName Microsoft.VisualBasic
    $domain = [Microsoft.VisualBasic.Interaction]::InputBox('Enter domain name (e.g. contoso.local)', 'Domain Name')
    $ouPath = [Microsoft.VisualBasic.Interaction]::InputBox('Enter OU path (leave blank for default placement)', 'OU Path (optional)')
    if ([string]::IsNullOrWhiteSpace($domain)) {
        Write-Log "    Domain join cancelled." ([System.Drawing.Color]::Yellow)
        return
    }
    try {
        if ([string]::IsNullOrWhiteSpace($ouPath)) {
            Add-Computer -DomainName $domain -Credential (Get-Credential -Message "Domain credentials for $domain") -ErrorAction Stop
        } else {
            Add-Computer -DomainName $domain -OUPath $ouPath -Credential (Get-Credential -Message "Domain credentials for $domain") -ErrorAction Stop
        }
        Write-Log "    Joined $domain successfully. Restart required."
    } catch {
        Write-Log "    ERROR joining domain: $_" ([System.Drawing.Color]::Red)
    }
}

# ============================================================
# SECTION: Management Software
# ============================================================
function Get-InstallerExtension {
    # Determine the correct extension for a downloaded installer.
    # 1. Try the URL path (works when the URL ends in .msi or .exe)
    # 2. Fall back to reading the file magic bytes:
    #    MSI files begin with D0 CF 11 E0 (OLE Compound Document)
    #    EXE/PE files begin with 4D 5A ("MZ")
    param([string]$Url, [string]$DownloadedPath, [string]$FallbackExtension)
    # Try URL hint first
    $urlPath = ($Url -split "[?#]")[0]   # strip query string and fragment
    if ($urlPath -match '\.msi$') { return ".msi" }
    if ($urlPath -match '\.exe$') { return ".exe" }
    # Read magic bytes from downloaded file
    try {
        $bytes = [System.IO.File]::ReadAllBytes($DownloadedPath)
        if ($bytes.Length -ge 4 -and
            $bytes[0] -eq 0xD0 -and $bytes[1] -eq 0xCF -and
            $bytes[2] -eq 0x11 -and $bytes[3] -eq 0xE0) {
            return ".msi"
        }
        if ($bytes.Length -ge 2 -and $bytes[0] -eq 0x4D -and $bytes[1] -eq 0x5A) {
            return ".exe"
        }
    } catch {}
    return $FallbackExtension
}

function Resolve-Installer {
    # Given a URL or local path, returns a local file path ready to execute.
    # For URLs, auto-detects whether the download is an MSI or EXE so the
    # correct installer method is used regardless of what the URL looks like.
    param([string]$Source, [string]$Label, [string]$FallbackExtension = ".exe")
    if ([string]::IsNullOrWhiteSpace($Source)) {
        Write-Log "    $Label`: no path or URL configured. Skipping." ([System.Drawing.Color]::Yellow)
        return $null
    }
    if ($Source -match '^https?://') {
        Write-Log "    Downloading $Label from URL..."
        # Download to a neutral .tmp file first, then rename with correct extension
        $tmpBase = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "PirumAgent_$Label.tmp")
        try {
            [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
            Invoke-WebRequest -Uri $Source -OutFile $tmpBase -UseBasicParsing -ErrorAction Stop
            $ext     = Get-InstallerExtension -Url $Source -DownloadedPath $tmpBase -FallbackExtension $FallbackExtension
            $tmpFile = $tmpBase -replace "\.tmp$", $ext
            if ($tmpFile -ne $tmpBase) {
                if (Test-Path $tmpFile) { Remove-Item $tmpFile -Force }
                Rename-Item -Path $tmpBase -NewName ([System.IO.Path]::GetFileName($tmpFile)) -ErrorAction Stop
            }
            Write-Log "    Download complete: $tmpFile  (detected type: $ext)"
            return $tmpFile
        } catch {
            Write-Log "    ERROR downloading $Label`: $_" ([System.Drawing.Color]::Red)
            if (Test-Path $tmpBase) { Remove-Item $tmpBase -Force -ErrorAction SilentlyContinue }
            return $null
        }
    } else {
        if (Test-Path $Source) {
            return $Source
        } else {
            Write-Log "    $Label installer not found at: $Source" ([System.Drawing.Color]::Red)
            return $null
        }
    }
}

function Invoke-InstallManagementSoftware {
    param(
        [bool]$DoNinja,
        [bool]$DoAction1,
        [bool]$DoIHC,
        [string]$NinjaSource,
        [string]$Action1Source,
        [string]$IHCSource
    )

    if ($DoNinja) {
        Write-Log ">>> Installing NinjaOne RMM agent..."
        $installer = Resolve-Installer -Source $NinjaSource -Label "NinjaOne" -FallbackExtension ".msi"
        if ($installer) {
            Start-Process msiexec.exe -ArgumentList "/i `"$installer`" /quiet /norestart" -Wait -ErrorAction SilentlyContinue
            Write-Log "    NinjaOne agent install complete."
        }
    }

    if ($DoAction1) {
        Write-Log ">>> Installing Action1 agent..."
        $installer = Resolve-Installer -Source $Action1Source -Label "Action1" -FallbackExtension ".msi"
        if ($installer) {
            Start-Process msiexec.exe -ArgumentList "/i `"$installer`" /quiet /norestart" -Wait -ErrorAction SilentlyContinue
            Write-Log "    Action1 agent install complete."
        }
    }

    if ($DoIHC) {
        Write-Log ">>> Installing Instant Housecall..."
        $installer = Resolve-Installer -Source $IHCSource -Label "IHC" -FallbackExtension ".msi"
        if ($installer) {
            if ($installer -match "\.msi$") {
                Start-Process msiexec.exe -ArgumentList "/i `"$installer`" /quiet /norestart" -Wait -ErrorAction SilentlyContinue
            } else {
                Start-Process $installer -Wait -ErrorAction SilentlyContinue
            }
            Write-Log "    Instant Housecall install complete."
        }
    }
}

# ============================================================
# SECTION: Personalize
# ============================================================
function Invoke-Personalize {
    param(
        [bool]$DoOEM,
        [bool]$OEMSetManufacturer, [string]$OEMManufacturer,
        [bool]$OEMSetPhone,        [string]$OEMPhone,
        [bool]$OEMSetHours,        [string]$OEMHours,
        [bool]$OEMSetURL,          [string]$OEMURL,
        [bool]$DoWallpaper,   [string]$WallpaperSrc,
        [bool]$DoLockscreen,  [string]$LockscreenSrc,
        [bool]$DoUserPictures,
        [bool]$DoReset
    )

    $boot   = Get-Partition | Where-Object { $_.IsBoot -eq $true }
    $OSDISK = $boot.DriveLetter + ":"
    $SID    = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value

    if ($DoReset) {
        Write-Log ">>> Resetting personalization to Windows defaults..."
        $oemPath = "HKLM:\Software\Microsoft\Windows\CurrentVersion\OEMInformation"
        'Logo','Manufacturer','SupportPhone','SupportHours','SupportURL' | ForEach-Object {
            Remove-ItemProperty -Path $oemPath -Name $_ -ErrorAction SilentlyContinue
        }
        $bgPath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
        'Wallpaper','WallpaperStyle' | ForEach-Object {
            Remove-ItemProperty -Path $bgPath -Name $_ -ErrorAction SilentlyContinue
        }
        Remove-ItemProperty -Path "HKLM:\Software\Policies\Microsoft\Windows\Personalization" -Name "LockScreenImage" -ErrorAction SilentlyContinue
        Remove-Item -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP" -Recurse -Force -ErrorAction SilentlyContinue
        Write-Log "    Personalization reset complete."
        return
    }

    if ($DoOEM) {
        Write-Log ">>> Applying OEM branding..."
        $oemPath = "HKLM:\Software\Microsoft\Windows\CurrentVersion\OEMInformation"
        Ensure-RegPath $oemPath
        $OEMLogo = "OEMLogo.bmp"
        $logoSrc = "$OSDISK\Pirum\defmedia\$OEMLogo"
        if (Test-Path $logoSrc) {
            Copy-Item $logoSrc "$OSDISK\windows\system32" -Force
            Copy-Item $logoSrc "$OSDISK\windows\system32\oobe\info" -Force
            Set-ItemProperty -Path $oemPath -Name Logo -Value "$OSDISK\Windows\System32\$OEMLogo"
        }
        if ($OEMSetManufacturer) { Set-ItemProperty -Path $oemPath -Name Manufacturer -Value $OEMManufacturer }
        if ($OEMSetPhone)        { Set-ItemProperty -Path $oemPath -Name SupportPhone  -Value $OEMPhone }
        if ($OEMSetHours)        { Set-ItemProperty -Path $oemPath -Name SupportHours  -Value $OEMHours }
        if ($OEMSetURL)          { Set-ItemProperty -Path $oemPath -Name SupportURL    -Value $OEMURL }
        Write-Log "    OEM info applied."
    }

    if ($DoWallpaper) {
        Write-Log ">>> Applying desktop wallpaper..."
        if ([string]::IsNullOrWhiteSpace($WallpaperSrc) -or -not (Test-Path $WallpaperSrc)) {
            Write-Log "    Wallpaper source not found: $WallpaperSrc" ([System.Drawing.Color]::Yellow)
        } else {
            $wallFile = Split-Path $WallpaperSrc -Leaf
            $bgDir = "C:\windows\system32\oobe\info\backgrounds"
            if (-not (Test-Path $bgDir)) { New-Item $bgDir -ItemType Directory -Force | Out-Null }
            Copy-Item $WallpaperSrc $bgDir -Force
            Copy-Item $WallpaperSrc "$OSDISK\Windows\Web\Screen" -Force -ErrorAction SilentlyContinue
            Copy-Item $WallpaperSrc "$OSDISK\Windows\Web\Wallpaper\Windows" -Force -ErrorAction SilentlyContinue
            $bgPath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
            Ensure-RegPath $bgPath
            Set-ItemProperty -Path $bgPath -Name Wallpaper      -Value "$OSDISK\Windows\Web\Wallpaper\Windows\$wallFile"
            Set-ItemProperty -Path $bgPath -Name WallpaperStyle -Value "2"
            Write-Log "    Wallpaper applied: $wallFile"
        }
    }

    if ($DoLockscreen) {
        Write-Log ">>> Applying lock screen..."
        if ([string]::IsNullOrWhiteSpace($LockscreenSrc) -or -not (Test-Path $LockscreenSrc)) {
            Write-Log "    Lock screen source not found: $LockscreenSrc" ([System.Drawing.Color]::Yellow)
        } else {
            $lsFile = Split-Path $LockscreenSrc -Leaf
            Copy-Item $LockscreenSrc "$OSDISK\Windows\System32" -Force
            $sysPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System"
            Ensure-RegPath $sysPath
            Set-ItemProperty -Path $sysPath -Name UseOEMBackground -Value 1
            $perPath = "HKLM:\Software\Policies\Microsoft\Windows\Personalization"
            Ensure-RegPath $perPath
            Set-ItemProperty -Path $perPath -Name LockScreenImage -Value "$OSDISK\Windows\System32\$lsFile"
            $cspPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP"
            Ensure-RegPath $cspPath
            New-ItemProperty -Path $cspPath -Name LockScreenImageStatus -Value 1 -PropertyType DWORD  -Force | Out-Null
            New-ItemProperty -Path $cspPath -Name LockScreenImagePath   -Value "$OSDISK\Windows\System32\$lsFile" -PropertyType STRING -Force | Out-Null
            New-ItemProperty -Path $cspPath -Name LockScreenImageUrl    -Value "$OSDISK\Windows\System32\$lsFile" -PropertyType STRING -Force | Out-Null
            Write-Log "    Lock screen applied: $lsFile"
        }
    }

    if ($DoUserPictures) {
        Write-Log ">>> Applying user account pictures..."
        $userPicDest = "$OSDISK\ProgramData\Microsoft\User Account Pictures"
        # Prefer C:\Pirum\media; fall back to C:\Pirum\defmedia
        $mediaPath   = "$OSDISK\Pirum\media"
        $defmediaPath= "$OSDISK\Pirum\defmedia"
        $picSource   = if (Test-Path "$mediaPath\user-32.png" -or Test-Path "$mediaPath\user-32.bmp") { $mediaPath } else { $defmediaPath }
        Write-Log "    Using picture source: $picSource"
        foreach ($ext in @("png","bmp")) {
            $src = "$picSource\user.$ext"
            if (Test-Path $src) { Copy-Item $src $userPicDest -Force }
        }
        $regBase = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AccountPicture\Users"
        Ensure-RegPath "$regBase\$SID"
        $userRegPath = "$regBase\$SID"
        $applied = 0
        foreach ($size in @("32","40","48","96","192","240")) {
            foreach ($ext in @("bmp","png")) {
                $src = "$picSource\user-$size.$ext"
                if (Test-Path $src) {
                    Copy-Item $src $userPicDest -Force
                    Set-ItemProperty -Path $userRegPath -Name "Image$size" -Value "$userPicDest\user-$size.$ext" -ErrorAction SilentlyContinue
                    $applied++
                }
            }
        }
        if ($applied -gt 0) {
            Write-Log "    User account pictures applied ($applied size variants from $picSource)."
        } else {
            Write-Log "    No user picture files found in $mediaPath or $defmediaPath. Skipping." ([System.Drawing.Color]::Yellow)
        }
    }
}

# ============================================================
# SECTION: Reclaim Windows - run selected hardening items
# Each item is a hashtable describing a tweak.
# Category: "Privacy" | "SystemTweaks" | "UITweaks" | "Bloatware"
# Advisory: brief note shown as tooltip, explaining impact
# Win11Support: $true means the tweak has a Win11-specific variant or note
# ============================================================
function Get-ReclaimItems {
    $osVer = Get-WinVersion
    return @(

        # ======================================================
        # PRIVACY
        # ======================================================
        [PSCustomObject]@{
            Key      = "DisableTelemetry"
            Label    = "Disable Telemetry"
            Category = "Privacy"
            Default  = $true
            Advisory = "RECOMMENDED. Prevents Windows from sending diagnostic/usage data to Microsoft. Safe for all business workstations."
            Action   = {
                Ensure-RegPath "HKLM:\Software\Policies\Microsoft\Windows\DataCollection"
                Set-ItemProperty -Path "HKLM:\Software\Policies\Microsoft\Windows\DataCollection" -Name "AllowTelemetry" -Type DWord -Value 0
            }
        }
        [PSCustomObject]@{
            Key      = "DisableWifiSense"
            Label    = "Disable Wi-Fi Sense"
            Category = "Privacy"
            Default  = $true
            Advisory = "RECOMMENDED. Prevents Windows from auto-connecting to shared hotspots and reporting Wi-Fi network data to Microsoft."
            Action   = {
                $p = "HKLM:\Software\Microsoft\PolicyManager\default\WiFi"
                Ensure-RegPath "$p\AllowWiFiHotSpotReporting"
                Ensure-RegPath "$p\AllowAutoConnectToWiFiSenseHotspots"
                Set-ItemProperty -Path "$p\AllowWiFiHotSpotReporting"        -Name "Value" -Type DWord -Value 0
                Set-ItemProperty -Path "$p\AllowAutoConnectToWiFiSenseHotspots" -Name "Value" -Type DWord -Value 0
            }
        }
        [PSCustomObject]@{
            Key      = "DisableLocationTracking"
            Label    = "Disable Location Tracking"
            Category = "Privacy"
            Default  = $true
            Advisory = "RECOMMENDED for business workstations. Disables Windows location services. Apps will no longer be able to request location data."
            Action   = {
                $sensorPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Sensor\Overrides\{BFA794E4-F964-4FDB-90F6-51056BFE4B44}"
                Ensure-RegPath $sensorPath
                Set-ItemProperty -Path $sensorPath -Name "SensorPermissionState" -Type DWord -Value 0
                Ensure-RegPath "HKLM:\System\CurrentControlSet\Services\lfsvc\Service\Configuration"
                Set-ItemProperty -Path "HKLM:\System\CurrentControlSet\Services\lfsvc\Service\Configuration" -Name "Status" -Type DWord -Value 0
            }
        }
        [PSCustomObject]@{
            Key      = "DisableFeedback"
            Label    = "Disable Feedback Prompts"
            Category = "Privacy"
            Default  = $true
            Advisory = "RECOMMENDED. Stops Windows from prompting users for feedback surveys. No functional impact."
            Action   = {
                Ensure-RegPath "HKCU:\Software\Microsoft\Siuf\Rules"
                Set-ItemProperty -Path "HKCU:\Software\Microsoft\Siuf\Rules" -Name "NumberOfSIUFInPeriod" -Type DWord -Value 0
            }
        }
        [PSCustomObject]@{
            Key      = "DisableAdvertisingID"
            Label    = "Disable Advertising ID"
            Category = "Privacy"
            Default  = $true
            Advisory = "RECOMMENDED. Disables the per-user advertising ID used by apps to show targeted ads. No functional impact on business software."
            Action   = {
                Ensure-RegPath "HKCU:\Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo"
                Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo" -Name "Enabled" -Type DWord -Value 0
            }
        }
        [PSCustomObject]@{
            Key      = "DisableCortana"
            Label    = "Disable Cortana"
            Category = "Privacy"
            Default  = $true
            Advisory = "RECOMMENDED for business. Cortana collects significant personal/usage data. Win11 note: Copilot is separate and not disabled here. Has no negative impact on normal use."
            Action   = {
                Ensure-RegPath "HKCU:\Software\Microsoft\Personalization\Settings"
                Set-ItemProperty -Path "HKCU:\Software\Microsoft\Personalization\Settings" -Name "AcceptedPrivacyPolicy" -Type DWord -Value 0
                Ensure-RegPath "HKCU:\Software\Microsoft\InputPersonalization"
                Set-ItemProperty -Path "HKCU:\Software\Microsoft\InputPersonalization" -Name "RestrictImplicitTextCollection" -Type DWord -Value 1
                Set-ItemProperty -Path "HKCU:\Software\Microsoft\InputPersonalization" -Name "RestrictImplicitInkCollection"  -Type DWord -Value 1
                Ensure-RegPath "HKCU:\Software\Microsoft\InputPersonalization\TrainedDataStore"
                Set-ItemProperty -Path "HKCU:\Software\Microsoft\InputPersonalization\TrainedDataStore" -Name "HarvestContacts" -Type DWord -Value 0
                $searchKey = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search"
                Ensure-RegPath $searchKey
                Set-ItemProperty -Path $searchKey -Name "AllowCortana" -Type DWord -Value 0
            }
        }
        [PSCustomObject]@{
            Key      = "DisableBingSearch"
            Label    = "Disable Bing Web Search in Start Menu"
            Category = "Privacy"
            Default  = $true
            Advisory = "RECOMMENDED for business. Prevents start menu searches from sending queries to Bing. Keeps local searches local. No productivity impact."
            Action   = {
                Ensure-RegPath "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search"
                Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search" -Name "BingSearchEnabled"  -Type DWord -Value 0
                Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search" -Name "CortanaConsent"      -Type DWord -Value 0
            }
        }
        [PSCustomObject]@{
            Key      = "RestrictWUP2P"
            Label    = "Restrict Windows Update P2P to Local Network Only"
            Category = "Privacy"
            Default  = $true
            Advisory = "RECOMMENDED. By default Windows can use your machine to distribute updates to random internet peers. This limits delivery optimization to LAN only, reducing bandwidth usage."
            Action   = {
                Ensure-RegPath "HKLM:\Software\Microsoft\Windows\CurrentVersion\DeliveryOptimization\Config"
                Set-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\DeliveryOptimization\Config" -Name "DODownloadMode" -Type DWord -Value 1
                Ensure-RegPath "HKCU:\Software\Microsoft\Windows\CurrentVersion\DeliveryOptimization"
                Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\DeliveryOptimization" -Name "SystemSettingsDownloadMode" -Type DWord -Value 3
            }
        }
        [PSCustomObject]@{
            Key      = "DisableDiagTrack"
            Label    = "Disable Diagnostics Tracking Service (DiagTrack)"
            Category = "Privacy"
            Default  = $true
            Advisory = "RECOMMENDED. Stops and permanently disables the Connected User Experiences and Telemetry service. Also disables WAP Push service and removes the AutoLogger ETL file. No end-user impact."
            Action   = {
                Stop-Service "DiagTrack"      -ErrorAction SilentlyContinue
                Stop-Service "dmwappushservice" -ErrorAction SilentlyContinue
                Set-Service  "DiagTrack"      -StartupType Disabled -ErrorAction SilentlyContinue
                Set-Service  "dmwappushservice" -StartupType Disabled -ErrorAction SilentlyContinue
                $autoLogDir = "$env:PROGRAMDATA\Microsoft\Diagnosis\ETLLogs\AutoLogger"
                $etlFile    = "$autoLogDir\AutoLogger-Diagtrack-Listener.etl"
                if (Test-Path $etlFile) { Remove-Item $etlFile -Force -ErrorAction SilentlyContinue }
                icacls $autoLogDir /deny SYSTEM:`(OI`)`(CI`)F 2>$null | Out-Null
            }
        }
        [PSCustomObject]@{
            Key      = "PreventBloatwareReturn"
            Label    = "Prevent Bloatware Auto-Reinstall (CloudContent Policy)"
            Category = "Privacy"
            Default  = $true
            Advisory = "RECOMMENDED. Sets a policy key that prevents Windows from silently reinstalling sponsored/suggested apps and prevents pre-installed app recommendations from re-appearing after removal."
            Action   = {
                $regPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent"
                Ensure-RegPath $regPath
                Set-ItemProperty -Path $regPath -Name "DisableWindowsConsumerFeatures" -Type DWord -Value 1
                $Holo = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Holographic"
                if (Test-Path $Holo) { Set-ItemProperty $Holo FirstRunSucceeded -Value 0 }
                reg load HKU\Default_User C:\Users\Default\NTUSER.DAT 2>$null
                $cdm = "Registry::HKU\Default_User\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"
                Set-ItemProperty -Path $cdm -Name "SystemPaneSuggestionsEnabled"  -Value 0 -ErrorAction SilentlyContinue
                Set-ItemProperty -Path $cdm -Name "PreInstalledAppsEnabled"        -Value 0 -ErrorAction SilentlyContinue
                Set-ItemProperty -Path $cdm -Name "OemPreInstalledAppsEnabled"     -Value 0 -ErrorAction SilentlyContinue
                reg unload HKU\Default_User 2>$null
            }
        }

        # ======================================================
        # SYSTEM TWEAKS
        # ======================================================
        [PSCustomObject]@{
            Key      = "EnableFirewall"
            Label    = "Ensure Windows Firewall is Enabled (All Profiles)"
            Category = "SystemTweaks"
            Default  = $true
            Advisory = "STRONGLY RECOMMENDED. Enables the Windows Firewall on Domain, Private, and Public profiles. Should always be on in a managed environment."
            Action   = { Set-NetFirewallProfile -Profile * -Enabled True }
        }
        [PSCustomObject]@{
            Key      = "EnsureDefenderEnabled"
            Label    = "Ensure Windows Defender is Active"
            Category = "SystemTweaks"
            Default  = $true
            Advisory = "STRONGLY RECOMMENDED. Removes the DisableAntiSpyware policy key if present, ensuring Defender stays active. Does not interfere with third-party AV co-existence."
            Action   = {
                $defPath = "HKLM:\Software\Policies\Microsoft\Windows Defender"
                if (Test-Path $defPath) {
                    Remove-ItemProperty -Path $defPath -Name "DisableAntiSpyware" -ErrorAction SilentlyContinue
                }
            }
        }
        [PSCustomObject]@{
            Key      = "DisableRDP"
            Label    = "Disable Remote Desktop (RDP)"
            Category = "SystemTweaks"
            Default  = $true
            Advisory = "RECOMMENDED. RDP is a significant attack vector. Disable unless the client specifically needs inbound RDP. NinjaOne and Instant Housecall handle remote support needs."
            Action   = {
                Set-ItemProperty -Path "HKLM:\System\CurrentControlSet\Control\Terminal Server" -Name "fDenyTSConnections" -Type DWord -Value 1
                Set-ItemProperty -Path "HKLM:\System\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp" -Name "UserAuthentication" -Type DWord -Value 1
            }
        }
        [PSCustomObject]@{
            Key      = "EnableUAC"
            Label    = "Ensure UAC is Set to Prompt (Not Silently Elevate)"
            Category = "SystemTweaks"
            Default  = $true
            Advisory = "RECOMMENDED. Sets UAC to prompt admin users on the secure desktop. Prevents malware from silently elevating without user awareness."
            Action   = {
                Set-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System" -Name "ConsentPromptBehaviorAdmin" -Type DWord -Value 5
                Set-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System" -Name "PromptOnSecureDesktop"       -Type DWord -Value 1
            }
        }
        [PSCustomObject]@{
            Key      = "DisableWUAutoRestart"
            Label    = "Disable Forced Windows Update Auto-Restart"
            Category = "SystemTweaks"
            Default  = $true
            Advisory = "RECOMMENDED for managed environments. Prevents Windows Update from force-rebooting outside business hours. Restart scheduling should be handled through your RMM."
            Action   = {
                Ensure-RegPath "HKLM:\Software\Microsoft\WindowsUpdate\UX\Settings"
                Set-ItemProperty -Path "HKLM:\Software\Microsoft\WindowsUpdate\UX\Settings" -Name "UxOption" -Type DWord -Value 1
            }
        }
        [PSCustomObject]@{
            Key      = "DisableScheduledTelemetryTasks"
            Label    = "Disable Unnecessary Telemetry Scheduled Tasks"
            Category = "SystemTweaks"
            Default  = $true
            Advisory = "RECOMMENDED. Disables Xbox game-save telemetry, CEIP, and device management client tasks. These serve no purpose on business machines."
            Action   = {
                @("XblGameSaveTask","Consolidator","UsbCeip","DmClient","DmClientOnScenarioDownload") | ForEach-Object {
                    Get-ScheduledTask -TaskName $_ -ErrorAction SilentlyContinue | Disable-ScheduledTask -ErrorAction SilentlyContinue
                }
            }
        }
        [PSCustomObject]@{
            Key      = "InstallDotNet35"
            Label    = "Install .NET Framework 3.5"
            Category = "SystemTweaks"
            Default  = $true
            Advisory = "RECOMMENDED. Many business and LOB applications still depend on .NET 3.5. Installs via DISM from Windows Update. Safe to run on both Win10 and Win11."
            Action   = { DISM /Online /Enable-Feature /FeatureName:NetFx3 /All /Quiet /NoRestart }
        }
        [PSCustomObject]@{
            Key      = "EnableVSS"
            Label    = "Enable Volume Shadow Copy Service (VSS)"
            Category = "SystemTweaks"
            Default  = $true
            Advisory = "RECOMMENDED. VSS is required for Windows backup, File History, and most third-party backup agents including NinjaOne's backup integration. Ensure it is set to Manual (triggered on demand)."
            Action   = {
                Set-Service -Name "VSS"   -StartupType Manual -ErrorAction SilentlyContinue
                Set-Service -Name "swprv" -StartupType Manual -ErrorAction SilentlyContinue
            }
        }
        [PSCustomObject]@{
            Key      = "SetEasternTime"
            Label    = "Set Time Zone to Eastern Standard Time"
            Category = "SystemTweaks"
            Default  = $false
            Advisory = "OPTIONAL here - also available as a top-level step. Only check this if you want time zone set as part of the Reclaim Windows pass."
            Action   = { Set-TimeZone -Id "Eastern Standard Time" }
        }
        [PSCustomObject]@{
            Key      = "DisableWin11Copilot"
            Label    = "Disable Windows Copilot Button (Win11 only)"
            Category = "SystemTweaks"
            Default  = $false
            Advisory = "OPTIONAL on Win11. Removes the Copilot sidebar button from the taskbar. Appropriate for most managed business endpoints. No effect on Win10."
            Action   = {
                if ((Get-WinVersion) -eq "Win11") {
                    Ensure-RegPath "HKCU:\Software\Policies\Microsoft\Windows\WindowsCopilot"
                    Set-ItemProperty -Path "HKCU:\Software\Policies\Microsoft\Windows\WindowsCopilot" -Name "TurnOffWindowsCopilot" -Type DWord -Value 1
                }
            }
        }
        [PSCustomObject]@{
            Key      = "EnableWakeOnLAN"
            Label    = "Enable Wake-on-LAN (WoL)"
            Category = "SystemTweaks"
            Default  = $true
            Advisory = "RECOMMENDED for managed endpoints. Enables WoL on all capable ethernet adapters via the Windows driver settings. IMPORTANT: WoL must also be enabled in the machine BIOS/UEFI firmware - this script handles the OS side only. Most Dell/HP/Lenovo machines ship with WoL on in BIOS by default."
            Action   = {
                $adapters = Get-NetAdapter -Physical | Where-Object { $_.Status -eq "Up" -or $_.Status -eq "Disconnected" }
                $enabled  = 0
                foreach ($adapter in $adapters) {
                    try {
                        $adapterPower = Get-WmiObject MSPower_DeviceWakeEnable -Namespace root\wmi -ErrorAction SilentlyContinue |
                            Where-Object { $_.InstanceName -like "*$($adapter.InterfaceDescription)*" }
                        if ($adapterPower) {
                            $adapterPower.Enable = $true
                            $adapterPower.Put() | Out-Null
                        }
                        $pnpDev = Get-PnpDevice -FriendlyName "*$($adapter.InterfaceDescription)*" -ErrorAction SilentlyContinue
                        if ($pnpDev) {
                            $nicRegBase = "HKLM:\SYSTEM\CurrentControlSet\Control\Class\{4D36E972-E325-11CE-BFC1-08002BE10318}"
                            Get-ChildItem $nicRegBase -ErrorAction SilentlyContinue | ForEach-Object {
                                Set-ItemProperty -Path $_.PSPath -Name "WakeOnMagicPacket"  -Value 1 -ErrorAction SilentlyContinue
                                Set-ItemProperty -Path $_.PSPath -Name "*WakeOnMagicPacket" -Value 1 -ErrorAction SilentlyContinue
                                Set-ItemProperty -Path $_.PSPath -Name "WakeOnPattern"      -Value 1 -ErrorAction SilentlyContinue
                            }
                        }
                        $devMgr = Get-WmiObject Win32_NetworkAdapter | Where-Object { $_.NetConnectionID -eq $adapter.Name }
                        if ($devMgr) { $devMgr.SetPowerState(1, $null) | Out-Null }
                        $enabled++
                    } catch {
                        # Non-fatal: some adapters do not support WoL
                    }
                }
                Write-Host "WoL configured on $enabled adapter(s). Verify BIOS/UEFI WoL setting manually."
            }
        }

        [PSCustomObject]@{
            Key      = "DisableTransparencyAnimations"
            Label    = "Disable Transparency and Animations"
            Category = "SystemTweaks"
            Default  = $true
            Advisory = "RECOMMENDED for performance. Disables window transparency, taskbar transparency, and UI animations. Visually simpler and noticeably faster on lower-spec machines."
            Action   = {
                # Transparency
                Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "EnableTransparency" -Type DWord -Value 0 -ErrorAction SilentlyContinue
                # Animations and visual effects - VisualFXSetting 3 = custom, 2 = best performance
                Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects" -Name "VisualFXSetting" -Type DWord -Value 3 -ErrorAction SilentlyContinue
                # Selectively disable animations while keeping useful rendering
                $apPath = "HKCU:\Control Panel\Desktop\WindowMetrics"
                Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "UserPreferencesMask" -Type Binary -Value ([byte[]](0x90,0x12,0x03,0x80,0x10,0x00,0x00,0x00)) -ErrorAction SilentlyContinue
                Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "MenuShowDelay" -Type String -Value "0" -ErrorAction SilentlyContinue
                Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "ListviewShadow" -Type DWord -Value 0 -ErrorAction SilentlyContinue
                Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "TaskbarAnimations" -Type DWord -Value 0 -ErrorAction SilentlyContinue
                Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\DWM" -Name "EnableAeroPeek" -Type DWord -Value 0 -ErrorAction SilentlyContinue
            }
        }
        [PSCustomObject]@{
            Key      = "EnableCoreIsolation"
            Label    = "Enable Core Isolation / Memory Integrity (HVCI)"
            Category = "SystemTweaks"
            Default  = $true
            Advisory = "RECOMMENDED for business endpoints. Enables Hypervisor-Protected Code Integrity (HVCI/Memory Integrity). Requires TPM 2.0 and Secure Boot. A restart is required before it takes effect. May block incompatible drivers - check Device Security after reboot."
            Action   = {
                $hvciPath = "HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard\Scenarios\HypervisorEnforcedCodeIntegrity"
                Ensure-RegPath $hvciPath
                Set-ItemProperty -Path $hvciPath -Name "Enabled" -Type DWord -Value 1
                Set-ItemProperty -Path $hvciPath -Name "Locked"  -Type DWord -Value 0
                # Credential Guard companion key
                $dgPath = "HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard"
                Ensure-RegPath $dgPath
                Set-ItemProperty -Path $dgPath -Name "EnableVirtualizationBasedSecurity" -Type DWord -Value 1
                Set-ItemProperty -Path $dgPath -Name "RequirePlatformSecurityFeatures"   -Type DWord -Value 1
                Write-Host "Core Isolation enabled. Restart required for Memory Integrity to take effect."
            }
        }
        [PSCustomObject]@{
            Key      = "DisableSMBv1"
            Label    = "Disable SMBv1 Protocol"
            Category = "SystemTweaks"
            Default  = $true
            Advisory = "STRONGLY RECOMMENDED. SMBv1 is the attack vector for WannaCry and similar ransomware. Should never be enabled on a managed business endpoint. Disabling it has no impact on modern networks."
            Action   = {
                Set-SmbServerConfiguration -EnableSMB1Protocol $false -Force -ErrorAction SilentlyContinue
                Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters" -Name "SMB1" -Type DWord -Value 0 -ErrorAction SilentlyContinue
                Disable-WindowsOptionalFeature -Online -FeatureName "SMB1Protocol" -NoRestart -ErrorAction SilentlyContinue | Out-Null
            }
        }
        [PSCustomObject]@{
            Key      = "EnableSMBSigning"
            Label    = "Enable SMB Signing (Require Security Signature)"
            Category = "SystemTweaks"
            Default  = $true
            Advisory = "RECOMMENDED. Prevents man-in-the-middle attacks on SMB file share connections by requiring cryptographic signing on all SMB traffic. Low risk to enable; high value for security."
            Action   = {
                Set-SmbClientConfiguration -RequireSecuritySignature $true  -Force -ErrorAction SilentlyContinue
                Set-SmbServerConfiguration -RequireSecuritySignature $false -Force -ErrorAction SilentlyContinue  # server side: require on client, enable but not mandate on server
                Set-SmbServerConfiguration -EnableSecuritySignature  $true  -Force -ErrorAction SilentlyContinue
            }
        }

        [PSCustomObject]@{
            Key      = "SetDNSServers"
            Label    = "Set DNS Servers (configure selection on Reclaim tab)"
            Category = "SystemTweaks"
            Default  = $true
            Advisory = "RECOMMENDED. Sets explicit DNS servers on all active ethernet adapters. Choice of provider configured on the Reclaim Windows tab. Prevents ISP DNS from being the default."
            Action   = {
                $dnsServers = if ($script:SelectedDNS) { $script:SelectedDNS } else { @("1.1.1.1","1.0.0.1") }
                $adapters = Get-NetAdapter -Physical | Where-Object { $_.Status -eq "Up" }
                foreach ($adapter in $adapters) {
                    try {
                        Set-DnsClientServerAddress -InterfaceIndex $adapter.InterfaceIndex -ServerAddresses $dnsServers -ErrorAction Stop
                        Write-Host "DNS set on $($adapter.Name): $($dnsServers -join ', ')"
                    } catch {
                        Write-Host "Could not set DNS on $($adapter.Name): $_"
                    }
                }
            }
        }

        # ======================================================
        # UI TWEAKS
        # ======================================================
        [PSCustomObject]@{
            Key      = "DisableActionCenter"
            Label    = "Disable Action Center / Notification Center"
            Category = "UITweaks"
            Default  = $false
            Advisory = "OPTIONAL. Completely removes the Notification Center from the taskbar. Can frustrate users who rely on app notifications. Consider leaving enabled unless client requests it off."
            Action   = {
                Ensure-RegPath "HKCU:\Software\Policies\Microsoft\Windows\Explorer"
                Set-ItemProperty -Path "HKCU:\Software\Policies\Microsoft\Windows\Explorer" -Name "DisableNotificationCenter" -Type DWord -Value 1
                Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\PushNotifications" -Name "ToastEnabled" -Type DWord -Value 0
            }
        }
        [PSCustomObject]@{
            Key      = "DisableAutoplay"
            Label    = "Disable AutoPlay and AutoRun (All Drives)"
            Category = "UITweaks"
            Default  = $true
            Advisory = "RECOMMENDED for security. AutoRun is a common malware vector via USB drives. Disabling both AutoPlay and AutoRun reduces risk with minimal usability impact."
            Action   = {
                Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\AutoplayHandlers" -Name "DisableAutoplay" -Type DWord -Value 1
                Ensure-RegPath "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
                Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer" -Name "NoDriveTypeAutoRun" -Type DWord -Value 255
            }
        }
        [PSCustomObject]@{
            Key      = "DisableStickyKeys"
            Label    = "Disable Sticky Keys Prompt"
            Category = "UITweaks"
            Default  = $true
            Advisory = "RECOMMENDED. Prevents the Sticky Keys accessibility dialog from appearing when Shift is pressed multiple times. Annoying for most users and rarely needed."
            Action   = {
                Set-ItemProperty -Path "HKCU:\Control Panel\Accessibility\StickyKeys" -Name "Flags" -Type String -Value "506"
            }
        }
        [PSCustomObject]@{
            Key      = "ShowFileExtensions"
            Label    = "Show Known File Extensions"
            Category = "UITweaks"
            Default  = $true
            Advisory = "RECOMMENDED for security. Hiding extensions makes it easier for malicious files to disguise themselves (e.g. 'invoice.pdf.exe' shows as 'invoice.pdf'). Essential for user awareness."
            Action   = {
                Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "HideFileExt" -Type DWord -Value 0
            }
        }
        [PSCustomObject]@{
            Key      = "ExplorerOpenThisPC"
            Label    = "Set File Explorer Default View to 'This PC'"
            Category = "UITweaks"
            Default  = $true
            Advisory = "OPTIONAL. Opens File Explorer to 'This PC' (drives) instead of Quick Access. Preferred by many business users. Personal preference item."
            Action   = {
                Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "LaunchTo" -Type DWord -Value 1
            }
        }
        [PSCustomObject]@{
            Key      = "ShowHiddenFiles"
            Label    = "Show Hidden Files and Folders"
            Category = "UITweaks"
            Default  = $false
            Advisory = "OPTIONAL. Reveals hidden system files in Explorer. Useful for technicians but can cause confusion for end users. Leave unchecked for standard user machines."
            Action   = {
                Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "Hidden" -Type DWord -Value 1
            }
        }
        [PSCustomObject]@{
            Key      = "ClassicPhotoViewer"
            Label    = "Restore Classic Windows Photo Viewer as Default"
            Category = "UITweaks"
            Default  = $false
            Advisory = "OPTIONAL. Restores the faster legacy Photo Viewer instead of the Photos app for JPG/PNG/BMP/GIF. Works on Win10. On Win11 the Photos app is deeply integrated; this still works but may be overridden."
            Action   = {
                if (!(Test-Path "HKCR:")) { New-PSDrive -Name HKCR -PSProvider Registry -Root HKEY_CLASSES_ROOT | Out-Null }
                foreach ($type in @("Paint.Picture","giffile","jpegfile","pngfile")) {
                    New-Item -Path "HKCR:\$type\shell\open"         -Force | Out-Null
                    New-Item -Path "HKCR:\$type\shell\open\command" -Force | Out-Null
                    Set-ItemProperty -Path "HKCR:\$type\shell\open" -Name "MuiVerb" -Type ExpandString -Value "@%ProgramFiles%\Windows Photo Viewer\photoviewer.dll,-3043"
                    Set-ItemProperty -Path "HKCR:\$type\shell\open\command" -Name "(Default)" -Type ExpandString -Value "%SystemRoot%\System32\rundll32.exe `"%ProgramFiles%\Windows Photo Viewer\PhotoViewer.dll`", ImageView_Fullscreen %1"
                }
                New-Item -Path "HKCR:\Applications\photoviewer.dll\shell\open\command"    -Force | Out-Null
                New-Item -Path "HKCR:\Applications\photoviewer.dll\shell\open\DropTarget" -Force | Out-Null
                Set-ItemProperty -Path "HKCR:\Applications\photoviewer.dll\shell\open" -Name "MuiVerb" -Type String -Value "@photoviewer.dll,-3043"
                Set-ItemProperty -Path "HKCR:\Applications\photoviewer.dll\shell\open\command" -Name "(Default)" -Type ExpandString -Value "%SystemRoot%\System32\rundll32.exe `"%ProgramFiles%\Windows Photo Viewer\PhotoViewer.dll`", ImageView_Fullscreen %1"
                Set-ItemProperty -Path "HKCR:\Applications\photoviewer.dll\shell\open\DropTarget" -Name "Clsid" -Type String -Value "{FFE2A43C-56B9-4bf5-9A79-CC6D4285608A}"
            }
        }
        [PSCustomObject]@{
            Key      = "Win11LeftTaskbar"
            Label    = "Move Taskbar Icons to Left (Win11 only)"
            Category = "UITweaks"
            Default  = ($osVer -eq "Win11")
            Advisory = "Win11 only. Moves the Start button and taskbar icons to the left edge to match the traditional Windows 10 layout. Helps users transitioning from Win10."
            Action   = {
                if ((Get-WinVersion) -eq "Win11") {
                    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "TaskbarAl" -Type DWord -Value 0
                }
            }
        }
        [PSCustomObject]@{
            Key      = "Win11DisableWidgets"
            Label    = "Disable Taskbar Widgets (Win11 only)"
            Category = "UITweaks"
            Default  = ($osVer -eq "Win11")
            Advisory = "Win11 only. Removes the Widgets news/weather panel button from the taskbar. The Widgets panel transmits usage data and is typically not appropriate for business workstations."
            Action   = {
                if ((Get-WinVersion) -eq "Win11") {
                    Ensure-RegPath "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
                    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "TaskbarDa" -Type DWord -Value 0
                    Ensure-RegPath "HKLM:\SOFTWARE\Policies\Microsoft\Dsh"
                    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Dsh" -Name "AllowNewsAndInterests" -Type DWord -Value 0
                }
            }
        }
        [PSCustomObject]@{
            Key      = "Win11DisableChat"
            Label    = "Disable Taskbar Chat / Teams Button (Win11 only)"
            Category = "UITweaks"
            Default  = ($osVer -eq "Win11")
            Advisory = "Win11 only. Removes the built-in Teams Chat button from the taskbar. Most business environments deploy Teams via M365; the built-in consumer version just causes confusion."
            Action   = {
                if ((Get-WinVersion) -eq "Win11") {
                    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "TaskbarMn" -Type DWord -Value 0 -ErrorAction SilentlyContinue
                }
            }
        }

        # ======================================================
        # BLOATWARE REMOVAL
        # ======================================================
        [PSCustomObject]@{
            Key      = "RemoveBloatware"
            Label    = "Remove Pre-installed Bloatware Apps"
            Category = "Bloatware"
            Default  = $true
            Advisory = "RECOMMENDED. Removes Xbox apps, consumer games, Bing apps, and sponsored third-party apps that serve no business purpose. Does NOT remove core productivity apps like Camera or Calculator."
            Action   = {
                $BloatList = @(
                    "Microsoft.BingNews","Microsoft.GetHelp","Microsoft.Getstarted",
                    "Microsoft.Microsoft3DViewer","Microsoft.MicrosoftOfficeHub",
                    "Microsoft.MicrosoftSolitaireCollection","Microsoft.NetworkSpeedTest",
                    "Microsoft.News","Microsoft.OneConnect","Microsoft.People",
                    "Microsoft.Print3D","Microsoft.RemoteDesktop","Microsoft.StorePurchaseApp",
                    "Microsoft.Office.Todo.List","Microsoft.Whiteboard","Microsoft.WindowsAlarms",
                    "Microsoft.WindowsFeedbackHub","Microsoft.WindowsMaps","Microsoft.Xbox.TCUI",
                    "Microsoft.XboxApp","Microsoft.XboxGameOverlay","Microsoft.XboxIdentityProvider",
                    "Microsoft.XboxSpeechToTextOverlay","Microsoft.ZuneMusic","Microsoft.ZuneVideo",
                    "Microsoft.3DBuilder","Microsoft.WindowsPhone","Microsoft.BingWeather",
                    "MicrosoftTeams",
                    "*EclipseManager*","*ActiproSoftwareLLC*","*AdobeSystemsIncorporated.AdobePhotoshopExpress*",
                    "*Duolingo-LearnLanguagesforFree*","*PandoraMediaInc*","*CandyCrush*",
                    "*Wunderlist*","*Flipboard*","*Twitter*","*Facebook*","*Spotify*",
                    "*Minecraft*","*Royal Revolt*","*Sway*","*Dolby*","*Windows.CBSPreview*"
                )
                foreach ($bloat in $BloatList) {
                    Get-AppxPackage -Name $bloat -ErrorAction SilentlyContinue | Remove-AppxPackage -ErrorAction SilentlyContinue
                    Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue |
                        Where-Object DisplayName -like $bloat |
                        Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue
                }
            }
        }
        [PSCustomObject]@{
            Key      = "DisableOneDrive"
            Label    = "Disable OneDrive (Policy Block)"
            Category = "Bloatware"
            Default  = $false
            Advisory = "OPTIONAL. Prevents OneDrive from running or syncing. Use only for clients who are NOT using OneDrive/SharePoint for file storage. If they use M365 with OneDrive, leave this unchecked."
            Action   = {
                Ensure-RegPath "HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive"
                Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive" -Name "DisableFileSyncNGSC" -Type DWord -Value 1
            }
        }
    )
}

# ============================================================
# MAIN GUI
# ============================================================
function Show-MainForm {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $osVer        = Get-WinVersion
    $reclaimItems = Get-ReclaimItems

    # Brand colors
    $clrPurple  = [System.Drawing.ColorTranslator]::FromHtml("#582a72")
    $clrGreen   = [System.Drawing.ColorTranslator]::FromHtml("#44722a")
    $clrGold    = [System.Drawing.ColorTranslator]::FromHtml("#a99639")
    $clrLav     = [System.Drawing.ColorTranslator]::FromHtml("#76508c")
    $clrDarkBg  = [System.Drawing.ColorTranslator]::FromHtml("#1e1028")
    $clrPaneBg  = [System.Drawing.ColorTranslator]::FromHtml("#f7f4fa")
    $clrWhite   = [System.Drawing.Color]::White
    $clrGray    = [System.Drawing.ColorTranslator]::FromHtml("#444444")
    $clrWarn    = [System.Drawing.ColorTranslator]::FromHtml("#cc6600")

    $segUI  = New-Object System.Drawing.Font("Segoe UI", 9)
    $segB   = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $segSm  = New-Object System.Drawing.Font("Segoe UI", 8)
    $segHdr = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)

    $ttip = New-Object System.Windows.Forms.ToolTip
    $ttip.AutoPopDelay = 9000
    $ttip.InitialDelay = 400
    $ttip.ReshowDelay  = 200
    $ttip.ShowAlways   = $true

    # ---- Main window ----
    # ClientSize defines the inner canvas exactly. No Dock on any major panel.
    # Header: static at (0,0), 62px tall.
    # Tabs:   static at (0,62), anchored all four sides.
    # Footer: anchored Bottom+Left+Right, floats at the bottom edge on resize.
    $form                  = New-Object System.Windows.Forms.Form
    $form.Text             = "Pirum Consulting LLC  -  PC Setup Tool  ($osVer detected)"
    $form.ClientSize       = New-Object System.Drawing.Size(1060, 730)
    $form.MinimumSize      = New-Object System.Drawing.Size(900, 600)
    $form.StartPosition    = "CenterScreen"
    $form.BackColor        = $clrPaneBg
    $form.Font             = $segUI

    # ---- Header: fixed position, no Dock ----
    $pnlHeader             = New-Object System.Windows.Forms.Panel
    $pnlHeader.Location    = New-Object System.Drawing.Point(0, 0)
    $pnlHeader.Size        = New-Object System.Drawing.Size(1060, 62)
    $pnlHeader.Anchor      = ([System.Windows.Forms.AnchorStyles]::Top -bor `
                              [System.Windows.Forms.AnchorStyles]::Left -bor `
                              [System.Windows.Forms.AnchorStyles]::Right)
    $pnlHeader.BackColor   = $clrPurple
    $form.Controls.Add($pnlHeader)

    $lblTitle              = New-Object System.Windows.Forms.Label
    $lblTitle.Text         = "Pirum Consulting LLC  |  PC Setup & Configuration Tool"
    $lblTitle.Font         = $segHdr
    $lblTitle.ForeColor    = $clrWhite
    $lblTitle.AutoSize     = $true
    $lblTitle.Location     = New-Object System.Drawing.Point(12, 8)
    $pnlHeader.Controls.Add($lblTitle)

    $lblSub                = New-Object System.Windows.Forms.Label
    $lblSub.Text           = "Veteran Owned & Operated  |  330-597-0550  |  pirumllc.com  |  OS: $osVer"
    $lblSub.Font           = $segSm
    $lblSub.ForeColor      = $clrGold
    $lblSub.AutoSize       = $true
    $lblSub.Location       = New-Object System.Drawing.Point(14, 38)
    $pnlHeader.Controls.Add($lblSub)

    # ---- Tab control: starts below header, anchored all sides ----
    $tabs                  = New-Object System.Windows.Forms.TabControl
    $tabs.Location         = New-Object System.Drawing.Point(0, 62)
    $tabs.Size             = New-Object System.Drawing.Size(1060, 618)
    $tabs.Anchor           = ([System.Windows.Forms.AnchorStyles]::Top    -bor `
                              [System.Windows.Forms.AnchorStyles]::Bottom -bor `
                              [System.Windows.Forms.AnchorStyles]::Left   -bor `
                              [System.Windows.Forms.AnchorStyles]::Right)
    $tabs.Font             = $segUI
    $tabs.Padding          = New-Object System.Drawing.Point(10, 4)
    $form.Controls.Add($tabs)

    # ---- Footer: anchored to bottom edge, floats correctly on resize ----
    $pnlBottom           = New-Object System.Windows.Forms.Panel
    $pnlBottom.Location  = New-Object System.Drawing.Point(0, 680)
    $pnlBottom.Size      = New-Object System.Drawing.Size(1060, 50)
    $pnlBottom.Anchor    = ([System.Windows.Forms.AnchorStyles]::Bottom -bor `
                            [System.Windows.Forms.AnchorStyles]::Left   -bor `
                            [System.Windows.Forms.AnchorStyles]::Right)
    $pnlBottom.BackColor = $clrPurple
    $form.Controls.Add($pnlBottom)

    # Helper: make a scrollable tab panel
    function New-TabPage([string]$Title) {
        $tp              = New-Object System.Windows.Forms.TabPage
        $tp.Text         = $Title
        $tp.BackColor    = $clrPaneBg
        $tp.AutoScroll   = $true
        $tabs.TabPages.Add($tp)
        return $tp
    }

    # Helper: section divider label
    function Add-SectionLabel([System.Windows.Forms.Control]$Parent, [string]$Text, [int]$Y) {
        $lbl             = New-Object System.Windows.Forms.Label
        $lbl.Text        = "  $Text"
        $lbl.Font        = $segB
        $lbl.BackColor   = $clrPurple
        $lbl.ForeColor   = $clrWhite
        $lbl.Size        = New-Object System.Drawing.Size(960, 20)
        $lbl.Location    = New-Object System.Drawing.Point(6, $Y)
        $Parent.Controls.Add($lbl)
        return $Y + 22
    }

    # Helper: add checkbox row with tooltip
    function Add-CheckRow([System.Windows.Forms.Control]$Parent, [string]$Text, [bool]$Checked, [string]$Tip, [int]$Y, [int]$Indent = 16) {
        $cb              = New-Object System.Windows.Forms.CheckBox
        $cb.Text         = $Text
        $cb.Checked      = $Checked
        $cb.Size         = New-Object System.Drawing.Size(940, 19)
        $cb.Location     = New-Object System.Drawing.Point($Indent, $Y)
        $cb.ForeColor    = $clrGray
        if ($Tip) { $ttip.SetToolTip($cb, $Tip) }
        $Parent.Controls.Add($cb)
        return $cb
    }

    # ================================================================
    # TAB 2: Main Steps
    # ================================================================
    $tpMain  = New-TabPage "  1. Main Steps  "
    $pMain   = New-Object System.Windows.Forms.Panel
    $pMain.AutoScroll = $true
    $pMain.Dock = "Fill"
    $tpMain.Controls.Add($pMain)
    $y = 8

    $y = Add-SectionLabel $pMain "Core Setup" $y
    $cbSetName = Add-CheckRow $pMain "Set PC Name  (uses values from PC Naming tab)" $true "Renames the computer. Fill in the PC Naming tab first." $y; $y += 22

    # Time zone row: checkbox + inline dropdown
    $cbSetTime = New-Object System.Windows.Forms.CheckBox
    $cbSetTime.Text = "Set Time Zone:"
    $cbSetTime.Checked = $true
    $cbSetTime.Location = New-Object System.Drawing.Point(16, $y)
    $cbSetTime.Size = New-Object System.Drawing.Size(130, 19)
    $cbSetTime.ForeColor = $clrGray
    $pMain.Controls.Add($cbSetTime)

    $cmbTimeZone = New-Object System.Windows.Forms.ComboBox
    $cmbTimeZone.DropDownStyle = "DropDownList"
    $cmbTimeZone.Location = New-Object System.Drawing.Point(150, ($y - 1))
    $cmbTimeZone.Size = New-Object System.Drawing.Size(320, 22)
    $usTimeZones = @(
        "Eastern Standard Time",
        "Central Standard Time",
        "Mountain Standard Time",
        "US Mountain Standard Time",
        "Pacific Standard Time",
        "Alaskan Standard Time",
        "Hawaiian Standard Time",
        "Atlantic Standard Time"
    )
    $usTZLabels = @{
        "Eastern Standard Time"      = "Eastern  (ET)  - New York, Miami, Atlanta"
        "Central Standard Time"      = "Central  (CT)  - Chicago, Dallas, Houston"
        "Mountain Standard Time"     = "Mountain  (MT)  - Denver, Phoenix (observes DST)"
        "US Mountain Standard Time"  = "Mountain  (MT)  - Arizona (no DST)"
        "Pacific Standard Time"      = "Pacific  (PT)  - Los Angeles, Seattle"
        "Alaskan Standard Time"      = "Alaska  (AKT)  - Anchorage"
        "Hawaiian Standard Time"     = "Hawaii  (HT)  - Honolulu"
        "Atlantic Standard Time"     = "Atlantic  (AT)  - Puerto Rico, Virgin Islands"
    }
    foreach ($tzId in $usTimeZones) { $cmbTimeZone.Items.Add($usTZLabels[$tzId]) | Out-Null }
    $cmbTimeZone.SelectedIndex = 0   # Eastern default
    $ttip.SetToolTip($cmbTimeZone, "Select the time zone for this machine.")
    $pMain.Controls.Add($cmbTimeZone)
    $y += 26

    $cbPower = Add-CheckRow $pMain "Apply Pirum Power Profile  (no sleep on AC, managed DC timeouts)" $true "Creates a custom power scheme. On AC power the machine never sleeps. DC standby at 30 min." $y; $y += 26

    $y = Add-SectionLabel $pMain "Applications" $y
    $cbInstallChoco = Add-CheckRow $pMain "Install Chocolatey  (required before any app installs)" $true "Installs the Chocolatey package manager. Automatically skipped if already installed." $y; $y += 22
    $cbInstallApps  = Add-CheckRow $pMain "Install Applications  (configure on the Applications tab)" $true "Installs all checked apps from the Applications tab." $y; $y += 26

    $y = Add-SectionLabel $pMain "System Hardening" $y
    $cbReclaim = Add-CheckRow $pMain "Run Reclaim Windows  (configure on the Reclaim Windows tab)" $true "Runs all checked privacy, system, UI, and bloatware items from the Reclaim Windows tab." $y; $y += 26

    $y = Add-SectionLabel $pMain "Layout & Personalization" $y
    $cbLayout      = Add-CheckRow $pMain "Apply Layout Design  (taskbar/start menu for new user profiles)"  $true "Applies LayoutModification XML and AppAssociations XML from C:\Pirum\xml\. Delegates to the same functions as the Personalize tab." $y; $y += 22
    $cbPersonalize = Add-CheckRow $pMain "Apply Personalization  (configure on the Personalize tab)"        $true "Applies OEM branding, wallpaper, lock screen, and user pictures per selections on the Personalize tab." $y; $y += 26

    $y = Add-SectionLabel $pMain "Management Software" $y
    $cbMgmt = Add-CheckRow $pMain "Install Management Software  (configure on the Management tab)" $true "Installs selected RMM agents and remote support tools." $y; $y += 26

    $y = Add-SectionLabel $pMain "Domain" $y
    $cbDomain = Add-CheckRow $pMain "Join Domain  (will prompt for domain name and credentials)" $false "Prompts for domain name and optional OU path, then joins the machine to the domain." $y; $y += 26

    $y = Add-SectionLabel $pMain "Security" $y
    $cbBitLocker = Add-CheckRow $pMain "Enable BitLocker on C:  (requires CustData USB drive connected)" $true "Enables BitLocker with TPM protector and saves the recovery key to the \BitLocker folder on the CustData USB drive. Drive must be connected before clicking Run." $y; $y += 26

    $y = Add-SectionLabel $pMain "Finish" $y
    $cbRestart = Add-CheckRow $pMain "Restart computer when all steps complete" $true "Prompts for confirmation before restarting." $y; $y += 10

    $lblNote = New-Object System.Windows.Forms.Label
    $lblNote.Text = "Hover over any checkbox for a description. Fill in the PC Naming tab before running. All steps run in the order shown."
    $lblNote.Font = $segSm
    $lblNote.ForeColor = $clrLav
    $lblNote.AutoSize = $true
    $lblNote.Location = New-Object System.Drawing.Point(16, ($y + 8))
    $pMain.Controls.Add($lblNote)

    # ================================================================
    # TAB 2: Applications
    # ================================================================
    # ================================================================
    # TAB 1: PC Naming
    # ================================================================
    $tpNaming = New-TabPage "  2. PC Naming  "
    $pNaming  = New-Object System.Windows.Forms.Panel
    $pNaming.AutoScroll = $true
    $pNaming.Dock = "Fill"
    $tpNaming.Controls.Add($pNaming)
    $yn = 8

    $yn = Add-SectionLabel $pNaming "Computer Name  (Pirum format: COMPANY-LOCATION-TYPE-ASSETID)" $yn

    # Helper to add a label+textbox pair
    function Add-NameField([System.Windows.Forms.Panel]$Parent, [string]$LabelText, [string]$DefaultVal, [string]$Tip, [int]$Y) {
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = $LabelText
        $lbl.Location = New-Object System.Drawing.Point(16, ($Y + 3))
        $lbl.Size = New-Object System.Drawing.Size(160, 18)
        $lbl.ForeColor = $clrGray
        $Parent.Controls.Add($lbl)
        $txt = New-Object System.Windows.Forms.TextBox
        $txt.Text = $DefaultVal
        $txt.Location = New-Object System.Drawing.Point(180, $Y)
        $txt.Size = New-Object System.Drawing.Size(160, 22)
        $txt.CharacterCasing = "Upper"
        if ($Tip) { $ttip.SetToolTip($txt, $Tip) }
        $Parent.Controls.Add($txt)
        return $txt
    }

    # Device Type dropdown
    $lblDT = New-Object System.Windows.Forms.Label
    $lblDT.Text = "Device Type:"
    $lblDT.Location = New-Object System.Drawing.Point(16, ($yn + 3))
    $lblDT.Size = New-Object System.Drawing.Size(160, 18)
    $lblDT.ForeColor = $clrGray
    $pNaming.Controls.Add($lblDT)
    $cmbDeviceType = New-Object System.Windows.Forms.ComboBox
    $cmbDeviceType.DropDownStyle = "DropDownList"
    $cmbDeviceType.Location = New-Object System.Drawing.Point(180, $yn)
    $cmbDeviceType.Size = New-Object System.Drawing.Size(220, 22)
    @("D - Desktop","L - Laptop","A - All-in-One","B - Tablet","V - Server","Z - Other") | ForEach-Object { $cmbDeviceType.Items.Add($_) | Out-Null }
    $cmbDeviceType.SelectedIndex = 0
    $ttip.SetToolTip($cmbDeviceType, "Device type code used in the computer name.")
    $pNaming.Controls.Add($cmbDeviceType)
    $yn += 28

    $txtCompany  = Add-NameField $pNaming "Company Initials:" "" "Up to 3 letters. Example: ABC" $yn; $yn += 28
    $txtLocation = Add-NameField $pNaming "Location Initials:" "" "Up to 3 letters. Example: YNG for Youngstown" $yn; $yn += 28
    $txtAssetID  = Add-NameField $pNaming "Asset ID:" "" "Up to 5 digits. Example: 00042" $yn; $yn += 28

    # Live preview label
    $lblPreviewHdr = New-Object System.Windows.Forms.Label
    $lblPreviewHdr.Text = "Preview:"
    $lblPreviewHdr.Location = New-Object System.Drawing.Point(16, ($yn + 3))
    $lblPreviewHdr.Size = New-Object System.Drawing.Size(160, 18)
    $lblPreviewHdr.ForeColor = $clrGray
    $pNaming.Controls.Add($lblPreviewHdr)

    $lblPreview = New-Object System.Windows.Forms.Label
    $lblPreview.Text = "---"
    $lblPreview.Location = New-Object System.Drawing.Point(180, ($yn + 1))
    $lblPreview.Size = New-Object System.Drawing.Size(500, 22)
    $lblPreview.Font = $segB
    $lblPreview.ForeColor = $clrGreen
    $pNaming.Controls.Add($lblPreview)
    $yn += 32

    $lblCurrent = New-Object System.Windows.Forms.Label
    $lblCurrent.Text = "Current name:  $env:COMPUTERNAME"
    $lblCurrent.Location = New-Object System.Drawing.Point(180, $yn)
    $lblCurrent.AutoSize = $true
    $lblCurrent.ForeColor = $clrLav
    $lblCurrent.Font = $segSm
    $pNaming.Controls.Add($lblCurrent)
    $yn += 24

    $lblNamingNote = New-Object System.Windows.Forms.Label
    $lblNamingNote.Text = "Fill in all four fields before running. The preview updates as you type. Max total length is 15 characters."
    $lblNamingNote.Location = New-Object System.Drawing.Point(16, ($yn + 8))
    $lblNamingNote.Size = New-Object System.Drawing.Size(700, 32)
    $lblNamingNote.ForeColor = $clrLav
    $lblNamingNote.Font = $segSm
    $pNaming.Controls.Add($lblNamingNote)

    # Update preview on any change
    $updatePreview = {
        $dt = if ($cmbDeviceType.SelectedItem) { ($cmbDeviceType.SelectedItem -split " - ")[0] } else { "" }
        $co = $txtCompany.Text.Trim().ToUpper()
        $lo = $txtLocation.Text.Trim().ToUpper()
        $ai = $txtAssetID.Text.Trim()
        if ($co -and $lo -and $dt -and $ai) {
            $preview = "$co-$lo-$dt-$ai"
            $lblPreview.Text = $preview
            $lblPreview.ForeColor = if ($preview.Length -le 15) { $clrGreen } else { $clrWarn }
        } else {
            $lblPreview.Text = "---"
            $lblPreview.ForeColor = $clrGray
        }
    }
    $cmbDeviceType.Add_SelectedIndexChanged($updatePreview)
    $txtCompany.Add_TextChanged($updatePreview)
    $txtLocation.Add_TextChanged($updatePreview)
    $txtAssetID.Add_TextChanged($updatePreview)

    $tpApps  = New-TabPage "  3. Applications  "
    $pApps   = New-Object System.Windows.Forms.Panel
    $pApps.AutoScroll = $true
    $pApps.Dock = "Fill"
    $tpApps.Controls.Add($pApps)
    $y2 = 8

    $y2 = Add-SectionLabel $pApps "Baseline Applications  (uncheck to skip for this machine)" $y2
    $appCheckboxes = @{}
    foreach ($app in $script:BaselineApps) {
        $cb = Add-CheckRow $pApps "$($app.Name)   [ choco: $($app.ChocoID) ]" $app.Default "Chocolatey package ID: $($app.ChocoID)" $y2
        $appCheckboxes[$app.ChocoID] = $cb
        $y2 += 22
    }

    $y2 += 10
    $y2 = Add-SectionLabel $pApps "Microsoft 365 / Office" $y2
    $cbInstallM365 = Add-CheckRow $pApps "Install Microsoft 365  (O365BusinessRetail, Monthly channel, en-us)" $true "Runs choco install microsoft-office-deployment with the O365BusinessRetail params below. This runs as a separate install pass and takes longer than standard apps." $y2; $y2 += 22
    $lblM365Cmd = New-Object System.Windows.Forms.Label
    $lblM365Cmd.Text      = "    choco install microsoft-office-deployment --params=`"'/Channel:Monthly /Language:en-us /Product:O365BusinessRetail /Exclude:Lync,Groove'`" -y"
    $lblM365Cmd.Font      = $segSm
    $lblM365Cmd.ForeColor = $clrLav
    $lblM365Cmd.AutoSize  = $true
    $lblM365Cmd.Location  = New-Object System.Drawing.Point(32, $y2)
    $pApps.Controls.Add($lblM365Cmd)
    $y2 += 30

    $y2 = Add-SectionLabel $pApps "Add Custom App by Chocolatey ID" $y2

    $lblCustom = New-Object System.Windows.Forms.Label
    $lblCustom.Text = "Chocolatey ID:"
    $lblCustom.Location = New-Object System.Drawing.Point(16, ($y2 + 2))
    $lblCustom.AutoSize = $true
    $lblCustom.ForeColor = $clrGray
    $pApps.Controls.Add($lblCustom)

    $txtCustom = New-Object System.Windows.Forms.TextBox
    $txtCustom.Location = New-Object System.Drawing.Point(108, $y2)
    $txtCustom.Size = New-Object System.Drawing.Size(200, 22)
    $ttip.SetToolTip($txtCustom, "Type a Chocolatey package ID (e.g. 'slack', 'vscode', 'googledrive') and click Install Now")
    $pApps.Controls.Add($txtCustom)

    $btnCustom = New-Object System.Windows.Forms.Button
    $btnCustom.Text = "Install Now"
    $btnCustom.Location = New-Object System.Drawing.Point(316, ($y2 - 1))
    $btnCustom.Size = New-Object System.Drawing.Size(100, 24)
    $btnCustom.BackColor = $clrGreen
    $btnCustom.ForeColor = $clrWhite
    $btnCustom.FlatStyle = "Flat"
    $btnCustom.Font = $segB
    $pApps.Controls.Add($btnCustom)

    $y2 += 30
    $lblChocoRef = New-Object System.Windows.Forms.Label
    $lblChocoRef.Text = "Browse packages: community.chocolatey.org/packages"
    $lblChocoRef.Location = New-Object System.Drawing.Point(16, $y2)
    $lblChocoRef.AutoSize = $true
    $lblChocoRef.ForeColor = $clrLav
    $lblChocoRef.Font = $segSm
    $pApps.Controls.Add($lblChocoRef)

    $y2 += 20
    $lblBaselineNote = New-Object System.Windows.Forms.Label
    $lblBaselineNote.Text = "To update the baseline list, edit the `$script:BaselineApps section at the top of PCSetup.ps1."
    $lblBaselineNote.Location = New-Object System.Drawing.Point(16, $y2)
    $lblBaselineNote.AutoSize = $true
    $lblBaselineNote.ForeColor = $clrLav
    $lblBaselineNote.Font = $segSm
    $pApps.Controls.Add($lblBaselineNote)

    # ================================================================
    # TAB 3: Reclaim Windows
    # ================================================================
    $tpReclaim  = New-TabPage "  4. Reclaim Windows  "
    $pReclaim   = New-Object System.Windows.Forms.Panel
    $pReclaim.AutoScroll = $true
    $pReclaim.Dock = "Fill"
    $tpReclaim.Controls.Add($pReclaim)
    $y3 = 8

    $reclaimCBs = @{}
    $categories = [ordered]@{
        "Privacy"      = "Privacy & Telemetry"
        "SystemTweaks" = "System Tweaks"
        "UITweaks"     = "UI Tweaks"
        "Bloatware"    = "Bloatware Removal"
    }
    foreach ($catKey in $categories.Keys) {
        $y3 = Add-SectionLabel $pReclaim $categories[$catKey] $y3
        foreach ($item in ($reclaimItems | Where-Object { $_.Category -eq $catKey })) {
            $cb = Add-CheckRow $pReclaim $item.Label $item.Default $item.Advisory $y3
            $reclaimCBs[$item.Key] = $cb
            $y3 += 22
        }
        $y3 += 6
    }

    # Select all / none buttons
    # DNS Provider selection
    $y3 += 8
    $y3 = Add-SectionLabel $pReclaim "DNS Provider  (used by Set DNS Servers tweak above)" $y3
    $dnsProviders = [ordered]@{
        "Cloudflare  (1.1.1.1 / 1.0.0.1)  - Fast, privacy-focused"    = @("1.1.1.1","1.0.0.1")
        "Google  (8.8.8.8 / 8.8.4.4)  - Reliable, widely used"        = @("8.8.8.8","8.8.4.4")
        "AdGuard  (94.140.14.14 / 94.140.15.15)  - Blocks ads/trackers" = @("94.140.14.14","94.140.15.15")
        "Quad9  (9.9.9.9 / 149.112.112.112)  - Blocks malware domains" = @("9.9.9.9","149.112.112.112")
        "OpenDNS  (208.67.222.222 / 208.67.220.220)  - Cisco, content filtering" = @("208.67.222.222","208.67.220.220")
    }
    $cmbDNS = New-Object System.Windows.Forms.ComboBox
    $cmbDNS.DropDownStyle = "DropDownList"
    $cmbDNS.Location = New-Object System.Drawing.Point(16, $y3)
    $cmbDNS.Size = New-Object System.Drawing.Size(700, 22)
    foreach ($key in $dnsProviders.Keys) { $cmbDNS.Items.Add($key) | Out-Null }
    $cmbDNS.SelectedIndex = 0
    $ttip.SetToolTip($cmbDNS, "Select DNS provider to apply to all active ethernet adapters. Only used if Set DNS Servers is checked above.")
    $pReclaim.Controls.Add($cmbDNS)
    $y3 += 30

    # Store selected DNS in script scope for reclaim action to read
    $script:SelectedDNS = @("1.1.1.1","1.0.0.1")  # Cloudflare default
    $cmbDNS.Add_SelectedIndexChanged({
        $selected = $cmbDNS.SelectedItem
        $script:SelectedDNS = $dnsProviders[$selected]
    })

    $y3 += 8
    $btnSelAll = New-Object System.Windows.Forms.Button
    $btnSelAll.Text = "Select All"
    $btnSelAll.Location = New-Object System.Drawing.Point(16, $y3)
    $btnSelAll.Size = New-Object System.Drawing.Size(100, 26)
    $btnSelAll.BackColor = $clrPurple
    $btnSelAll.ForeColor = $clrWhite
    $btnSelAll.FlatStyle = "Flat"
    $pReclaim.Controls.Add($btnSelAll)

    $btnSelNone = New-Object System.Windows.Forms.Button
    $btnSelNone.Text = "Deselect All"
    $btnSelNone.Location = New-Object System.Drawing.Point(124, $y3)
    $btnSelNone.Size = New-Object System.Drawing.Size(100, 26)
    $btnSelNone.BackColor = $clrLav
    $btnSelNone.ForeColor = $clrWhite
    $btnSelNone.FlatStyle = "Flat"
    $pReclaim.Controls.Add($btnSelNone)

    $btnSelAll.Add_Click({
        foreach ($cb in $reclaimCBs.Values) { $cb.Checked = $true }
    })
    $btnSelNone.Add_Click({
        foreach ($cb in $reclaimCBs.Values) { $cb.Checked = $false }
    })

    # ================================================================
    # TAB 4: Personalize
    # ================================================================
    $tpPersonalize = New-TabPage "  5. Personalize  "
    $pPers = New-Object System.Windows.Forms.Panel
    $pPers.AutoScroll = $true
    $pPers.Dock = "Fill"
    $tpPersonalize.Controls.Add($pPers)
    $y4 = 8

    # ---- OEM Branding ----
    $y4 = Add-SectionLabel $pPers "OEM Information  (visible in Settings > System > About)" $y4

    $cbOEM = Add-CheckRow $pPers "Apply OEM Information" $true "Writes branding fields to the OEM registry key. Uncheck individual items below to skip them." $y4; $y4 += 24

    # Helper: OEM field row - checkbox + label + textbox
    function Add-OEMField {
        param([System.Windows.Forms.Panel]$Parent, [string]$FieldLabel, [string]$Default, [string]$Tip, [int]$Y)
        $cb = New-Object System.Windows.Forms.CheckBox
        $cb.Checked  = $true
        $cb.Text     = $FieldLabel
        $cb.Location = New-Object System.Drawing.Point(32, $Y)
        $cb.Size     = New-Object System.Drawing.Size(160, 19)
        $cb.ForeColor = $clrGray
        $Parent.Controls.Add($cb)
        $txt = New-Object System.Windows.Forms.TextBox
        $txt.Text     = $Default
        $txt.Location = New-Object System.Drawing.Point(200, ($Y - 1))
        $txt.Size     = New-Object System.Drawing.Size(500, 22)
        if ($Tip) { $ttip.SetToolTip($txt, $Tip) }
        $Parent.Controls.Add($txt)
        return [PSCustomObject]@{ Checkbox = $cb; TextBox = $txt }
    }

    $oemMfr   = Add-OEMField $pPers "Manufacturer:"   "Pirum Consulting LLC"         "Company name shown in System Properties." $y4; $y4 += 26
    $oemPhone = Add-OEMField $pPers "Support Phone:"  "330-597-0550"                 "Support phone number shown in System Properties." $y4; $y4 += 26
    $oemHours = Add-OEMField $pPers "Support Hours:"  "Mon-Fri 9am-5pm"              "Support hours shown in System Properties." $y4; $y4 += 26
    $oemURL   = Add-OEMField $pPers "Support URL:"    "http://go.pirumllc.com/portal" "Support URL shown in System Properties." $y4; $y4 += 30

    # ---- Wallpaper ----
    $y4 = Add-SectionLabel $pPers "Desktop Wallpaper" $y4
    $cbWallpaper = Add-CheckRow $pPers "Apply Desktop Wallpaper" $false "Copies the selected image to Windows wallpaper directories and sets it for all users." $y4; $y4 += 22

    $lblWallSrc = New-Object System.Windows.Forms.Label
    $lblWallSrc.Text      = "    Image file:"
    $lblWallSrc.Location  = New-Object System.Drawing.Point(32, ($y4 + 3))
    $lblWallSrc.AutoSize  = $true
    $lblWallSrc.ForeColor = $clrGray
    $pPers.Controls.Add($lblWallSrc)
    $txtWallSrc = New-Object System.Windows.Forms.TextBox
    $txtWallSrc.Text     = "C:\Pirum\media\background.jpg"
    $txtWallSrc.Location = New-Object System.Drawing.Point(120, $y4)
    $txtWallSrc.Size     = New-Object System.Drawing.Size(680, 22)
    $ttip.SetToolTip($txtWallSrc, "Full path to the wallpaper image. JPG or PNG recommended. Use the Browse button to select.")
    $pPers.Controls.Add($txtWallSrc)
    $btnWallBrowse = New-Object System.Windows.Forms.Button
    $btnWallBrowse.Text      = "Browse..."
    $btnWallBrowse.Location  = New-Object System.Drawing.Point(808, ($y4 - 1))
    $btnWallBrowse.Size      = New-Object System.Drawing.Size(80, 24)
    $btnWallBrowse.BackColor = $clrLav
    $btnWallBrowse.ForeColor = $clrWhite
    $btnWallBrowse.FlatStyle = "Flat"
    $pPers.Controls.Add($btnWallBrowse)
    $btnWallBrowse.Add_Click({
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Title            = "Select Wallpaper Image"
        $dlg.Filter           = "Image files (*.jpg;*.jpeg;*.png;*.bmp)|*.jpg;*.jpeg;*.png;*.bmp|All files (*.*)|*.*"
        $dlg.InitialDirectory = if (Test-Path "C:\Pirum\media") { "C:\Pirum\media" } else { "C:\" }
        if ($dlg.ShowDialog() -eq "OK") { $txtWallSrc.Text = $dlg.FileName }
    })
    $y4 += 30

    # ---- Lock Screen ----
    $y4 = Add-SectionLabel $pPers "Lock Screen Image" $y4
    $cbLockscreen = Add-CheckRow $pPers "Apply Lock Screen Image" $false "Copies the selected image and sets it as the Windows lock screen via policy and PersonalizationCSP." $y4; $y4 += 22

    $lblLSSrc = New-Object System.Windows.Forms.Label
    $lblLSSrc.Text      = "    Image file:"
    $lblLSSrc.Location  = New-Object System.Drawing.Point(32, ($y4 + 3))
    $lblLSSrc.AutoSize  = $true
    $lblLSSrc.ForeColor = $clrGray
    $pPers.Controls.Add($lblLSSrc)
    $txtLSSrc = New-Object System.Windows.Forms.TextBox
    $txtLSSrc.Text     = "C:\Pirum\media\lockscreen.jpg"
    $txtLSSrc.Location = New-Object System.Drawing.Point(120, $y4)
    $txtLSSrc.Size     = New-Object System.Drawing.Size(680, 22)
    $ttip.SetToolTip($txtLSSrc, "Full path to the lock screen image. JPG or PNG recommended. Use the Browse button to select.")
    $pPers.Controls.Add($txtLSSrc)
    $btnLSBrowse = New-Object System.Windows.Forms.Button
    $btnLSBrowse.Text      = "Browse..."
    $btnLSBrowse.Location  = New-Object System.Drawing.Point(808, ($y4 - 1))
    $btnLSBrowse.Size      = New-Object System.Drawing.Size(80, 24)
    $btnLSBrowse.BackColor = $clrLav
    $btnLSBrowse.ForeColor = $clrWhite
    $btnLSBrowse.FlatStyle = "Flat"
    $pPers.Controls.Add($btnLSBrowse)
    $btnLSBrowse.Add_Click({
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Title            = "Select Lock Screen Image"
        $dlg.Filter           = "Image files (*.jpg;*.jpeg;*.png;*.bmp)|*.jpg;*.jpeg;*.png;*.bmp|All files (*.*)|*.*"
        $dlg.InitialDirectory = if (Test-Path "C:\Pirum\media") { "C:\Pirum\media" } else { "C:\" }
        if ($dlg.ShowDialog() -eq "OK") { $txtLSSrc.Text = $dlg.FileName }
    })
    $y4 += 30

    # ---- User Account Pictures ----
    $y4 = Add-SectionLabel $pPers "User Account Pictures" $y4
    $cbUserPics = Add-CheckRow $pPers "Apply User Account Pictures" $false "Sets the account picture for the current user SID. Prefers C:\Pirum\media\; falls back to C:\Pirum\defmedia\ if media files are not found there." $y4; $y4 += 22

    # Detect which source folder has pictures and show status
    $boot4 = Get-Partition | Where-Object { $_.IsBoot -eq $true }
    $osd4  = $boot4.DriveLetter + ":"
    $mediaHasPics   = (Test-Path "$osd4\Pirum\media\user-32.png")   -or (Test-Path "$osd4\Pirum\media\user-32.bmp")
    $defmediaHasPics= (Test-Path "$osd4\Pirum\defmedia\user-32.png") -or (Test-Path "$osd4\Pirum\defmedia\user-32.bmp")

    $lblPicSource = New-Object System.Windows.Forms.Label
    if ($mediaHasPics) {
        $lblPicSource.Text      = "    Source: C:\Pirum\media\  (custom pictures found)"
        $lblPicSource.ForeColor = $clrGreen
    } elseif ($defmediaHasPics) {
        $lblPicSource.Text      = "    Source: C:\Pirum\defmedia\  (no custom pictures found in media\, using defaults)"
        $lblPicSource.ForeColor = $clrGold
    } else {
        $lblPicSource.Text      = "    No user picture files found in C:\Pirum\media\ or C:\Pirum\defmedia\  (step will be skipped)"
        $lblPicSource.ForeColor = $clrWarn
    }
    $lblPicSource.Font     = $segSm
    $lblPicSource.AutoSize = $true
    $lblPicSource.Location = New-Object System.Drawing.Point(32, $y4)
    $pPers.Controls.Add($lblPicSource)
    $y4 += 24

    # ---- App Associations XML ----
    $y4 = Add-SectionLabel $pPers "Default App Associations" $y4
    $cbAppAssoc = Add-CheckRow $pPers "Apply App Associations XML  (sets default programs for file types and protocols)" $true "Reads AppAssociations.xml from C:\Pirum\xml\ and applies default app mappings via DISM. Skipped gracefully if file not present." $y4; $y4 += 22
    $lblAssocSrc = New-Object System.Windows.Forms.Label
    $lblAssocSrc.Text      = if (Test-Path "C:\Pirum\xml\AppAssociations.xml") { "    Found: C:\Pirum\xml\AppAssociations.xml" } else { "    Not found: C:\Pirum\xml\AppAssociations.xml  (will be skipped at run time)" }
    $lblAssocSrc.ForeColor = if (Test-Path "C:\Pirum\xml\AppAssociations.xml") { $clrGreen } else { $clrGold }
    $lblAssocSrc.Font      = $segSm
    $lblAssocSrc.AutoSize  = $true
    $lblAssocSrc.Location  = New-Object System.Drawing.Point(32, $y4)
    $pPers.Controls.Add($lblAssocSrc)
    $y4 += 28

    # ---- Layout Modification XML ----
    $y4 = Add-SectionLabel $pPers "Layout Modification (Taskbar / Start Menu for New Users)" $y4
    $cbLayoutMod = Add-CheckRow $pPers "Apply Layout Modification XML  (taskbar and Start menu layout for new user profiles)" $true "Reads LayoutModification.xml from C:\Pirum\xml\ and applies layout via Import-StartLayout. Affects new user profiles only. Skipped gracefully if file not present." $y4; $y4 += 22
    $lblLayoutSrc = New-Object System.Windows.Forms.Label
    $layoutFound = (Test-Path "C:\Pirum\xml\LayoutModification_Win10.xml") -or
                   (Test-Path "C:\Pirum\xml\LayoutModification_Win11.xml") -or
                   (Test-Path "C:\Pirum\xml\LayoutModification.xml")
    $lblLayoutSrc.Text      = if ($layoutFound) { "    Found: C:\Pirum\xml\LayoutModification_Win10/11.xml" } else { "    Not found in C:\Pirum\xml\  (will be skipped at run time)" }
    $lblLayoutSrc.ForeColor = if ($layoutFound) { $clrGreen } else { $clrGold }
    $lblLayoutSrc.Font      = $segSm
    $lblLayoutSrc.AutoSize  = $true
    $lblLayoutSrc.Location  = New-Object System.Drawing.Point(32, $y4)
    $pPers.Controls.Add($lblLayoutSrc)
    $y4 += 28

    # ---- Reset ----
    $y4 = Add-SectionLabel $pPers "Reset" $y4
    $cbResetPers = Add-CheckRow $pPers "Reset Personalization Back to Windows Defaults  (removes all Pirum branding)" $false "Removes OEM info, wallpaper policy, lock screen policy, and PersonalizationCSP entries." $y4; $y4 += 10

    $lblPersNote = New-Object System.Windows.Forms.Label
    $lblPersNote.Text      = "Browse buttons open a file picker. Paths can also be typed or pasted directly. Steps with missing files are skipped with a warning in the log."
    $lblPersNote.Font      = $segSm
    $lblPersNote.ForeColor = $clrLav
    $lblPersNote.AutoSize  = $true
    $lblPersNote.Location  = New-Object System.Drawing.Point(16, ($y4 + 8))
    $pPers.Controls.Add($lblPersNote)

    # ================================================================
    # TAB 5: Management Software
    # ================================================================
    $tpMgmt = New-TabPage "  6. Management  "
    $pMgmt = New-Object System.Windows.Forms.Panel
    $pMgmt.AutoScroll = $true
    $pMgmt.Dock = "Fill"
    $tpMgmt.Controls.Add($pMgmt)
    $y5 = 8

    $y5 = Add-SectionLabel $pMgmt "RMM & Remote Support Agents" $y5

    # Helper: add a checkbox + URL-or-path label + textbox + browse button row
    function Add-AgentRow {
        param(
            [System.Windows.Forms.Panel]$Parent,
            [string]$CheckLabel,
            [bool]$Checked,
            [string]$CheckTip,
            [string]$DefaultSource,
            [string]$SourceTip,
            [ref]$Y
        )
        $cb = Add-CheckRow $Parent $CheckLabel $Checked $CheckTip $Y.Value
        $Y.Value += 22

        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = "    URL or file path:"
        $lbl.Location = New-Object System.Drawing.Point(32, ($Y.Value + 3))
        $lbl.AutoSize = $true
        $lbl.ForeColor = $clrGray
        $Parent.Controls.Add($lbl)

        $txt = New-Object System.Windows.Forms.TextBox
        $txt.Text = $DefaultSource
        $txt.Location = New-Object System.Drawing.Point(148, $Y.Value)
        $txt.Size = New-Object System.Drawing.Size(730, 22)
        $ttip.SetToolTip($txt, $SourceTip)
        $Parent.Controls.Add($txt)

        $btn = New-Object System.Windows.Forms.Button
        $btn.Text = "Browse..."
        $btn.Location = New-Object System.Drawing.Point(886, ($Y.Value - 1))
        $btn.Size = New-Object System.Drawing.Size(80, 24)
        $btn.BackColor = $clrLav
        $btn.ForeColor = $clrWhite
        $btn.FlatStyle = "Flat"
        $btn.Tag = $txt   # store textbox reference on the button itself to avoid closure scoping issues
        $ttip.SetToolTip($btn, "Browse for a local installer file. You can also type or paste a URL directly into the text box.")
        $Parent.Controls.Add($btn)

        # Use $this.Tag to retrieve the textbox - reliable across all PS closure contexts
        $btn.Add_Click({
            $targetTxt = $this.Tag
            $dlg = New-Object System.Windows.Forms.OpenFileDialog
            $dlg.Title            = "Select Installer File"
            $dlg.Filter           = "Installer files (*.exe;*.msi)|*.exe;*.msi|All files (*.*)|*.*"
            $dlg.InitialDirectory = if (Test-Path "C:\Pirum\agents") { "C:\Pirum\agents" } else { "C:\" }
            if ($dlg.ShowDialog() -eq "OK") { $targetTxt.Text = $dlg.FileName }
        })

        $Y.Value += 30
        return [PSCustomObject]@{ Checkbox = $cb; TextBox = $txt }
    }

    $y5ref = [ref]$y5

    $ninjaRow  = Add-AgentRow $pMgmt "Install NinjaOne RMM Agent" $false `
        "Installs the NinjaOne RMM agent. Enter the URL or local path to the client-specific MSI." `
        "" `
        "Enter a download URL (https://...) or a local/UNC file path to the NinjaOne client MSI. Download client-specific installer from: app.ninjarmm.com > Administration > Installer" `
        $y5ref
    $cbNinja    = $ninjaRow.Checkbox
    $txtNinjaPath = $ninjaRow.TextBox

    $action1Row = Add-AgentRow $pMgmt "Install Action1 Agent" $false `
        "Installs the Action1 RMM agent. Enter the URL or local path to the MSI." `
        "" `
        "Enter a download URL (https://...) or a local/UNC file path to the Action1 agent MSI. Download from: app.action1.com > Endpoints > Deploy Agent" `
        $y5ref
    $cbAction1    = $action1Row.Checkbox
    $txtA1Path    = $action1Row.TextBox

    $ihcRow     = Add-AgentRow $pMgmt "Install Instant Housecall" $true `
        "Installs the Instant Housecall host application. Defaults to the Pirum IHC download URL." `
        "https://pirumllc.instanthousecall.com/dlsecure.cgi?sub=pirumllc&specialistPreference=jim@pirumllc.com" `
        "Enter a download URL (https://...) or a local file path to the IHC installer (MSI or EXE). The file type is detected automatically. The default URL downloads directly from the Pirum Instant Housecall portal." `
        $y5ref
    $cbIHC      = $ihcRow.Checkbox
    $txtIHCPath = $ihcRow.TextBox

    $y5 = $y5ref.Value

    $lblMgmtNote = New-Object System.Windows.Forms.Label
    $lblMgmtNote.Text = "Enter a URL (https://...) or a local/UNC file path for each agent. URLs are downloaded to a temp file before install. Leave blank to skip."
    $lblMgmtNote.Font = $segSm
    $lblMgmtNote.ForeColor = $clrLav
    $lblMgmtNote.AutoSize = $true
    $lblMgmtNote.Location = New-Object System.Drawing.Point(16, ($y5 + 8))
    $pMgmt.Controls.Add($lblMgmtNote)

    # ================================================================
    # TAB 6: Log Output
    # ================================================================
    $tpLog = New-TabPage "  7. Log  "

    $script:LogBox = New-Object System.Windows.Forms.RichTextBox
    $script:LogBox.Dock        = "Fill"
    $script:LogBox.BackColor   = $clrDarkBg
    $script:LogBox.ForeColor   = [System.Drawing.ColorTranslator]::FromHtml("#b0f080")
    $script:LogBox.Font        = New-Object System.Drawing.Font("Consolas", 9)
    $script:LogBox.ReadOnly    = $true
    $script:LogBox.ScrollBars  = "Vertical"
    $tpLog.Controls.Add($script:LogBox)

    # ================================================================
    # ---- Bottom action bar: populate buttons (panel created above) ----
    # ================================================================
    $btnRun              = New-Object System.Windows.Forms.Button
    $btnRun.Text         = "RUN SELECTED STEPS"
    $btnRun.Size         = New-Object System.Drawing.Size(190, 32)
    $btnRun.Location     = New-Object System.Drawing.Point(10, 9)
    $btnRun.BackColor    = $clrGreen
    $btnRun.ForeColor    = $clrWhite
    $btnRun.FlatStyle    = "Flat"
    $btnRun.Font         = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $pnlBottom.Controls.Add($btnRun)

    $btnExit             = New-Object System.Windows.Forms.Button
    $btnExit.Text        = "Exit"
    $btnExit.Size        = New-Object System.Drawing.Size(70, 32)
    $btnExit.Location    = New-Object System.Drawing.Point(208, 9)
    $btnExit.BackColor   = $clrLav
    $btnExit.ForeColor   = $clrWhite
    $btnExit.FlatStyle   = "Flat"
    $pnlBottom.Controls.Add($btnExit)

    $btnSaveLog          = New-Object System.Windows.Forms.Button
    $btnSaveLog.Text     = "Save Log"
    $btnSaveLog.Size     = New-Object System.Drawing.Size(80, 32)
    $btnSaveLog.Location = New-Object System.Drawing.Point(286, 9)
    $btnSaveLog.BackColor = $clrGold
    $btnSaveLog.ForeColor = $clrDark
    $btnSaveLog.FlatStyle = "Flat"
    $btnSaveLog.Font      = $segUI
    $pnlBottom.Controls.Add($btnSaveLog)

    $lblStatus           = New-Object System.Windows.Forms.Label
    $lblStatus.Text      = "Ready. Configure options above, then click RUN SELECTED STEPS."
    $lblStatus.ForeColor = $clrGold
    $lblStatus.Font      = $segUI
    $lblStatus.AutoSize  = $true
    $lblStatus.Location  = New-Object System.Drawing.Point(378, 16)
    $pnlBottom.Controls.Add($lblStatus)

    # ================================================================
    # Event: custom install button
    # ================================================================
    $btnCustom.Add_Click({
        $id = $txtCustom.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($id)) {
            [System.Windows.Forms.MessageBox]::Show("Enter a Chocolatey package ID first.", "Missing Input", "OK", "Warning") | Out-Null
            return
        }
        $tabs.SelectedTab = $tpLog
        Invoke-InstallCustomApp -ChocoID $id
    })

    # ================================================================
    # Event: RUN button
    # ================================================================
    $btnRun.Add_Click({
        $btnRun.Enabled  = $false
        $lblStatus.Text  = "Running..."
        $tabs.SelectedTab = $tpLog
        $script:LogBox.Clear()
        Write-Log "=================================================="
        Write-Log " Pirum Consulting LLC - PC Setup Tool"
        Write-Log " $(Get-Date)   OS: $(Get-WinVersion)"
        Write-Log "=================================================="

        try {
            # 1. Set PC Name
            if ($cbSetName.Checked) {
                $dt = if ($cmbDeviceType.SelectedItem) { ($cmbDeviceType.SelectedItem -split " - ")[0] } else { "" }
                Invoke-SetPCName -DeviceType $dt -Company $txtCompany.Text.Trim().ToUpper() `
                                 -Location $txtLocation.Text.Trim().ToUpper() -AssetID $txtAssetID.Text.Trim()
            }

            # 2. Set Time Zone
            if ($cbSetTime.Checked) {
                # $usTimeZones array index matches $cmbTimeZone dropdown index
                $tzId = $usTimeZones[$cmbTimeZone.SelectedIndex]
                if (-not $tzId) { $tzId = "Eastern Standard Time" }
                Invoke-SetTimeZone -TimeZoneId $tzId
            }

            # 3. Power Profile
            if ($cbPower.Checked) { Invoke-SetPowerProfile }

            # 3b. Defender exclusion for tech tools USB drive
            Invoke-AddDefenderExclusion

            # 4. Chocolatey
            if ($cbInstallChoco.Checked -or $cbInstallApps.Checked -or $cbInstallM365.Checked) { Invoke-InstallChoco }

            # 5. Install Apps
            if ($cbInstallApps.Checked) {
                Write-Log ">>> Installing baseline applications..."
                $selected = $appCheckboxes.GetEnumerator() |
                    Where-Object { $_.Value.Checked } |
                    ForEach-Object { $_.Key }
                Invoke-InstallApps -SelectedIDs @($selected)
            }

            # 5b. Install Microsoft 365
            if ($cbInstallM365.Checked) { Invoke-InstallM365 }

            # 6. Reclaim Windows
            if ($cbReclaim.Checked) {
                Write-Log ">>> Running Reclaim Windows tweaks..."
                foreach ($item in $reclaimItems) {
                    if ($reclaimCBs[$item.Key].Checked) {
                        Write-Log "    Applying: $($item.Label)"
                        try { & $item.Action }
                        catch { Write-Log "    ERROR: $_" ([System.Drawing.Color]::Red) }
                    }
                }
                Write-Log "    Reclaim Windows complete."
            }

            # 7. Layout Design
            if ($cbLayout.Checked) { Invoke-LayoutDesign }

            # 8. Personalize
            if ($cbPersonalize.Checked) {
                Write-Log ">>> Applying personalization..."
                Invoke-Personalize `
                    -DoOEM               $cbOEM.Checked                `
                    -OEMSetManufacturer  $oemMfr.Checkbox.Checked      `
                    -OEMManufacturer     $oemMfr.TextBox.Text          `
                    -OEMSetPhone         $oemPhone.Checkbox.Checked    `
                    -OEMPhone            $oemPhone.TextBox.Text        `
                    -OEMSetHours         $oemHours.Checkbox.Checked    `
                    -OEMHours            $oemHours.TextBox.Text        `
                    -OEMSetURL           $oemURL.Checkbox.Checked      `
                    -OEMURL              $oemURL.TextBox.Text          `
                    -DoWallpaper         $cbWallpaper.Checked          `
                    -WallpaperSrc        $txtWallSrc.Text              `
                    -DoLockscreen        $cbLockscreen.Checked         `
                    -LockscreenSrc       $txtLSSrc.Text                `
                    -DoUserPictures      $cbUserPics.Checked           `
                    -DoReset             $cbResetPers.Checked
            }

            # 8b. App Associations XML
            if ($cbAppAssoc.Checked) {
                Invoke-ApplyAppAssociations -XmlPath "C:\Pirum\xml\AppAssociations.xml"
            }

            # 8c. Layout Modification XML
            if ($cbLayoutMod.Checked) {
                Invoke-ApplyLayoutModification -XmlBasePath "C:\Pirum\xml"
            }

            # 9. Management Software
            if ($cbMgmt.Checked) {
                Write-Log ">>> Installing management software..."
                Invoke-InstallManagementSoftware `
                    -DoNinja       $cbNinja.Checked   `
                    -DoAction1     $cbAction1.Checked `
                    -DoIHC         $cbIHC.Checked     `
                    -NinjaSource   $txtNinjaPath.Text `
                    -Action1Source $txtA1Path.Text    `
                    -IHCSource     $txtIHCPath.Text
            }

            # 9b. System Restore Point (before any domain join or reboot)
            Invoke-CreateRestorePoint

            # 9c. Log TPM and Secure Boot status
            Invoke-LogSecurityHardwareStatus

            # 9d. BitLocker
            if ($cbBitLocker.Checked) { Invoke-EnableBitLocker }

            # 10. Join Domain
            if ($cbDomain.Checked) { Invoke-JoinDomain }

            Write-Log ""
            Write-Log "=================================================="
            Write-Log " All selected steps complete."
            Write-Log "=================================================="
            $lblStatus.Text = "All steps complete. Check log for details."

        } catch {
            Write-Log "CRITICAL ERROR: $_" ([System.Drawing.Color]::Red)
            $lblStatus.Text = "Error encountered. See log."
        }

        $btnRun.Enabled = $true

        # Restart prompt
        if ($cbRestart.Checked) {
            $res = [System.Windows.Forms.MessageBox]::Show(
                "All selected steps are complete.`n`nRestart the computer now?",
                "Restart",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )
            if ($res -eq "Yes") { Restart-Computer -Force }
        }
    })

    $btnExit.Add_Click({ $form.Close() })

    $btnSaveLog.Add_Click({
        if ([string]::IsNullOrWhiteSpace($script:LogBox.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Nothing in the log to save yet.", "Save Log", "OK", "Information") | Out-Null
            return
        }
        if (-not (Test-Path "C:\Pirum")) { New-Item "C:\Pirum" -ItemType Directory -Force | Out-Null }
        $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
        $logPath   = "C:\Pirum\PCSetup_Log_$timestamp.txt"
        try {
            $script:LogBox.Text | Out-File -FilePath $logPath -Encoding UTF8 -Force
            $lblStatus.Text = "Log saved: $logPath"
            [System.Windows.Forms.MessageBox]::Show("Log saved to:`n$logPath", "Log Saved", "OK", "Information") | Out-Null
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to save log:`n$_", "Error", "OK", "Error") | Out-Null
        }
    })

    [System.Windows.Forms.Application]::EnableVisualStyles()
    $form.ShowDialog() | Out-Null
}

# ============================================================
# ENTRY POINT
# ============================================================
# Re-launch as admin if needed
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]"Administrator")) {
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    Exit
}

Show-MainForm
