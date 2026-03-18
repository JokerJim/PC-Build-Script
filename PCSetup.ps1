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
    [PSCustomObject]@{ Name = "Microsoft 365 / Office"; ChocoID = "microsoft-office-deployment"; Default = $true  }
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
            & choco install $id --params=$params -y 2>&1 | ForEach-Object { Write-Log "      $_" }
        } else {
            & choco install $id -y 2>&1 | ForEach-Object { Write-Log "      $_" }
        }
    }
    Write-Log "    App installation complete."
}

# ============================================================
# SECTION: Install custom Choco app by ID
# ============================================================
function Invoke-InstallCustomApp {
    param([string]$ChocoID)
    if ([string]::IsNullOrWhiteSpace($ChocoID)) { return }
    Invoke-InstallChoco
    Write-Log ">>> Installing custom package: $ChocoID"
    & choco install $ChocoID -y 2>&1 | ForEach-Object { Write-Log "      $_" }
    Write-Log "    Done."
}

# ============================================================
# SECTION: Power Profile
# ============================================================
function Invoke-SetPowerProfile {
    Write-Log ">>> Applying Pirum Power Management profile..."
    $schemeGUID = "381b4222-f694-41f0-9685-ff5bb260aaaa"
    POWERCFG -DUPLICATESCHEME 381b4222-f694-41f0-9685-ff5bb260df2e $schemeGUID 2>$null
    POWERCFG -CHANGENAME $schemeGUID "Pirum Power Management" 2>$null
    POWERCFG -SETACTIVE $schemeGUID
    POWERCFG -Change -monitor-timeout-ac 30
    POWERCFG -Change -monitor-timeout-dc 10
    POWERCFG -Change -disk-timeout-ac 30
    POWERCFG -Change -disk-timeout-dc 5
    POWERCFG -Change -standby-timeout-ac 0
    POWERCFG -Change -standby-timeout-dc 30
    POWERCFG -Change -hibernate-timeout-ac 0
    POWERCFG -Change -hibernate-timeout-dc 0
    Write-Log "    Power profile applied (AC: no sleep/hibernate; DC: sleep at 30 min)."
}

# ============================================================
# SECTION: Layout Design (new user profiles)
# ============================================================
function Invoke-LayoutDesign {
    Write-Log ">>> Applying default layout for new user profiles..."
    $boot   = Get-Partition | Where-Object { $_.IsBoot -eq $true }
    $OSDISK = $boot.DriveLetter + ":"
    $xmlPath   = "C:\Pirum\PC-Build-Script-master\LayoutModification.xml"
    $assocPath = "C:\Pirum\PC-Build-Script-master\AppAssociations.xml"
    if (Test-Path $xmlPath) {
        Import-StartLayout -LayoutPath $xmlPath -MountPath "$OSDISK\" -ErrorAction SilentlyContinue
        Write-Log "    Start/taskbar layout applied."
    } else {
        Write-Log "    WARNING: $xmlPath not found. Skipping layout." ([System.Drawing.Color]::Yellow)
    }
    if (Test-Path $assocPath) {
        dism /online /Import-DefaultAppAssociations:$assocPath | Out-Null
        Write-Log "    Default app associations applied."
    } else {
        Write-Log "    WARNING: $assocPath not found. Skipping app associations." ([System.Drawing.Color]::Yellow)
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
function Resolve-Installer {
    # Given a URL or local path, returns a local file path ready to execute.
    # If it is a URL, downloads to a temp file and returns that path.
    # Returns $null and logs an error if the source cannot be resolved.
    param([string]$Source, [string]$Label, [string]$Extension)
    if ([string]::IsNullOrWhiteSpace($Source)) {
        Write-Log "    $Label`: no path or URL configured. Skipping." ([System.Drawing.Color]::Yellow)
        return $null
    }
    if ($Source -match '^https?://') {
        Write-Log "    Downloading $Label from URL..."
        $tmpFile = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "PirumAgent_$Label$Extension")
        try {
            [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
            Invoke-WebRequest -Uri $Source -OutFile $tmpFile -UseBasicParsing -ErrorAction Stop
            Write-Log "    Download complete: $tmpFile"
            return $tmpFile
        } catch {
            Write-Log "    ERROR downloading $Label`: $_" ([System.Drawing.Color]::Red)
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
        $installer = Resolve-Installer -Source $NinjaSource -Label "NinjaOne" -Extension ".msi"
        if ($installer) {
            Start-Process msiexec.exe -ArgumentList "/i `"$installer`" /quiet /norestart" -Wait -ErrorAction SilentlyContinue
            Write-Log "    NinjaOne agent install complete."
        }
    }

    if ($DoAction1) {
        Write-Log ">>> Installing Action1 agent..."
        $installer = Resolve-Installer -Source $Action1Source -Label "Action1" -Extension ".msi"
        if ($installer) {
            Start-Process msiexec.exe -ArgumentList "/i `"$installer`" /quiet /norestart" -Wait -ErrorAction SilentlyContinue
            Write-Log "    Action1 agent install complete."
        }
    }

    if ($DoIHC) {
        Write-Log ">>> Installing Instant Housecall..."
        $installer = Resolve-Installer -Source $IHCSource -Label "IHC" -Extension ".exe"
        if ($installer) {
            Start-Process $installer -Wait -ErrorAction SilentlyContinue
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
        [bool]$DoWallpaper,
        [bool]$DoLockscreen,
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
        $lsPerPath = "HKLM:\Software\Policies\Microsoft\Windows\Personalization"
        Remove-ItemProperty -Path $lsPerPath -Name "LockScreenImage" -ErrorAction SilentlyContinue
        $cspPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP"
        Remove-Item -Path $cspPath -Recurse -Force -ErrorAction SilentlyContinue
        Write-Log "    Personalization reset complete."
        return
    }

    # OEM Information
    if ($DoOEM) {
        Write-Log ">>> Applying OEM branding..."
        $oemPath   = "HKLM:\Software\Microsoft\Windows\CurrentVersion\OEMInformation"
        $OEMLogo   = "OEMLogo.bmp"
        $logoSrc   = "$OSDISK\Pirum\defmedia\$OEMLogo"
        if (Test-Path $logoSrc) {
            Copy-Item $logoSrc "$OSDISK\windows\system32" -Force
            Copy-Item $logoSrc "$OSDISK\windows\system32\oobe\info" -Force
            Set-ItemProperty -Path $oemPath -Name Logo -Value "$OSDISK\Windows\System32\$OEMLogo"
        }
        Set-ItemProperty -Path $oemPath -Name Manufacturer  -Value "Pirum Consulting LLC"
        Set-ItemProperty -Path $oemPath -Name SupportPhone  -Value "330-597-0550"
        Set-ItemProperty -Path $oemPath -Name SupportHours  -Value "Mon-Fri 9am-5pm"
        Set-ItemProperty -Path $oemPath -Name SupportURL    -Value "http://go.pirumllc.com/portal"
        Write-Log "    OEM info set (Pirum Consulting LLC, 330-597-0550)."
    }

    # Desktop Wallpaper
    if ($DoWallpaper) {
        Write-Log ">>> Applying desktop wallpaper..."
        $Wallpaper  = "background.jpg"
        $wallSrc    = "$OSDISK\Pirum\media\$Wallpaper"
        if (Test-Path $wallSrc) {
            $bgDir = "c:\windows\system32\oobe\info\backgrounds"
            if (-not (Test-Path $bgDir)) { New-Item $bgDir -ItemType Directory -Force | Out-Null }
            Copy-Item $wallSrc $bgDir -Force
            Copy-Item $wallSrc "$OSDISK\Windows\Web\Screen" -Force -ErrorAction SilentlyContinue
            Copy-Item $wallSrc "$OSDISK\Windows\Web\Wallpaper\Windows" -Force -ErrorAction SilentlyContinue
            $bgPath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
            Ensure-RegPath $bgPath
            Set-ItemProperty -Path $bgPath -Name Wallpaper      -Value "$OSDISK\Windows\Web\Wallpaper\Windows\$Wallpaper"
            Set-ItemProperty -Path $bgPath -Name WallpaperStyle -Value "2"
            Write-Log "    Wallpaper applied."
        } else {
            Write-Log "    Wallpaper source not found: $wallSrc" ([System.Drawing.Color]::Yellow)
        }
    }

    # Lock Screen
    if ($DoLockscreen) {
        Write-Log ">>> Applying lock screen..."
        $LockScreen = "lockscreen.jpg"
        $lsSrc      = "$OSDISK\Pirum\media\$LockScreen"
        if (Test-Path $lsSrc) {
            Copy-Item $lsSrc "$OSDISK\Windows\System32" -Force
            $sysPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System"
            Ensure-RegPath $sysPath
            Set-ItemProperty -Path $sysPath -Name UseOEMBackground -Value 1
            $perPath = "HKLM:\Software\Policies\Microsoft\Windows\Personalization"
            Ensure-RegPath $perPath
            Set-ItemProperty -Path $perPath -Name LockScreenImage -Value "$OSDISK\Windows\System32\$LockScreen"
            $cspParent = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion"
            $cspPath   = "$cspParent\PersonalizationCSP"
            Ensure-RegPath $cspPath
            New-ItemProperty -Path $cspPath -Name LockScreenImageStatus -Value 1          -PropertyType DWORD  -Force | Out-Null
            New-ItemProperty -Path $cspPath -Name LockScreenImagePath   -Value "$OSDISK\Windows\System32\$LockScreen" -PropertyType STRING -Force | Out-Null
            New-ItemProperty -Path $cspPath -Name LockScreenImageUrl    -Value "$OSDISK\Windows\System32\$LockScreen" -PropertyType STRING -Force | Out-Null
            Write-Log "    Lock screen applied."
        } else {
            Write-Log "    Lock screen source not found: $lsSrc" ([System.Drawing.Color]::Yellow)
        }
    }

    # User Account Pictures
    if ($DoUserPictures) {
        Write-Log ">>> Applying user account pictures..."
        $userPicDest = "$OSDISK\ProgramData\Microsoft\User Account Pictures"
        foreach ($ext in @("png","bmp")) {
            $src = "$OSDISK\Pirum\media\user.$ext"
            if (Test-Path $src) { Copy-Item $src $userPicDest -Force }
        }
        $regBase = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AccountPicture\Users"
        Ensure-RegPath "$regBase\$SID"
        $userRegPath = "$regBase\$SID"
        foreach ($size in @("32","40","48","96","192","240")) {
            foreach ($ext in @("bmp","png")) {
                $src = "$OSDISK\Pirum\media\user-$size.$ext"
                if (Test-Path $src) {
                    Copy-Item $src $userPicDest -Force
                    Set-ItemProperty -Path $userRegPath -Name "Image$size" -Value "$userPicDest\user-$size.$ext" -ErrorAction SilentlyContinue
                }
            }
        }
        Write-Log "    User account pictures applied."
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
    $cbLayout      = Add-CheckRow $pMain "Apply Layout Design  (taskbar/start menu for new user profiles)"  $true "Imports LayoutModification.xml and AppAssociations.xml from C:\Pirum\PC-Build-Script-master\." $y; $y += 22
    $cbPersonalize = Add-CheckRow $pMain "Apply Personalization  (configure on the Personalize tab)"        $true "Applies OEM branding, wallpaper, lock screen, and user pictures per selections on the Personalize tab." $y; $y += 26

    $y = Add-SectionLabel $pMain "Management Software" $y
    $cbMgmt = Add-CheckRow $pMain "Install Management Software  (configure on the Management tab)" $true "Installs selected RMM agents and remote support tools." $y; $y += 26

    $y = Add-SectionLabel $pMain "Domain" $y
    $cbDomain = Add-CheckRow $pMain "Join Domain  (will prompt for domain name and credentials)" $false "Prompts for domain name and optional OU path, then joins the machine to the domain." $y; $y += 26

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
    $lblBaselineNote.Text = "To update the baseline list, edit the `$script:BaselineApps section at the top of PCSetup_v2.ps1."
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

    $y4 = Add-SectionLabel $pPers "OEM Branding" $y4
    $cbOEM = Add-CheckRow $pPers "Apply OEM Information  (Manufacturer, support phone 330-597-0550, hours, URL)" $true "Sets the branding info visible in Settings > System > About. Also copies OEMLogo.bmp if found in C:\Pirum\defmedia\." $y4; $y4 += 26

    $y4 = Add-SectionLabel $pPers "Backgrounds" $y4
    $cbWallpaper  = Add-CheckRow $pPers "Apply Desktop Wallpaper       source: C:\Pirum\media\background.jpg" $true "Copies wallpaper to Windows web/screen directories and sets it for the current user profile." $y4; $y4 += 22
    $cbLockscreen = Add-CheckRow $pPers "Apply Lock Screen Image         source: C:\Pirum\media\lockscreen.jpg"      $true "Copies lockscreen.jpg to System32, sets it via policy and PersonalizationCSP registry keys." $y4; $y4 += 26

    $y4 = Add-SectionLabel $pPers "User Account Pictures" $y4
    $cbUserPics = Add-CheckRow $pPers "Apply User Account Pictures   source: C:\Pirum\media\user-[32|40|48|96|192|240].png/.bmp" $true "Sets account picture for the current user SID. Requires sized image files in C:\Pirum\media\." $y4; $y4 += 26

    $y4 = Add-SectionLabel $pPers "Reset" $y4
    $cbResetPers = Add-CheckRow $pPers "Reset Personalization Back to Windows Defaults  (removes Pirum branding)" $false "Removes OEM info, wallpaper policy, lock screen policy, and PersonalizationCSP entries. Use when repurposing a machine." $y4; $y4 += 10

    $lblPersNote = New-Object System.Windows.Forms.Label
    $lblPersNote.Text = "All source files must be in place before running. Steps with missing source files are skipped gracefully with a warning in the log."
    $lblPersNote.Font = $segSm
    $lblPersNote.ForeColor = $clrLav
    $lblPersNote.AutoSize = $true
    $lblPersNote.Location = New-Object System.Drawing.Point(16, ($y4 + 8))
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

    # Helper: add a checkbox + URL-or-path label + textbox row
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
        $txt.Location = New-Object System.Drawing.Point(148, ($Y.Value))
        $txt.Size = New-Object System.Drawing.Size(820, 22)
        $ttip.SetToolTip($txt, $SourceTip)
        $Parent.Controls.Add($txt)
        $Y.Value += 28
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
        "Enter a download URL (https://...) or a local file path to the IHC setup EXE. The default URL downloads directly from the Pirum Instant Housecall portal." `
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

    $lblStatus           = New-Object System.Windows.Forms.Label
    $lblStatus.Text      = "Ready. Configure options above, then click RUN SELECTED STEPS."
    $lblStatus.ForeColor = $clrGold
    $lblStatus.Font      = $segUI
    $lblStatus.AutoSize  = $true
    $lblStatus.Location  = New-Object System.Drawing.Point(290, 16)
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

            # 4. Chocolatey
            if ($cbInstallChoco.Checked -or $cbInstallApps.Checked) { Invoke-InstallChoco }

            # 5. Install Apps
            if ($cbInstallApps.Checked) {
                Write-Log ">>> Installing baseline applications..."
                $selected = $appCheckboxes.GetEnumerator() |
                    Where-Object { $_.Value.Checked } |
                    ForEach-Object { $_.Key }
                Invoke-InstallApps -SelectedIDs @($selected)
            }

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
                    -DoOEM         $cbOEM.Checked        `
                    -DoWallpaper   $cbWallpaper.Checked  `
                    -DoLockscreen  $cbLockscreen.Checked `
                    -DoUserPictures $cbUserPics.Checked  `
                    -DoReset       $cbResetPers.Checked
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
