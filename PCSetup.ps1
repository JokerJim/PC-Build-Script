Set-ExecutionPolicy remotesigned
# Begin by creating the various functions which will be called at the end of the script. You can create additional functions if needed.
function SetPCName {
    # In our MSP we designate all systems in the format devicetype-companyname-assetid for example DT-MSP-000001 keep in mind that this is the maximum length Windows allows for system names
    # This function creates VisualBasic pop-up prompts which ask for this information to be input. You can hange these as needed to suite your MSP
    Add-Type -AssemblyName Microsoft.VisualBasic
    $DeviceType = [Microsoft.VisualBasic.Interaction]::InputBox('Enter Device Type (L)aptop, (A)ll-in-one, Ta(B)let, Ser(V)er, (D)esktop or (Z)other', 'Device Type')
    $CompanyName = [Microsoft.VisualBasic.Interaction]::InputBox('Enter Company Initials (Max 3 letters)', 'Company Initials')
    $LocationName = [Microsoft.VisualBasic.Interaction]::InputBox('Enter Location Initials (Max 3 letters)', 'Location Initials')
    $AssetID = [Microsoft.VisualBasic.Interaction]::InputBox('Enter a Asset ID (Max 5 digits)', 'Asset ID')
    Write-Output "The asset ID is $AssetID"
    Write-Output "$CompanyName-$LocationName-$DeviceType$AssetID"
    Rename-Computer -NewName "$CompanyName-$LocationName-$DeviceType-$AssetID"
}

function InstallChoco {
    # Ask for elevated permissions if required
    If (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")) {
        Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
        Exit
        }
    # Install Chocolatey to allow automated installation of packages  
    # Set-ExecutionPolicy Bypass -Scope Process -Force; iex ((New-Object System.Net.WebClient).DownloadString('http://www.hruskaj.com/chocolateyinstall.ps1'))
    Set-ExecutionPolicy Bypass -Scope Process -Force; iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
    }

function InstallApps {
    # Install the first set of applications. these are quick so ive added them separately
    choco install googlechrome zoom adobereader 7zip microsoft-edge firefox notepadplusplus unchecky -y
    # Install Office365 applications. This takes a while so is done separately. You can change the options here by following the instructions here: https://chocolatey.org/packages/microsoft-office-deployment
    # choco install microsoft-office-deployment --params="'/Channel:Monthly /Language:en-us /64bit /Product:O365BusinessRetail /Exclude:Lync,Groove'" -y
    choco install microsoft-office-deployment --params="'/Channel:Monthly /Language:en-us /Product:O365BusinessRetail /Exclude:Lync,Groove'" -y
}

function ReclaimWindows10 {
    # Ask for elevated permissions if required
    If (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")) {
        Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
        Exit
    }


    # Massive deployment section. There are stacks of customization options here. Un-hash the ones your want to apply.
    ##########
    # Privacy Settings
    ##########

    # Disable Telemetry
    Write-Host "Disabling Telemetry..."
    Set-ItemProperty -Path "HKLM:\Software\Policies\Microsoft\Windows\DataCollection" -Name "AllowTelemetry" -Type DWord -Value 0

    # Enable Telemetry
    # Remove-ItemProperty -Path "HKLM:\Software\Policies\Microsoft\Windows\DataCollection" -Name "AllowTelemetry"

    # Disable Wi-Fi Sense
    Write-Host "Disabling Wi-Fi Sense..."
    If (!(Test-Path "HKLM:\Software\Microsoft\PolicyManager\default\WiFi\AllowWiFiHotSpotReporting")) {
        New-Item -Path "HKLM:\Software\Microsoft\PolicyManager\default\WiFi\AllowWiFiHotSpotReporting" -Force | Out-Null
    }
    Set-ItemProperty -Path "HKLM:\Software\Microsoft\PolicyManager\default\WiFi\AllowWiFiHotSpotReporting" -Name "Value" -Type DWord -Value 0
    Set-ItemProperty -Path "HKLM:\Software\Microsoft\PolicyManager\default\WiFi\AllowAutoConnectToWiFiSenseHotspots" -Name "Value" -Type DWord -Value 0

    # Enable Wi-Fi Sense
    # Set-ItemProperty -Path "HKLM:\Software\Microsoft\PolicyManager\default\WiFi\AllowWiFiHotSpotReporting" -Name "Value" -Type DWord -Value 1
    # Set-ItemProperty -Path "HKLM:\Software\Microsoft\PolicyManager\default\WiFi\AllowAutoConnectToWiFiSenseHotspots" -Name "Value" -Type DWord -Value 1

    # Disable SmartScreen Filter
    # Write-Host "Disabling SmartScreen Filter..."
    # Set-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer" -Name "SmartScreenEnabled" -Type String -Value "Off"
    # Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\AppHost" -Name "EnableWebContentEvaluation" -Type DWord -Value 0

    # Enable SmartScreen Filter
    # Set-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer" -Name "SmartScreenEnabled" -Type String -Value "RequireAdmin"
    # Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\AppHost" -Name "EnableWebContentEvaluation"

    # Disable Bing Search in Start Menu
    # Write-Host "Disabling Bing Search in Start Menu..."
    # Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search" -Name "BingSearchEnabled" -Type DWord -Value 0

    # Enable Bing Search in Start Menu
    # Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search" -Name "BingSearchEnabled"

    # Disable Location Tracking
    Write-Host "Disabling Location Tracking..."
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Sensor\Overrides\{BFA794E4-F964-4FDB-90F6-51056BFE4B44}" -Name "SensorPermissionState" -Type DWord -Value 0
    Set-ItemProperty -Path "HKLM:\System\CurrentControlSet\Services\lfsvc\Service\Configuration" -Name "Status" -Type DWord -Value 0

    # Enable Location Tracking
    # Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Sensor\Overrides\{BFA794E4-F964-4FDB-90F6-51056BFE4B44}" -Name "SensorPermissionState" -Type DWord -Value 1
    # Set-ItemProperty -Path "HKLM:\System\CurrentControlSet\Services\lfsvc\Service\Configuration" -Name "Status" -Type DWord -Value 1

    # Disable Feedback
    Write-Host "Disabling Feedback..."
    If (!(Test-Path "HKCU:\Software\Microsoft\Siuf\Rules")) {
        New-Item -Path "HKCU:\Software\Microsoft\Siuf\Rules" -Force | Out-Null
    }
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Siuf\Rules" -Name "NumberOfSIUFInPeriod" -Type DWord -Value 0

    # Enable Feedback
    # Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Siuf\Rules" -Name "NumberOfSIUFInPeriod"

    # Disable Advertising ID
    Write-Host "Disabling Advertising ID..."
    If (!(Test-Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo")) {
        New-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo" | Out-Null
    }
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo" -Name "Enabled" -Type DWord -Value 0

    # Enable Advertising ID
    # Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo" -Name "Enabled"

    # Disable Cortana
    Write-Host "Disabling Cortana..."
    If (!(Test-Path "HKCU:\Software\Microsoft\Personalization\Settings")) {
        New-Item -Path "HKCU:\Software\Microsoft\Personalization\Settings" -Force | Out-Null
    }
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Personalization\Settings" -Name "AcceptedPrivacyPolicy" -Type DWord -Value 0
    If (!(Test-Path "HKCU:\Software\Microsoft\InputPersonalization")) {
        New-Item -Path "HKCU:\Software\Microsoft\InputPersonalization" -Force | Out-Null
    }
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\InputPersonalization" -Name "RestrictImplicitTextCollection" -Type DWord -Value 1
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\InputPersonalization" -Name "RestrictImplicitInkCollection" -Type DWord -Value 1
    If (!(Test-Path "HKCU:\Software\Microsoft\InputPersonalization\TrainedDataStore")) {
        New-Item -Path "HKCU:\Software\Microsoft\InputPersonalization\TrainedDataStore" -Force | Out-Null
    }
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\InputPersonalization\TrainedDataStore" -Name "HarvestContacts" -Type DWord -Value 0

    # Enable Cortana
    # Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Personalization\Settings" -Name "AcceptedPrivacyPolicy"
    # Set-ItemProperty -Path "HKCU:\Software\Microsoft\InputPersonalization" -Name "RestrictImplicitTextCollection" -Type DWord -Value 0
    # Set-ItemProperty -Path "HKCU:\Software\Microsoft\InputPersonalization" -Name "RestrictImplicitInkCollection" -Type DWord -Value 0
    # Remove-ItemProperty -Path "HKCU:\Software\Microsoft\InputPersonalization\TrainedDataStore" -Name "HarvestContacts"

	#From Remove Windows10 Bloatware
	#Stops Cortana from being used as part of your Windows Search Function
	Write-Host "Stopping Cortana from being used as part of your Windows Search Function"
	$Search = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search'
	If (Test-Path $Search) {
		Set-ItemProperty $Search AllowCortana -Value 0
	}
          
    # Restrict Windows Update P2P only to local network
    Write-Host "Restricting Windows Update P2P only to local network..."
    Set-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\DeliveryOptimization\Config" -Name "DODownloadMode" -Type DWord -Value 1
    If (!(Test-Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\DeliveryOptimization")) {
        New-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\DeliveryOptimization" | Out-Null
    }
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\DeliveryOptimization" -Name "SystemSettingsDownloadMode" -Type DWord -Value 3

    # Unrestrict Windows Update P2P
    # Remove-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\DeliveryOptimization\Config" -Name "DODownloadMode"
    # Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\DeliveryOptimization" -Name "SystemSettingsDownloadMode"

    # Remove AutoLogger file and restrict directory
    Write-Host "Removing AutoLogger file and restricting directory..."
    $autoLoggerDir = "$env:PROGRAMDATA\Microsoft\Diagnosis\ETLLogs\AutoLogger"
    If (Test-Path "$autoLoggerDir\AutoLogger-Diagtrack-Listener.etl") {
        Remove-Item "$autoLoggerDir\AutoLogger-Diagtrack-Listener.etl"
    }
    icacls $autoLoggerDir /deny SYSTEM:`(OI`)`(CI`)F | Out-Null

    # Unrestrict AutoLogger directory
    # $autoLoggerDir = "$env:PROGRAMDATA\Microsoft\Diagnosis\ETLLogs\AutoLogger"
    # icacls $autoLoggerDir /grant:r SYSTEM:`(OI`)`(CI`)F | Out-Null

    # Stop and disable Diagnostics Tracking Service
    Write-Host "Stopping and disabling Diagnostics Tracking Service..."
    Stop-Service "DiagTrack"
    Set-Service "DiagTrack" -StartupType Disabled

    # Enable and start Diagnostics Tracking Service
    # Set-Service "DiagTrack" -StartupType Automatic
    # Start-Service "DiagTrack"

    # Stop and disable WAP Push Service
    Write-Host "Stopping and disabling WAP Push Service..."
    Stop-Service "dmwappushservice"
    Set-Service "dmwappushservice" -StartupType Disabled

    # Enable and start WAP Push Service
    # Set-Service "dmwappushservice" -StartupType Automatic
    # Start-Service "dmwappushservice"
    # Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\dmwappushservice" -Name "DelayedAutoStart" -Type DWord -Value 1

	#From Remove Windows10 Bloatware
	Write-Host "Adding Registry key to prevent bloatware apps from returning"
	#Prevents bloatware applications from returning
	$registryPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent"
	If (!(Test-Path $registryPath)) {
		Mkdir $registryPath
		New-ItemProperty $registryPath DisableWindowsConsumerFeatures -Value 1 
	}          

	#From Remove Windows10 Bloatware
	Write-Host "Setting Mixed Reality Portal value to 0 so that you can uninstall it in Settings"
	$Holo = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Holographic'    
	If (Test-Path $Holo) {
		Set-ItemProperty $Holo FirstRunSucceeded -Value 0
	}
      
	#From Remove Windows10 Bloatware
	#Loads the registry keys/values below into the NTUSER.DAT file which prevents the apps from redownloading. Credit to a60wattfish
	reg load HKU\Default_User C:\Users\Default\NTUSER.DAT
	Set-ItemProperty -Path Registry::HKU\Default_User\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager -Name SystemPaneSuggestionsEnabled -Value 0
	Set-ItemProperty -Path Registry::HKU\Default_User\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager -Name PreInstalledAppsEnabled -Value 0
	Set-ItemProperty -Path Registry::HKU\Default_User\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager -Name OemPreInstalledAppsEnabled -Value 0
	reg unload HKU\Default_User
      
 	#From Remove Windows10 Bloatware
	#Disables scheduled tasks that are considered unnecessary 
	Write-Host "Disabling scheduled tasks"
	#Get-ScheduledTask -TaskName XblGameSaveTaskLogon | Disable-ScheduledTask
	Get-ScheduledTask -TaskName XblGameSaveTask | Disable-ScheduledTask
	Get-ScheduledTask -TaskName Consolidator | Disable-ScheduledTask
	Get-ScheduledTask -TaskName UsbCeip | Disable-ScheduledTask
	Get-ScheduledTask -TaskName DmClient | Disable-ScheduledTask
	Get-ScheduledTask -TaskName DmClientOnScenarioDownload | Disable-ScheduledTask


    ##########
    # Service Tweaks
    ##########

    # Lower UAC level
    # Write-Host "Lowering UAC level..."
    # Set-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System" -Name "ConsentPromptBehaviorAdmin" -Type DWord -Value 0
    # Set-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System" -Name "PromptOnSecureDesktop" -Type DWord -Value 0

    # Raise UAC level
    # Set-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System" -Name "ConsentPromptBehaviorAdmin" -Type DWord -Value 5
    # Set-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System" -Name "PromptOnSecureDesktop" -Type DWord -Value 1

    # Enable sharing mapped drives between users
    # Write-Host "Enabling sharing mapped drives between users..."
    # Set-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System" -Name "EnableLinkedConnections" -Type DWord -Value 1

    # Disable sharing mapped drives between users
    # Remove-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System" -Name "EnableLinkedConnections"

    # Disable Firewall
    # Write-Host "Disabling Firewall..."
    # Set-NetFirewallProfile -Profile * -Enabled False

    # Enable Firewall
    Set-NetFirewallProfile -Profile * -Enabled True

    # Disable Windows Defender
    # Write-Host "Disabling Windows Defender..."
    # Set-ItemProperty -Path "HKLM:\Software\Policies\Microsoft\Windows Defender" -Name "DisableAntiSpyware" -Type DWord -Value 1

    # Enable Windows Defender
    # Remove-ItemProperty -Path "HKLM:\Software\Policies\Microsoft\Windows Defender" -Name "DisableAntiSpyware"

    # Disable Windows Update automatic restart
    Write-Host "Disabling Windows Update automatic restart..."
    Set-ItemProperty -Path "HKLM:\Software\Microsoft\WindowsUpdate\UX\Settings" -Name "UxOption" -Type DWord -Value 1

    # Enable Windows Update automatic restart
    # Set-ItemProperty -Path "HKLM:\Software\Microsoft\WindowsUpdate\UX\Settings" -Name "UxOption" -Type DWord -Value 0

    # Stop and disable Home Groups services
    # Write-Host "Stopping and disabling Home Groups services..."
    # Stop-Service "HomeGroupListener"
    # Set-Service "HomeGroupListener" -StartupType Disabled
    # Stop-Service "HomeGroupProvider"
    # Set-Service "HomeGroupProvider" -StartupType Disabled

    # Enable and start Home Groups services
    # Set-Service "HomeGroupListener" -StartupType Manual
    # Set-Service "HomeGroupProvider" -StartupType Manual
    # Start-Service "HomeGroupProvider"

    # Disable Remote Assistance
    # Write-Host "Disabling Remote Assistance..."
    # Set-ItemProperty -Path "HKLM:\System\CurrentControlSet\Control\Remote Assistance" -Name "fAllowToGetHelp" -Type DWord -Value 0

    # Enable Remote Assistance
    # Set-ItemProperty -Path "HKLM:\System\CurrentControlSet\Control\Remote Assistance" -Name "fAllowToGetHelp" -Type DWord -Value 1

    # Enable Remote Desktop w/o Network Level Authentication
    # Write-Host "Enabling Remote Desktop w/o Network Level Authentication..."
    # Set-ItemProperty -Path "HKLM:\System\CurrentControlSet\Control\Terminal Server" -Name "fDenyTSConnections" -Type DWord -Value 0
    # Set-ItemProperty -Path "HKLM:\System\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp" -Name "UserAuthentication" -Type DWord -Value 0

    # Disable Remote Desktop
    Set-ItemProperty -Path "HKLM:\System\CurrentControlSet\Control\Terminal Server" -Name "fDenyTSConnections" -Type DWord -Value 1
    Set-ItemProperty -Path "HKLM:\System\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp" -Name "UserAuthentication" -Type DWord -Value 1



    ##########
    # UI Tweaks
    ##########

    # Disable Action Center
    Write-Host "Disabling Action Center..."
    If (!(Test-Path "HKCU:\Software\Policies\Microsoft\Windows\Explorer")) {
      New-Item -Path "HKCU:\Software\Policies\Microsoft\Windows\Explorer" | Out-Null
    }
    Set-ItemProperty -Path "HKCU:\Software\Policies\Microsoft\Windows\Explorer" -Name "DisableNotificationCenter" -Type DWord -Value 1
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\PushNotifications" -Name "ToastEnabled" -Type DWord -Value 0

    # Enable Action Center
    # Remove-ItemProperty -Path "HKCU:\Software\Policies\Microsoft\Windows\Explorer" -Name "DisableNotificationCenter"
    # Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\PushNotifications" -Name "ToastEnabled"

    # Disable Lock screen
    #Write-Host "Disabling Lock screen..."
    #If (!(Test-Path "HKLM:\Software\Policies\Microsoft\Windows\Personalization")) {
    #  New-Item -Path "HKLM:\Software\Policies\Microsoft\Windows\Personalization" | Out-Null
    #}
    #Set-ItemProperty -Path "HKLM:\Software\Policies\Microsoft\Windows\Personalization" -Name "NoLockScreen" -Type DWord -Value 1

    # Enable Lock screen
    # Remove-ItemProperty -Path "HKLM:\Software\Policies\Microsoft\Windows\Personalization" -Name "NoLockScreen"

    # Disable Autoplay
    Write-Host "Disabling Autoplay..."
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\AutoplayHandlers" -Name "DisableAutoplay" -Type DWord -Value 1

    # Enable Autoplay
    # Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\AutoplayHandlers" -Name "DisableAutoplay" -Type DWord -Value 0

    # Disable Autorun for all drives
     Write-Host "Disabling Autorun for all drives..."
     If (!(Test-Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer")) {
       New-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer" | Out-Null
    }
     Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer" -Name "NoDriveTypeAutoRun" -Type DWord -Value 255

    # Enable Autorun
    # Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer" -Name "NoDriveTypeAutoRun"

    #Disable Sticky keys prompt
    Write-Host "Disabling Sticky keys prompt..."
    Set-ItemProperty -Path "HKCU:\Control Panel\Accessibility\StickyKeys" -Name "Flags" -Type String -Value "506"

    # Enable Sticky keys prompt
    # Set-ItemProperty -Path "HKCU:\Control Panel\Accessibility\StickyKeys" -Name "Flags" -Type String -Value "510"

    # Hide Search button / box
    # Write-Host "Hiding Search Box / Button..."
    # Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search" -Name "SearchboxTaskbarMode" -Type DWord -Value 0

    # Show Search button / box
    # Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search" -Name "SearchboxTaskbarMode"

    # Hide Task View button
    # Write-Host "Hiding Task View button..."
    # Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "ShowTaskViewButton" -Type DWord -Value 0

    # Show Task View button
    Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "ShowTaskViewButton"

    # Show small icons in taskbar
    # Write-Host "Showing small icons in taskbar..."
    # Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "TaskbarSmallIcons" -Type DWord -Value 1

    # Show large icons in taskbar
    # Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "TaskbarSmallIcons"

    # Show titles in taskbar
    # Write-Host "Showing titles in taskbar..."
    # Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "TaskbarGlomLevel" -Type DWord -Value 1

    # Hide titles in taskbar
    # Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "TaskbarGlomLevel"

    # Show all tray icons
    # Write-Host "Showing all tray icons..."
    # Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer" -Name "EnableAutoTray" -Type DWord -Value 0

    # Hide tray icons as needed
    # Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer" -Name "EnableAutoTray"

    # Show known file extensions
    Write-Host "Showing known file extensions..."
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "HideFileExt" -Type DWord -Value 0

    # Hide known file extensions
    # Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "HideFileExt" -Type DWord -Value 1

    # Show hidden files
    #Write-Host "Showing hidden files..."
    #Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "Hidden" -Type DWord -Value 1

    # Hide hidden files
    # Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "Hidden" -Type DWord -Value 2

    # Change default Explorer view to "Computer"
    Write-Host "Changing default Explorer view to `"Computer`"..."
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "LaunchTo" -Type DWord -Value 1

    # Change default Explorer view to "Quick Access"
    # Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "LaunchTo"

    # Show Computer shortcut on desktop
    # Write-Host "Showing Computer shortcut on desktop..."
    # If (!(Test-Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\ClassicStartMenu")) {
    #   New-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\ClassicStartMenu" | Out-Null
    # }
    # Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\ClassicStartMenu" -Name "{20D04FE0-3AEA-1069-A2D8-08002B30309D}" -Type DWord -Value 0
    # Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel" -Name "{20D04FE0-3AEA-1069-A2D8-08002B30309D}" -Type DWord -Value 0

    # Hide Computer shortcut from desktop
    # Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\ClassicStartMenu" -Name "{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
    # Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel" -Name "{20D04FE0-3AEA-1069-A2D8-08002B30309D}"

    # Remove Desktop icon from computer namespace
    # Write-Host "Removing Desktop icon from computer namespace..."
    # Remove-Item -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}" -Recurse -ErrorAction SilentlyContinue

    # Add Desktop icon to computer namespace
    # New-Item -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}"

    # Remove Documents icon from computer namespace
    # Write-Host "Removing Documents icon from computer namespace..."
    # Remove-Item -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{d3162b92-9365-467a-956b-92703aca08af}" -Recurse -ErrorAction SilentlyContinue
    # Remove-Item -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{A8CDFF1C-4878-43be-B5FD-F8091C1C60D0}" -Recurse -ErrorAction SilentlyContinue

    # Add Documents icon to computer namespace
    # New-Item -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{d3162b92-9365-467a-956b-92703aca08af}"
    # New-Item -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{A8CDFF1C-4878-43be-B5FD-F8091C1C60D0}"

    # Remove Downloads icon from computer namespace
    # Write-Host "Removing Downloads icon from computer namespace..."
    # Remove-Item -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{088e3905-0323-4b02-9826-5d99428e115f}" -Recurse -ErrorAction SilentlyContinue
    # Remove-Item -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{374DE290-123F-4565-9164-39C4925E467B}" -Recurse -ErrorAction SilentlyContinue

    # Add Downloads icon to computer namespace
    # New-Item -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{088e3905-0323-4b02-9826-5d99428e115f}"
    # New-Item -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{374DE290-123F-4565-9164-39C4925E467B}"

    # Remove Music icon from computer namespace
    # Write-Host "Removing Music icon from computer namespace..."
    # Remove-Item -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{3dfdf296-dbec-4fb4-81d1-6a3438bcf4de}" -Recurse -ErrorAction SilentlyContinue
    # Remove-Item -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{1CF1260C-4DD0-4ebb-811F-33C572699FDE}" -Recurse -ErrorAction SilentlyContinue

    # Add Music icon to computer namespace
    # New-Item -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{3dfdf296-dbec-4fb4-81d1-6a3438bcf4de}"
    # New-Item -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{1CF1260C-4DD0-4ebb-811F-33C572699FDE}"

    # Remove Pictures icon from computer namespace
    #Write-Host "Removing Pictures icon from computer namespace..."
    #Remove-Item -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{24ad3ad4-a569-4530-98e1-ab02f9417aa8}" -Recurse -ErrorAction SilentlyContinue
    #Remove-Item -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{3ADD1653-EB32-4cb0-BBD7-DFA0ABB5ACCA}" -Recurse -ErrorAction SilentlyContinue

    # Add Pictures icon to computer namespace
    # New-Item -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{24ad3ad4-a569-4530-98e1-ab02f9417aa8}"
    # New-Item -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{3ADD1653-EB32-4cb0-BBD7-DFA0ABB5ACCA}"

    # Remove Videos icon from computer namespace
    #Write-Host "Removing Videos icon from computer namespace..."
    #Remove-Item -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{f86fa3ab-70d2-4fc7-9c99-fcbf05467f3a}" -Recurse -ErrorAction SilentlyContinue
    #Remove-Item -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{A0953C92-50DC-43bf-BE83-3742FED03C9C}" -Recurse -ErrorAction SilentlyContinue

    # Add Videos icon to computer namespace
    # New-Item -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{f86fa3ab-70d2-4fc7-9c99-fcbf05467f3a}"
    # New-Item -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{A0953C92-50DC-43bf-BE83-3742FED03C9C}"

    ## Add secondary en-US keyboard
    #Write-Host "Adding secondary en-US keyboard..."
    #$langs = Get-WinUserLanguageList
    #$langs.Add("en-US")
    #Set-WinUserLanguageList $langs -Force

    # Remove secondary en-US keyboard
    # $langs = Get-WinUserLanguageList
    # Set-WinUserLanguageList ($langs | ? {$_.LanguageTag -ne "en-US"}) -Force



    ##########
    # Remove unwanted applications
    ##########

    # Disable OneDrive
    # Write-Host "Disabling OneDrive..."
    # If (!(Test-Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive")) {
    #     New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive" | Out-Null
    # }
    # Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive" -Name "DisableFileSyncNGSC" -Type DWord -Value 1

    # Enable OneDrive
    # Remove-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive" -Name "DisableFileSyncNGSC"

    # Uninstall OneDrive
    # Write-Host "Uninstalling OneDrive..."
    # Stop-Process -Name OneDrive -ErrorAction SilentlyContinue
    # Start-Sleep -s 3
    # $onedrive = "$env:SYSTEMROOT\SysWOW64\OneDriveSetup.exe"
    # If (!(Test-Path $onedrive)) {
    #     $onedrive = "$env:SYSTEMROOT\System32\OneDriveSetup.exe"
    # }
    # Start-Process $onedrive "/uninstall" -NoNewWindow -Wait
    # Start-Sleep -s 3
    # Stop-Process -Name explorer -ErrorAction SilentlyContinue
    # Start-Sleep -s 3
    # Remove-Item "$env:USERPROFILE\OneDrive" -Force -Recurse -ErrorAction SilentlyContinue
    # Remove-Item "$env:LOCALAPPDATA\Microsoft\OneDrive" -Force -Recurse -ErrorAction SilentlyContinue
    # Remove-Item "$env:PROGRAMDATA\Microsoft OneDrive" -Force -Recurse -ErrorAction SilentlyContinue
    # If (Test-Path "$env:SYSTEMDRIVE\OneDriveTemp") {
    #     Remove-Item "$env:SYSTEMDRIVE\OneDriveTemp" -Force -Recurse -ErrorAction SilentlyContinue
    # }
    # If (!(Test-Path "HKCR:")) {
    #     New-PSDrive -Name HKCR -PSProvider Registry -Root HKEY_CLASSES_ROOT | Out-Null
    # }
    # Remove-Item -Path "HKCR:\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}" -Recurse -ErrorAction SilentlyContinue
    # Remove-Item -Path "HKCR:\Wow6432Node\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}" -Recurse -ErrorAction SilentlyContinue

    # Install OneDrive
    # $onedrive = "$env:SYSTEMROOT\SysWOW64\OneDriveSetup.exe"
    # If (!(Test-Path $onedrive)) {
    #   $onedrive = "$env:SYSTEMROOT\System32\OneDriveSetup.exe"
    # }
    # Start-Process $onedrive -NoNewWindow

	## ## ## ## ## ## ## ## ## ##
	#From "Remove Windows10Bloatware.ps1" script
            $Bloatware = @(
    
                #Unnecessary Windows 10 AppX Apps
                "Microsoft.BingNews"
                "Microsoft.GetHelp"
                "Microsoft.Getstarted"
                #"Microsoft.Messaging"
                "Microsoft.Microsoft3DViewer"
                "Microsoft.MicrosoftOfficeHub"
                "Microsoft.MicrosoftSolitaireCollection"
                "Microsoft.NetworkSpeedTest"
                "Microsoft.News"
                #"Microsoft.Office.Lens"
                #"Microsoft.Office.OneNote"
                #"Microsoft.Office.Sway"
                "Microsoft.OneConnect"
                "Microsoft.People"
                "Microsoft.Print3D"
                "Microsoft.RemoteDesktop"
                #"Microsoft.SkypeApp"
                "Microsoft.StorePurchaseApp"
                "Microsoft.Office.Todo.List"
                "Microsoft.Whiteboard"
                "Microsoft.WindowsAlarms"
                #"Microsoft.WindowsCamera"
                #"microsoft.windowscommunicationsapps"
                "Microsoft.WindowsFeedbackHub"
                "Microsoft.WindowsMaps"
                #"Microsoft.WindowsSoundRecorder"
                "Microsoft.Xbox.TCUI"
                "Microsoft.XboxApp"
                "Microsoft.XboxGameOverlay"
                "Microsoft.XboxIdentityProvider"
                "Microsoft.XboxSpeechToTextOverlay"
                "Microsoft.ZuneMusic"
                "Microsoft.ZuneVideo"

				#From MatStocks PC Build Script
				"Microsoft.3DBuilder"
				"Microsoft.WindowsPhone"
				#"Microsoft.AppConnector"
				#"Microsoft.ConnectivityStore"
				#"Microsoft.CommsPhone"

				
				
                #Sponsored Windows 10 AppX Apps
                #Add sponsored/featured apps to remove in the "*AppName*" format
                "*EclipseManager*"
                "*ActiproSoftwareLLC*"
                "*AdobeSystemsIncorporated.AdobePhotoshopExpress*"
                "*Duolingo-LearnLanguagesforFree*"
                "*PandoraMediaInc*"
                "*CandyCrush*"
                "*Wunderlist*"
                "*Flipboard*"
                "*Twitter*"
                "*Facebook*"
                "*Spotify*"
                "*Minecraft*"
                "*Royal Revolt*"
                "*Sway*"
                "*Dolby*"
                "*Windows.CBSPreview*"
                
                #Optional: Typically not removed but you can if you need to for some reason
                #"*Microsoft.Advertising.Xaml_10.1712.5.0_x64__8wekyb3d8bbwe*"
                #"*Microsoft.Advertising.Xaml_10.1712.5.0_x86__8wekyb3d8bbwe*"
                #"*Microsoft.BingWeather*"
                #"*Microsoft.MSPaint*"
                #"*Microsoft.MicrosoftStickyNotes*"
                #"*Microsoft.Windows.Photos*"
                #"*Microsoft.WindowsCalculator*"
                #"*Microsoft.WindowsStore*"
            )
            foreach ($Bloat in $Bloatware) {
                Get-AppxPackage -Name $Bloat| Remove-AppxPackage
                Get-AppxProvisionedPackage -Online | Where-Object DisplayName -like $Bloat | Remove-AppxProvisionedPackage -Online
                Write-Host "Trying to remove $Bloat."
                Write-Host "Bloatware removed!"
            }
	## ## ## ## ## ## ## ## ## ##
	

#	Replaced by previous section of code	
#	# Uninstall default Microsoft applications
#    Write-Host "Uninstalling default Microsoft applications..."
#    Get-AppxPackage "Microsoft.3DBuilder" | Remove-AppxPackage
#    Get-AppxPackage "Microsoft.BingFinance" | Remove-AppxPackage
#    Get-AppxPackage "Microsoft.BingNews" | Remove-AppxPackage
#    Get-AppxPackage "Microsoft.BingSports" | Remove-AppxPackage
#    Get-AppxPackage "Microsoft.BingWeather" | Remove-AppxPackage
#    Get-AppxPackage "Microsoft.Getstarted" | Remove-AppxPackage
#    Get-AppxPackage "Microsoft.MicrosoftOfficeHub" | Remove-AppxPackage
#    Get-AppxPackage "Microsoft.MicrosoftSolitaireCollection" | Remove-AppxPackage
#    # Get-AppxPackage "Microsoft.Office.OneNote" | Remove-AppxPackage
#    Get-AppxPackage "Microsoft.People" | Remove-AppxPackage
#    # Get-AppxPackage "Microsoft.SkypeApp" | Remove-AppxPackage
#    Get-AppxPackage "Microsoft.Windows.Photos" | Remove-AppxPackage
#    Get-AppxPackage "Microsoft.WindowsAlarms" | Remove-AppxPackage
#    # Get-AppxPackage "Microsoft.WindowsCamera" | Remove-AppxPackage
#    # Get-AppxPackage "microsoft.windowscommunicationsapps" | Remove-AppxPackage
#    Get-AppxPackage "Microsoft.WindowsMaps" | Remove-AppxPackage
#    Get-AppxPackage "Microsoft.WindowsPhone" | Remove-AppxPackage
#    # Get-AppxPackage "Microsoft.WindowsSoundRecorder" | Remove-AppxPackage
#    Get-AppxPackage "Microsoft.XboxApp" | Remove-AppxPackage
#    Get-AppxPackage "Microsoft.ZuneMusic" | Remove-AppxPackage
#    Get-AppxPackage "Microsoft.ZuneVideo" | Remove-AppxPackage
#    # Get-AppxPackage "Microsoft.AppConnector" | Remove-AppxPackage
#    # Get-AppxPackage "Microsoft.ConnectivityStore" | Remove-AppxPackage
#    # Get-AppxPackage "Microsoft.Office.Sway" | Remove-AppxPackage
#    # Get-AppxPackage "Microsoft.Messaging" | Remove-AppxPackage
#    # Get-AppxPackage "Microsoft.CommsPhone" | Remove-AppxPackage
#    # Get-AppxPackage "9E2F88E3.Twitter" | Remove-AppxPackage
#    Get-AppxPackage "king.com.CandyCrushSodaSaga" | Remove-AppxPackage
#    Get-AppxPackage "king.com.CandyCrushSaga" | Remove-AppxPackage
#    Get-AppxPackage "king.com.CandyCrushFriends" | Remove-AppxPackage

    # Install default Microsoft applications
    # Add-AppxPackage -DisableDevelopmentMode -Register "$($(Get-AppXPackage -AllUsers "Microsoft.3DBuilder").InstallLocation)\AppXManifest.xml"
    # Add-AppxPackage -DisableDevelopmentMode -Register "$($(Get-AppXPackage -AllUsers "Microsoft.BingFinance").InstallLocation)\AppXManifest.xml"
    # Add-AppxPackage -DisableDevelopmentMode -Register "$($(Get-AppXPackage -AllUsers "Microsoft.BingNews").InstallLocation)\AppXManifest.xml"
    # Add-AppxPackage -DisableDevelopmentMode -Register "$($(Get-AppXPackage -AllUsers "Microsoft.BingSports").InstallLocation)\AppXManifest.xml"
    # Add-AppxPackage -DisableDevelopmentMode -Register "$($(Get-AppXPackage -AllUsers "Microsoft.BingWeather").InstallLocation)\AppXManifest.xml"
    # Add-AppxPackage -DisableDevelopmentMode -Register "$($(Get-AppXPackage -AllUsers "Microsoft.Getstarted").InstallLocation)\AppXManifest.xml"
    # Add-AppxPackage -DisableDevelopmentMode -Register "$($(Get-AppXPackage -AllUsers "Microsoft.MicrosoftOfficeHub").InstallLocation)\AppXManifest.xml"
    # Add-AppxPackage -DisableDevelopmentMode -Register "$($(Get-AppXPackage -AllUsers "Microsoft.MicrosoftSolitaireCollection").InstallLocation)\AppXManifest.xml"
    # Add-AppxPackage -DisableDevelopmentMode -Register "$($(Get-AppXPackage -AllUsers "Microsoft.Office.OneNote").InstallLocation)\AppXManifest.xml"
    # Add-AppxPackage -DisableDevelopmentMode -Register "$($(Get-AppXPackage -AllUsers "Microsoft.People").InstallLocation)\AppXManifest.xml"
    # Add-AppxPackage -DisableDevelopmentMode -Register "$($(Get-AppXPackage -AllUsers "Microsoft.SkypeApp").InstallLocation)\AppXManifest.xml"
    # Add-AppxPackage -DisableDevelopmentMode -Register "$($(Get-AppXPackage -AllUsers "Microsoft.Windows.Photos").InstallLocation)\AppXManifest.xml"
    # Add-AppxPackage -DisableDevelopmentMode -Register "$($(Get-AppXPackage -AllUsers "Microsoft.WindowsAlarms").InstallLocation)\AppXManifest.xml"
    # Add-AppxPackage -DisableDevelopmentMode -Register "$($(Get-AppXPackage -AllUsers "Microsoft.WindowsCamera").InstallLocation)\AppXManifest.xml"
    # Add-AppxPackage -DisableDevelopmentMode -Register "$($(Get-AppXPackage -AllUsers "Microsoft.windowscommunicationsapps").InstallLocation)\AppXManifest.xml"
    # Add-AppxPackage -DisableDevelopmentMode -Register "$($(Get-AppXPackage -AllUsers "Microsoft.WindowsMaps").InstallLocation)\AppXManifest.xml"
    # Add-AppxPackage -DisableDevelopmentMode -Register "$($(Get-AppXPackage -AllUsers "Microsoft.WindowsPhone").InstallLocation)\AppXManifest.xml"
    # Add-AppxPackage -DisableDevelopmentMode -Register "$($(Get-AppXPackage -AllUsers "Microsoft.WindowsSoundRecorder").InstallLocation)\AppXManifest.xml"
    # Add-AppxPackage -DisableDevelopmentMode -Register "$($(Get-AppXPackage -AllUsers "Microsoft.XboxApp").InstallLocation)\AppXManifest.xml"
    # Add-AppxPackage -DisableDevelopmentMode -Register "$($(Get-AppXPackage -AllUsers "Microsoft.ZuneMusic").InstallLocation)\AppXManifest.xml"
    # Add-AppxPackage -DisableDevelopmentMode -Register "$($(Get-AppXPackage -AllUsers "Microsoft.ZuneVideo").InstallLocation)\AppXManifest.xml"
    # Add-AppxPackage -DisableDevelopmentMode -Register "$($(Get-AppXPackage -AllUsers "Microsoft.AppConnector").InstallLocation)\AppXManifest.xml"
    # Add-AppxPackage -DisableDevelopmentMode -Register "$($(Get-AppXPackage -AllUsers "Microsoft.ConnectivityStore").InstallLocation)\AppXManifest.xml"
    # Add-AppxPackage -DisableDevelopmentMode -Register "$($(Get-AppXPackage -AllUsers "Microsoft.Office.Sway").InstallLocation)\AppXManifest.xml"
    # Add-AppxPackage -DisableDevelopmentMode -Register "$($(Get-AppXPackage -AllUsers "Microsoft.Messaging").InstallLocation)\AppXManifest.xml"
    # Add-AppxPackage -DisableDevelopmentMode -Register "$($(Get-AppXPackage -AllUsers "Microsoft.CommsPhone").InstallLocation)\AppXManifest.xml"
    # In case you have removed them for good, you can try to restore the files using installation medium as follows
    # New-Item C:\Mnt -Type Directory | Out-Null
    # dism /Mount-Image /ImageFile:D:\sources\install.wim /index:1 /ReadOnly /MountDir:C:\Mnt
    # robocopy /S /SEC /R:0 "C:\Mnt\Program Files\WindowsApps" "C:\Program Files\WindowsApps"
    # dism /Unmount-Image /Discard /MountDir:C:\Mnt
    # Remove-Item -Path C:\Mnt -Recurse

    # Uninstall Windows Media Player
    # Write-Host "Uninstalling Windows Media Player..."
    # dism /online /Disable-Feature /FeatureName:MediaPlayback /Quiet /NoRestart

    # Install Windows Media Player
    # dism /online /Enable-Feature /FeatureName:MediaPlayback /Quiet /NoRestart

    # Uninstall Work Folders Client
    # Write-Host "Uninstalling Work Folders Client..."
    # dism /online /Disable-Feature /FeatureName:WorkFolders-Client /Quiet /NoRestart

    # Install Work Folders Client
    # dism /online /Enable-Feature /FeatureName:WorkFolders-Client /Quiet /NoRestart

    # Set Photo Viewer as default for bmp, gif, jpg and png
    Write-Host "Setting Photo Viewer as default for bmp, gif, jpg, png and tif..."
    If (!(Test-Path "HKCR:")) {
        New-PSDrive -Name HKCR -PSProvider Registry -Root HKEY_CLASSES_ROOT | Out-Null
    }
    ForEach ($type in @("Paint.Picture", "giffile", "jpegfile", "pngfile")) {
        New-Item -Path $("HKCR:\$type\shell\open") -Force | Out-Null
        New-Item -Path $("HKCR:\$type\shell\open\command") | Out-Null
        Set-ItemProperty -Path $("HKCR:\$type\shell\open") -Name "MuiVerb" -Type ExpandString -Value "@%ProgramFiles%\Windows Photo Viewer\photoviewer.dll,-3043"
        Set-ItemProperty -Path $("HKCR:\$type\shell\open\command") -Name "(Default)" -Type ExpandString -Value "%SystemRoot%\System32\rundll32.exe `"%ProgramFiles%\Windows Photo Viewer\PhotoViewer.dll`", ImageView_Fullscreen %1"
    }

    # Remove or reset default open action for bmp, gif, jpg and png
    # If (!(Test-Path "HKCR:")) {
    #   New-PSDrive -Name HKCR -PSProvider Registry -Root HKEY_CLASSES_ROOT | Out-Null
    # }
    # Remove-Item -Path "HKCR:\Paint.Picture\shell\open" -Recurse
    # Remove-ItemProperty -Path "HKCR:\giffile\shell\open" -Name "MuiVerb"
    # Set-ItemProperty -Path "HKCR:\giffile\shell\open" -Name "CommandId" -Type String -Value "IE.File"
    # Set-ItemProperty -Path "HKCR:\giffile\shell\open\command" -Name "(Default)" -Type String -Value "`"$env:SystemDrive\Program Files\Internet Explorer\iexplore.exe`" %1"
    # Set-ItemProperty -Path "HKCR:\giffile\shell\open\command" -Name "DelegateExecute" -Type String -Value "{17FE9752-0B5A-4665-84CD-569794602F5C}"
    # Remove-Item -Path "HKCR:\jpegfile\shell\open" -Recurse
    # Remove-Item -Path "HKCR:\pngfile\shell\open" -Recurse

    # Show Photo Viewer in "Open with..."
    Write-Host "Showing Photo Viewer in `"Open with...`""
    If (!(Test-Path "HKCR:")) {
        New-PSDrive -Name HKCR -PSProvider Registry -Root HKEY_CLASSES_ROOT | Out-Null
    }
    New-Item -Path "HKCR:\Applications\photoviewer.dll\shell\open\command" -Force | Out-Null
    New-Item -Path "HKCR:\Applications\photoviewer.dll\shell\open\DropTarget" -Force | Out-Null
    Set-ItemProperty -Path "HKCR:\Applications\photoviewer.dll\shell\open" -Name "MuiVerb" -Type String -Value "@photoviewer.dll,-3043"
    Set-ItemProperty -Path "HKCR:\Applications\photoviewer.dll\shell\open\command" -Name "(Default)" -Type ExpandString -Value "%SystemRoot%\System32\rundll32.exe `"%ProgramFiles%\Windows Photo Viewer\PhotoViewer.dll`", ImageView_Fullscreen %1"
    Set-ItemProperty -Path "HKCR:\Applications\photoviewer.dll\shell\open\DropTarget" -Name "Clsid" -Type String -Value "{FFE2A43C-56B9-4bf5-9A79-CC6D4285608A}"

    # Remove Photo Viewer from "Open with..."
    # If (!(Test-Path "HKCR:")) {
    #   New-PSDrive -Name HKCR -PSProvider Registry -Root HKEY_CLASSES_ROOT | Out-Null
    # }
    # Remove-Item -Path "HKCR:\Applications\photoviewer.dll\shell\open" -Recurse

	#From Remove Windows10 Bloatware script
	#Installs .NET 3.5
	Write-Host "Initializing the installation of .NET 3.5..."
	DISM /Online /Enable-Feature /FeatureName:NetFx3 /All
	Write-Host ".NET 3.5 has been successfully installed!"

    }

# Uploads a default layout to all NEW users that log into the system. Effects task bar and start menu
function LayoutDesign {
	$boot=Get-Partition | Where-Object {$_.IsBoot -eq 'True'}
	$OSDISK = $boot.driveletter + ":"

    If (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")) {
        Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
        Exit
    }
    Import-StartLayout -LayoutPath "c:\Pirum\PC-Build-Script-master\LayoutModification.xml" -MountPath "$OSDISK\"
    }
    
function ApplyDefaultApps {
    dism /online /Import-DefaultAppAssociations:c:\Pirum\PC-Build-Script-master\AppAssociations.xml
}

# Custom power profile used for our customers. Ensures systems do not go to sleep.
function Power {
    POWERCFG -DUPLICATESCHEME 381b4222-f694-41f0-9685-ff5bb260df2e 381b4222-f694-41f0-9685-ff5bb260aaaa
    POWERCFG -CHANGENAME 381b4222-f694-41f0-9685-ff5bb260aaaa "Pirum Power Management"
    POWERCFG -SETACTIVE 381b4222-f694-41f0-9685-ff5bb260aaaa
    POWERCFG -Change -monitor-timeout-ac 15
    POWERCFG -CHANGE -monitor-timeout-dc 5
    POWERCFG -CHANGE -disk-timeout-ac 30
    POWERCFG -CHANGE -disk-timeout-dc 5
    POWERCFG -CHANGE -standby-timeout-ac 0
    POWERCFG -CHANGE -standby-timeout-dc 30
    POWERCFG -CHANGE -hibernate-timeout-ac 0
    POWERCFG -CHANGE -hibernate-timeout-dc 0
}

function JoinDomain {
    add-computer -domainname "spcs.local" -OUPath "OU=SPCS Computer Lab,DC=spcs,DC=local" -Credential SPCS\Administrator 
}

function SetTime {
    Set-TimeZone -Id "Eastern Standard Time" 
}

function personalize{
	$boot=Get-Partition | Where-Object {$_.IsBoot -eq 'True'}
	$OSDISK = $boot.driveletter + ":"

	$SID=[System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value

	# OEM bitmap
	$OEMLogo = "OEMLogo.bmp"
	$strPath = "HKLM:\Software\Microsoft\Windows\CurrentVersion\OEMInformation"
	if(Test-Path -Path "$OSDISK\Pirum\defmedia\$OEMLogo") {
		copy-item "$OSDISK\Pirum\defmedia\$OEMLogo" "$OSDISK\windows\system32"
		copy-item "$OSDISK\Pirum\defmedia\$OEMLogo" "$OSDISK\windows\system32\oobe\info"
		Set-ItemProperty -Path $strPath -Name Logo -Value "$OSDISK\Windows\System32\OEMlogo.bmp"
	}

	#set OEM information
	Set-ItemProperty -Path $strPath -Name Manufacturer -Value "Pirum Consulting LLC"
	Set-ItemProperty -Path $strPath -Name SupportPhone -Value "330-597-0450"
	Set-ItemProperty -Path $strPath -Name SupportHours -Value "Mon-Fri 9am-5pm"
	Set-ItemProperty -Path $strPath -Name SupportURL -Value "http://pirumllc.itclientportal.com"
	
	
	# background
	$Wallpaper = "backgroundDefault.jpg"
	$strPath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
	if(Test-Path -Path "$OSDISK\Pirum\media\$Wallpaper") {
		If (-not(Test-Path "c:\windows\system32\oobe\info\backgrounds")){New-item "c:\windows\system32\oobe\info\backgrounds" -type directory}
		copy-item "$OSDISK\Pirum\media\$Wallpaper" "$OSDISK\windows\system32\oobe\info\backgrounds"
		copy-item "$OSDISK\Pirum\media\$Wallpaper" "$OSDISK\Windows\Web\Screen"
		copy-item "$OSDISK\Pirum\media\$Wallpaper" "$OSDISK\Windows\Web\Wallpaper\Windows"

		New-Item -Path $subPath -Name "System" -Force
		Set-ItemProperty -Path $strPath -Name Wallpaper -value "$OSDISK\Windows\Web\Wallpaper\Windows\$Wallpaper"
		Set-ItemProperty -Path $strPath -Name WallpaperStyle -value "2"
	
	}
	
	#lockscreen
	$LockScreen = "lockscreen.jpg"
	if(Test-Path -Path "$OSDISK\Pirum\media\$LockScreen") {
		copy-item "$OSDISK\Pirum\media\$LockScreen" "$OSDISK\Windows\System32"

		$strPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System"
		Set-ItemProperty -Path $strPath -Name UseOEMBackground -value 1

		$strPath = "HKLM:\Software\Policies\Microsoft\Windows"
		if(!(Test-Path "$strPath\Personalization")) {
			New-Item -Path $strPath -Name "Personalization" -Force
		}
		Set-ItemProperty -Path "$strPath\Personalization" -Name LockScreenImage -value "$OSDISK\Windows\System32\$LockScreen"

		$strPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion"
		if(!(Test-Path "$strPath\PersonalizationCSP")) {
			New-Item -Path $strPath -Name "PersonalizationCSP" -Force
		}
		New-ItemProperty -Path $subPath6 -Name LockScreenImageStatus -Value 1 -PropertyType DWORD -Force
		New-ItemProperty -Path $subPath6 -Name LockScreenImagePath -Value "$OSDISK\Windows\System32\$LockScreen" -PropertyType STRING -Force
		New-ItemProperty -Path $subPath6 -Name LockScreenImageUrl -Value "$OSDISK\Windows\System32\$LockScreen" -PropertyType STRING -Force
	}


	# user account pictures
	$sizelist = "32","40","48","96","192","240"
	if(Test-Path -Path "$OSDISK\Pirum\media\user.png"){
		copy-item "$OSDISK\Pirum\media\user.png" "$OSDISK\ProgramData\Microsoft\User Account Pictures"
	}
	if(Test-Path -Path "$OSDISK\Pirum\media\user.bmp"){
		copy-item "$OSDISK\Pirum\media\user.bmp" "$OSDISK\ProgramData\Microsoft\User Account Pictures"
	}

	$strPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AccountPicture\Users"
	New-Item -Path $strPath -Name "$SID" -Force
	$strPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AccountPicture\Users\$SID"
	foreach($size in $sizelist){
		if(Test-Path -Path "$OSDISK\Pirum\media\user-$size.bmp"){
			copy-item "$OSDISK\Pirum\media\user-$size.bmp" "$OSDISK\ProgramData\Microsoft\User Account Pictures"
			Set-ItemProperty -Path $strPath -Name "Image$size" -value "$OSDISK\ProgramData\Microsoft\User Account Pictures\user-$size.bmp"
		}
		if(Test-Path -Path "$OSDISK\Pirum\media\user-$size.png"){
			copy-item "$OSDISK\Pirum\media\user-$size.png" "$OSDISK\ProgramData\Microsoft\User Account Pictures"
			Set-ItemProperty -Path $strPath -Name "Image$size" -value "$OSDISK\ProgramData\Microsoft\User Account Pictures\user-$size.png"
		}
	}
}

function RestartPC{
    ##########
    # Restart
    ##########
    Write-Host
    Write-Host "Press any key to restart your system..." -ForegroundColor Black -BackgroundColor White
    $key = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    Write-Host "Restarting..."
    Restart-Computer
}

function Branding{
#Invoke-WebRequest -Uri "https://pirumllc.com/Pirum+consulting+logo+colors.png" -OutFile "c:\windows\system32\intechlogo.bmp"
copy-item "$OSDISK\Pirum\media\$OEMLogo" "$OSDISK\windows\system32"
New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation" -Name "Manufacturer" -Value "Pirum Consulting LLC"  -PropertyType "String" -Force
}

InstallChoco
InstallApps
ReclaimWindows10
LayoutDesign
ApplyDefaultApps
Power
Branding
SetPCName
SetTime
personalize
#JoinDomain
RestartPC
