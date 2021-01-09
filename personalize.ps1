function personalize{
	$boot=Get-Partition | Where-Object {$_.IsBoot -eq 'True'}
	$OSDISK = $boot.driveletter + ":"

	$SID=[System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
	$Wallpaper = "backgroundDefault.jpg"
	$OEMLogo = "OEMLogo.BMP"
	$UserBMP32 = "user32.bmp"
	$UserBMP40 = "user40.bmp"
	$UserBMP48 = "user48.bmp"
	$UserBMP96 = "user96.bmp"
	$UserBMP192 = "user192.bmp"
	$UserBMP240 = "user240.bmp"
	$LockScreen = "lockscreen.jpg"

	# copy the OEM bitmap
	If (-not(Test-Path "c:\windows\system32\oobe\info\backgrounds")){New-item "c:\windows\system32\oobe\info\backgrounds" -type directory}

	copy-item "$OSDISK\Pirum\media\$OEMLogo" "$OSDISK\windows\system32"
	copy-item "$OSDISK\Pirum\media\$UserBMP32" "$OSDISK\ProgramData\Microsoft\User Account Pictures"
	copy-item "$OSDISK\Pirum\media\$UserBMP40" "$OSDISK\ProgramData\Microsoft\User Account Pictures"
	copy-item "$OSDISK\Pirum\media\$UserBMP48" "$OSDISK\ProgramData\Microsoft\User Account Pictures"
	copy-item "$OSDISK\Pirum\media\$UserBMP96" "$OSDISK\ProgramData\Microsoft\User Account Pictures"
	copy-item "$OSDISK\Pirum\media\$UserBMP192" "$OSDISK\ProgramData\Microsoft\User Account Pictures"
	copy-item "$OSDISK\Pirum\media\$UserBMP240" "$OSDISK\ProgramData\Microsoft\User Account Pictures"
	copy-item "$OSDISK\Pirum\media\$OEMLogo" "$OSDISK\windows\system32\oobe\info"
	copy-item "$OSDISK\Pirum\media\$Wallpaper" "$OSDISK\windows\system32\oobe\info\backgrounds"
	copy-item "$OSDISK\Pirum\media\$Wallpaper" "$OSDISK\Windows\Web\Screen"
	copy-item "$OSDISK\Pirum\media\$Wallpaper" "$OSDISK\Windows\Web\Wallpaper\Windows"
	copy-item "$OSDISK\Pirum\media\$LockScreen" "$OSDISK\Windows\System32"

	# make required registry changes
	$strPath = "HKLM:\Software\Microsoft\Windows\CurrentVersion\OEMInformation"
	$strPath2 = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Authentication\LogonUI\Background"
	$subPath3 = "HKLM:\Software\Policies\Microsoft\Windows"
	$strPath3 = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Personalization"
	$subPath4 = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies"
	$strPath4 = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
	$subPath5 = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AccountPicture\Users"
	$strPath5 = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AccountPicture\Users\$SID"
	$subPath6 = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP"

	#set OEM information
	Set-ItemProperty -Path $strPath -Name Logo -Value "$OSDISK\Windows\System32\OEMlogo.bmp"
	Set-ItemProperty -Path $strPath -Name Manufacturer -Value "Pirum Consulting LLC"
	Set-ItemProperty -Path $strPath -Name SupportPhone -Value "330-597-0450"
	Set-ItemProperty -Path $strPath -Name SupportHours -Value "Mon-Fri 9am-5pm"
	Set-ItemProperty -Path $strPath -Name SupportURL -Value "http://pirumllc.itclientportal.com"
	Set-ItemProperty -Path $strPath2 -Name OEMBackground -value 1
	
	#set lockscreen
    if(!(Test-Path $subPath3)) {
		New-Item -Path $subPath3 -Name "Personalization" -Force
	}
	Set-ItemProperty -Path $strPath3 -Name LockScreenImage -value "$OSDISK\Windows\Web\Screen\$Wallpaper"

    if(!(Test-Path $subPath6)) {
        New-Item -Path $subPath6 -Force | Out-Null
    }
	New-ItemProperty -Path $subPath6 -Name LockScreenImageStatus -Value 1 -PropertyType DWORD -Force | Out-Null
	New-ItemProperty -Path $subPath6 -Name LockScreenImagePath -Value "$OSDISK\Windows\System32\$LockScreen" -PropertyType STRING -Force | Out-Null
	New-ItemProperty -Path $subPath6 -Name LockScreenImageUrl -Value "$OSDISK\Windows\System32\$LockScreen" -PropertyType STRING -Force | Out-Null
	

	#set wallpaper
	New-Item -Path $subPath4 -Name "System" -Force
	Set-ItemProperty -Path $strPath4 -Name Wallpaper -value "$OSDISK\Windows\Web\Wallpaper\Windows\$Wallpaper"
	Set-ItemProperty -Path $strPath4 -Name WallpaperStyle -value "2"
	
	#set user icon for lockscreen
	New-Item -Path $subPath5 -Name "$SID" -Force
	Set-ItemProperty -Path $strPath5 -Name Image32 -value "$OSDISK\ProgramData\Microsoft\User Account Pictures\$UserBMP32"
	Set-ItemProperty -Path $strPath5 -Name Image40 -value "$OSDISK\ProgramData\Microsoft\User Account Pictures\$UserBMP40"
	Set-ItemProperty -Path $strPath5 -Name Image48 -value "$OSDISK\ProgramData\Microsoft\User Account Pictures\$UserBMP48"
	Set-ItemProperty -Path $strPath5 -Name Image96 -value "$OSDISK\ProgramData\Microsoft\User Account Pictures\$UserBMP96"
	Set-ItemProperty -Path $strPath5 -Name Image192 -value "$OSDISK\ProgramData\Microsoft\User Account Pictures\$UserBMP192"
	Set-ItemProperty -Path $strPath5 -Name Image240 -value "$OSDISK\ProgramData\Microsoft\User Account Pictures\$UserBMP240"
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

personalize
RestartPC