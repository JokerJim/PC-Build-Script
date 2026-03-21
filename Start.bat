@echo off
setlocal

REM ============================================================
REM  Pirum Consulting LLC - Quick Launch
REM  Run from the extracted GitHub archive folder.
REM  Downloads the latest script package then launches setup.
REM ============================================================

echo.
echo   Pirum Consulting LLC - PC Setup Quick Launch
echo   ----------------------------------------------
echo.

REM -- Set execution policy --
powershell.exe -NoProfile -Command "Set-ExecutionPolicy RemoteSigned -Scope LocalMachine -Force" >nul 2>&1

REM -- Download latest build script from GitHub --
echo [....] Downloading latest build script...
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command ^
    "& { Start-Process powershell.exe -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File ""%~dp0download.ps1""' -Verb RunAs -Wait }"
echo [ OK ] Download complete.

REM -- Launch the setup configurator --
echo [....] Launching setup configurator...
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command ^
    "& { Start-Process powershell.exe -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File ""C:\Pirum\PC-Build-Script-master\PCSetup.ps1""' -Verb RunAs }"

echo [ OK ] Done. Setup is launching.
echo.

REM -- 30-second countdown then auto-close; any key exits immediately --
powershell.exe -NoProfile -Command "for ($i=30;$i-gt 0;$i--){Write-Host \"`rClosing in $i seconds... (press any key to close now)  \" -NoNewline;if([Console]::KeyAvailable){[Console]::ReadKey($true)|Out-Null;break};Start-Sleep -Milliseconds 1000};Write-Host ''"

endlocal

REM ============================================================
REM  VERSION HISTORY
REM ============================================================
REM
REM  v0.2  - Replaced pause with 30-second auto-close countdown.
REM          Pressing any key closes immediately.
REM
REM  v0.1  - Initial version. Sets execution policy, runs
REM          download.ps1 elevated, launches PCSetup.ps1 elevated.
