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

REM ── Set execution policy ──
powershell.exe -NoProfile -Command "Set-ExecutionPolicy RemoteSigned -Scope LocalMachine -Force" >nul 2>&1

REM ── Download latest build script from GitHub ──
echo [....] Downloading latest build script...
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command ^
    "& { Start-Process powershell.exe -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File ""%~dp0download.ps1""' -Verb RunAs -Wait }"
echo [ OK ] Download complete.

REM ── Launch the setup configurator ──
echo [....] Launching setup configurator...
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command ^
    "& { Start-Process powershell.exe -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File ""C:\Pirum\PC-Build-Script-master\PCSetup.ps1""' -Verb RunAs }"

echo [ OK ] Done. This window can be closed.
echo.
pause
endlocal
