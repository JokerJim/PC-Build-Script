@echo off
setlocal

REM ============================================================
REM  Pirum Consulting LLC - PC Setup Launcher
REM  Run this file from the USB drive or network share.
REM  It copies media files, downloads the latest build script,
REM  and launches the setup configurator.
REM ============================================================

echo.
echo   Pirum Consulting LLC - PC Setup Launcher
echo   ------------------------------------------
echo.

REM ── Create local Pirum directory structure ──
echo [....] Creating C:\Pirum folder structure...
if not exist "C:\Pirum"          mkdir "C:\Pirum"
if not exist "C:\Pirum\media"    mkdir "C:\Pirum\media"
if not exist "C:\Pirum\defmedia" mkdir "C:\Pirum\defmedia"
if not exist "C:\Pirum\agents"   mkdir "C:\Pirum\agents"
echo [ OK ] Folders ready.

REM ── Copy media files from the same location as this bat file ──
echo [....] Copying media files...
if exist "%~dp0media\*"    xcopy /Y /Q "%~dp0media\*"    "C:\Pirum\media\"
if exist "%~dp0defmedia\*" xcopy /Y /Q "%~dp0defmedia\*" "C:\Pirum\defmedia\"
echo [ OK ] Media files copied.

REM ── Copy agent installers if present alongside this script ──
if exist "%~dp0agents\*" (
    echo [....] Copying agent installers...
    xcopy /Y /Q "%~dp0agents\*" "C:\Pirum\agents\"
    echo [ OK ] Agent installers copied.
)

REM ── Set PowerShell execution policy ──
echo [....] Setting PowerShell execution policy...
powershell.exe -NoProfile -Command "Set-ExecutionPolicy RemoteSigned -Scope LocalMachine -Force" >nul 2>&1
echo [ OK ] Execution policy set.

REM ── Download latest build script from GitHub ──
echo [....] Downloading latest build script from GitHub...
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command ^
    "& { Start-Process powershell.exe -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File ""%~dp0download.ps1""' -Verb RunAs -Wait }"
echo [ OK ] Download complete.

REM ── Brief pause to ensure extraction has finished ──
echo [....] Waiting for extraction to complete...
timeout /t 5 /nobreak >nul
echo [ OK ] Ready.

REM ── Launch the setup configurator ──
echo [....] Launching PCSetup configurator...
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command ^
    "& { Start-Process powershell.exe -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File ""C:\Pirum\PC-Build-Script-master\PCSetup.ps1""' -Verb RunAs }"

echo [ OK ] Configurator launched. This window can be closed.
echo.
pause
endlocal
