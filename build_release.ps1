$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $projectRoot

$pyInstallerCmd = "python -m PyInstaller"
$isccCandidates = @(
    "C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
    "C:\Program Files\Inno Setup 6\ISCC.exe"
)
$iscc = $isccCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $iscc) {
    throw "ISCC.exe was not found. Inno Setup 6 is required."
}

Write-Host "Cleaning previous build folders..." -ForegroundColor Cyan
Remove-Item -Recurse -Force "$projectRoot\build" -ErrorAction SilentlyContinue
Remove-Item -Recurse -Force "$projectRoot\dist" -ErrorAction SilentlyContinue
Remove-Item -Recurse -Force "$projectRoot\installer-output" -ErrorAction SilentlyContinue

Write-Host "Building WinSys Guardian Advanced executable..." -ForegroundColor Cyan
cmd /c "$pyInstallerCmd --noconfirm --clean --onedir --windowed --name WGA --icon `"assets\wga-icon.ico`" --add-data `"assets\wga-icon.ico;assets`" --add-data `"installers_manifest.json;.`" --add-data `"version.json;.`" app.py"

Write-Host "Creating portable update package..." -ForegroundColor Cyan
$portableZip = "$projectRoot\installer-output\WGA-portable.zip"
New-Item -ItemType Directory -Force "$projectRoot\installer-output" | Out-Null
Compress-Archive -Path "$projectRoot\dist\WGA\*" -DestinationPath $portableZip -Force

Write-Host "Compiling Inno Setup installer..." -ForegroundColor Cyan
& $iscc "$projectRoot\WGAInstaller.iss"

Write-Host "Done. Check installer-output for the setup and WGA-portable.zip update package." -ForegroundColor Green
