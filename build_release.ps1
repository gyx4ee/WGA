$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $projectRoot

$pyInstallerCandidates = @(
    "$env:APPDATA\Python\Python313\Scripts\pyinstaller.exe",
    "$env:APPDATA\Python\Python312\Scripts\pyinstaller.exe",
    "$env:LOCALAPPDATA\Programs\Python\Python313\Scripts\pyinstaller.exe",
    "$env:LOCALAPPDATA\Programs\Python\Python312\Scripts\pyinstaller.exe"
)
$pyInstallerExe = $pyInstallerCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
$pyInstallerCmd = if ($pyInstallerExe) { "`"$pyInstallerExe`"" } else { "python -m PyInstaller" }
$isccCandidates = @(
    "C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
    "C:\Program Files\Inno Setup 6\ISCC.exe",
    "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe"
)
$iscc = $isccCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1

Write-Host "Cleaning previous build folders..." -ForegroundColor Cyan
Remove-Item -Recurse -Force "$projectRoot\build" -ErrorAction SilentlyContinue
Remove-Item -Recurse -Force "$projectRoot\dist" -ErrorAction SilentlyContinue
Remove-Item -Recurse -Force "$projectRoot\installer-output" -ErrorAction SilentlyContinue

Write-Host "Building WinSys Guardian Advanced executable..." -ForegroundColor Cyan
cmd /c "$pyInstallerCmd --noconfirm --clean --onedir --windowed --name WGA --icon `"assets\wga-icon.ico`" --add-data `"assets;assets`" --add-data `"installers_manifest.json;.`" --add-data `"version.json;.`" --add-data `"third_party\open-shell\portable\PFiles\Open-Shell;third_party\open-shell\PFiles\Open-Shell`" --add-data `"third_party\open-shell\LICENSE.txt;third_party\open-shell`" app.py"

Write-Host "Creating portable update package..." -ForegroundColor Cyan
$portableZip = "$projectRoot\installer-output\WGA-portable.zip"
New-Item -ItemType Directory -Force "$projectRoot\installer-output" | Out-Null
Compress-Archive -Path "$projectRoot\dist\WGA\*" -DestinationPath $portableZip -Force

if ($iscc) {
    Write-Host "Compiling Inno Setup installer..." -ForegroundColor Cyan
    & $iscc "$projectRoot\WGAInstaller.iss"
}
else {
    Write-Host "Inno Setup was not found. Skipping Setup Wizard build." -ForegroundColor Yellow
}

Write-Host "Done. Check dist\WGA for the portable EXE and installer-output for WGA-portable.zip." -ForegroundColor Green
