$ErrorActionPreference = "Stop"

$appName = "Allokering"
$sourceDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$targetDir = Join-Path $env:LOCALAPPDATA $appName
$exePath = Join-Path $targetDir "$appName.exe"

if (-not (Test-Path (Join-Path $sourceDir "$appName.exe"))) {
    throw "Hittar inte $appName.exe i $sourceDir"
}

function New-AppShortcut {
    param(
        [string]$ShortcutPath,
        [string]$TargetPath,
        [string]$WorkingDirectory
    )

    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($ShortcutPath)
    $shortcut.TargetPath = $TargetPath
    $shortcut.WorkingDirectory = $WorkingDirectory
    $shortcut.Save()
}

New-Item -ItemType Directory -Force -Path $targetDir | Out-Null
Copy-Item -Path (Join-Path $sourceDir "*") -Destination $targetDir -Recurse -Force

$startMenuDir = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\$appName"
New-Item -ItemType Directory -Force -Path $startMenuDir | Out-Null

$desktopDir = [Environment]::GetFolderPath("Desktop")
New-AppShortcut -ShortcutPath (Join-Path $desktopDir "$appName.lnk") -TargetPath $exePath -WorkingDirectory $targetDir
New-AppShortcut -ShortcutPath (Join-Path $startMenuDir "$appName.lnk") -TargetPath $exePath -WorkingDirectory $targetDir

$uninstallScriptPath = Join-Path $targetDir "uninstall.ps1"
$uninstallBatchPath = Join-Path $targetDir "Avinstallera $appName.bat"

$uninstallScript = @'
$ErrorActionPreference = "Stop"

$appName = "Allokering"
$targetDir = Join-Path $env:LOCALAPPDATA $appName
$startMenuDir = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\$appName"
$desktopShortcut = Join-Path ([Environment]::GetFolderPath("Desktop")) "$appName.lnk"

Get-Process -Name $appName -ErrorAction SilentlyContinue | Stop-Process -Force

if (Test-Path $desktopShortcut) {
    Remove-Item -LiteralPath $desktopShortcut -Force
}

if (Test-Path $startMenuDir) {
    Remove-Item -LiteralPath $startMenuDir -Recurse -Force
}

Set-Location $env:TEMP
if (Test-Path $targetDir) {
    Remove-Item -LiteralPath $targetDir -Recurse -Force
}

Write-Host "Allokering ar avinstallerat."
'@

$uninstallBatch = @"
@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0uninstall.ps1"
pause
"@

Set-Content -Path $uninstallScriptPath -Value $uninstallScript -Encoding UTF8
Set-Content -Path $uninstallBatchPath -Value $uninstallBatch -Encoding ASCII
New-AppShortcut -ShortcutPath (Join-Path $startMenuDir "Avinstallera $appName.lnk") -TargetPath $uninstallBatchPath -WorkingDirectory $targetDir

Write-Host "Installerat till $targetDir"
