param(
    [string]$Version = "12.1.0"
)

$ErrorActionPreference = "Stop"

$scriptPath = Join-Path $PSScriptRoot "Allokering.iss"
$candidates = @()

$pathCommand = Get-Command "iscc.exe" -ErrorAction SilentlyContinue
if ($pathCommand) {
    $candidates += $pathCommand.Source
}

if ($env:ProgramFiles) {
    $candidates += Join-Path $env:ProgramFiles "Inno Setup 6\ISCC.exe"
}

if ($env:LOCALAPPDATA) {
    $candidates += Join-Path $env:LOCALAPPDATA "Programs\Inno Setup 6\ISCC.exe"
}

$programFilesX86 = [Environment]::GetEnvironmentVariable("ProgramFiles(x86)")
if ($programFilesX86) {
    $candidates += Join-Path $programFilesX86 "Inno Setup 6\ISCC.exe"
}

$iscc = $candidates | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1

if (-not $iscc) {
    Write-Host "Inno Setup compiler not found. Skipping Setup.exe."
    Write-Host "Installera Inno Setup 6 och kor build_windows.bat igen for att skapa release\Allokering-12.1.0-Setup.exe."
    exit 0
}

& $iscc "/DMyAppVersion=$Version" $scriptPath
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}
