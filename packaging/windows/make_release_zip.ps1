param(
    [string]$DistDir = "",
    [string]$Version = "12.1.0"
)

$ErrorActionPreference = "Stop"

function Invoke-WithRetry {
    param(
        [scriptblock]$Action,
        [string]$Description
    )

    for ($attempt = 1; $attempt -le 5; $attempt++) {
        try {
            & $Action
            return
        }
        catch {
            if ($attempt -eq 5) {
                throw "$Description failed after 5 attempts: $_"
            }
            Start-Sleep -Seconds 2
        }
    }
}

$root = Resolve-Path (Join-Path $PSScriptRoot "..\..")
$appName = "Allokering"
$releaseDir = Join-Path $root "release"
$packageDir = Join-Path $releaseDir $appName
$zipPath = Join-Path $releaseDir "$appName-$Version-win64.zip"

if ([string]::IsNullOrWhiteSpace($DistDir)) {
    $distDir = Join-Path $root "dist\$appName"
}
else {
    $distDir = (Resolve-Path $DistDir).Path
}

if (-not (Test-Path $distDir)) {
    throw "Build output not found: $distDir"
}

New-Item -ItemType Directory -Force -Path $releaseDir | Out-Null

Invoke-WithRetry -Description "Stage release folder" -Action {
    if (Test-Path $packageDir) {
        Remove-Item -LiteralPath $packageDir -Recurse -Force
    }
    New-Item -ItemType Directory -Force -Path $packageDir | Out-Null
    Copy-Item -Path (Join-Path $distDir "*") -Destination $packageDir -Recurse -Force

    foreach ($asset in @("Installera Allokering.bat", "install.ps1", "README_ANVANDARE.txt")) {
        Copy-Item -Path (Join-Path $PSScriptRoot $asset) -Destination $packageDir -Force
    }
}

Invoke-WithRetry -Description "Create release zip" -Action {
    if (Test-Path $zipPath) {
        Remove-Item -LiteralPath $zipPath -Force
    }
    Compress-Archive -Path (Join-Path $packageDir "*") -DestinationPath $zipPath -CompressionLevel Optimal
}

Write-Host "Created $zipPath"
Write-Host "Staged $packageDir"
