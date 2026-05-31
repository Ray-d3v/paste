param(
    [string]$Version = "0.1.0",
    [switch]$SkipBuild,
    [string]$Profile = "release"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Resolve-IsccPath {
    $command = Get-Command iscc -ErrorAction SilentlyContinue
    if ($command) {
        return $command.Source
    }

    $candidates = @(
        "$env:ProgramFiles(x86)\Inno Setup 6\ISCC.exe",
        "$env:ProgramFiles\Inno Setup 6\ISCC.exe"
    )

    foreach ($candidate in $candidates) {
        if (Test-Path $candidate) {
            return $candidate
        }
    }

    throw "Inno Setup 6 was not found. Install ISCC.exe and add it to PATH or the default install directory."
}

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$buildScript = Join-Path $repoRoot "tools\scripts\build-local-gpui-exe.ps1"
$issPath = Join-Path $repoRoot "deploy\installer\PasteWinUI.GPUI.iss"
$publishExe = Join-Path $repoRoot "artifacts\local-gpui\PasteWinUI.exe"

if (-not $SkipBuild) {
    Write-Host "Building Rust GPUI app..."
    & powershell -ExecutionPolicy Bypass -File $buildScript -NoLaunch -Profile $Profile
    if ($LASTEXITCODE -ne 0) {
        throw "Rust GPUI build failed. ExitCode=$LASTEXITCODE"
    }
}

if (-not (Test-Path $publishExe)) {
    throw "Rust GPUI EXE was not found: $publishExe"
}

$iscc = Resolve-IsccPath
Write-Host "Using ISCC: $iscc"
Write-Host "Building GPUI installer..."
& $iscc "/DMyAppVersion=$Version" $issPath

if ($LASTEXITCODE -ne 0) {
    throw "ISCC failed. ExitCode=$LASTEXITCODE"
}

$installerPath = Join-Path $repoRoot "deploy\installer\output\PasteWinUI-GPUI-Setup-$Version.exe"
Write-Host "Installer build completed."
Write-Host "Output: $installerPath"
