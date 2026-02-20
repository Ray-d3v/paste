param(
    [string]$Configuration = "Release",
    [string]$Runtime = "win-x64",
    [string]$Version = "0.1.0",
    [switch]$SkipPublish
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

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$projectPath = Join-Path $repoRoot "PasteWinUI\PasteWinUI.csproj"
$issPath = Join-Path $repoRoot "installer\PasteWinUI.iss"
$publishDir = Join-Path $repoRoot "PasteWinUI\bin\x64\$Configuration\net8.0-windows10.0.19041.0\$Runtime\publish"
$publishExe = Join-Path $publishDir "PasteWinUI.exe"

if (-not $SkipPublish) {
    Write-Host "Publishing $projectPath ..."
    dotnet publish $projectPath -c $Configuration -r $Runtime --self-contained true /p:WindowsAppSDKSelfContained=true
}

if (-not (Test-Path $publishExe)) {
    throw "Publish output was not found: $publishExe"
}

$iscc = Resolve-IsccPath
Write-Host "Using ISCC: $iscc"
Write-Host "Building installer ..."
& $iscc "/DMyAppVersion=$Version" $issPath

if ($LASTEXITCODE -ne 0) {
    throw "ISCC failed. ExitCode=$LASTEXITCODE"
}

$installerPath = Join-Path $repoRoot "installer\output\PasteWinUI-Setup-$Version.exe"
Write-Host "Installer build completed."
Write-Host "Output: $installerPath"
