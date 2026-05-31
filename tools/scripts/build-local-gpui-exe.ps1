param(
    [switch]$NoLaunch,
    [string]$Profile = "release",
    [string]$OutDir = "artifacts/local-gpui",
    [string]$TargetDir = "artifacts/cargo-target"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Add-WindowsSdkFxcToPath {
    if (Get-Command fxc.exe -ErrorAction SilentlyContinue) {
        return
    }

    $kitsRoot = "${env:ProgramFiles(x86)}\Windows Kits\10\bin"
    if (-not (Test-Path $kitsRoot)) {
        Write-Warning "Windows SDK bin directory was not found; release GPUI builds may fail without fxc.exe."
        return
    }

    $fxc = Get-ChildItem $kitsRoot -Recurse -Filter fxc.exe -ErrorAction SilentlyContinue |
        Where-Object { $_.FullName -match "\\x64\\fxc\.exe$" } |
        Sort-Object FullName -Descending |
        Select-Object -First 1

    if ($fxc) {
        $fxcDir = Split-Path -Parent $fxc.FullName
        $env:PATH = "$fxcDir;$env:PATH"
        Write-Host "Added Windows SDK FXC to PATH: $fxcDir"
    }
    else {
        Write-Warning "fxc.exe was not found under Windows SDK; release GPUI builds may fail."
    }
}

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$workspacePath = Join-Path $repoRoot "src-paste-gpui\Cargo.toml"
$targetDirPath = Join-Path $repoRoot $TargetDir
$publishDir = Join-Path $repoRoot $OutDir
$profileDir = if ($Profile -eq "release") { "release" } else { "debug" }
$builtExe = Join-Path $targetDirPath "$profileDir\PasteWinUI.exe"
$localExe = Join-Path $publishDir "PasteWinUI.exe"

if (-not (Test-Path $workspacePath)) {
    throw "Rust workspace not found: $workspacePath"
}

Write-Host "Stopping existing PasteWinUI processes..."
Get-Process -Name "PasteWinUI" -ErrorAction SilentlyContinue | Stop-Process -Force

New-Item -ItemType Directory -Path $targetDirPath -Force | Out-Null
New-Item -ItemType Directory -Path $publishDir -Force | Out-Null

Add-WindowsSdkFxcToPath

$env:CARGO_TARGET_DIR = $targetDirPath
$cargoArgs = @("build", "--manifest-path", $workspacePath, "-p", "paste-gpui-app")
if ($Profile -eq "release") {
    $cargoArgs += "--release"
}
elseif ($Profile -ne "debug") {
    throw "Unsupported Profile '$Profile'. Use 'debug' or 'release'."
}

Write-Host "Building Rust GPUI EXE..."
& cargo @cargoArgs
if ($LASTEXITCODE -ne 0) {
    throw "cargo build failed. ExitCode=$LASTEXITCODE"
}

if (-not (Test-Path $builtExe)) {
    throw "EXE was not generated: $builtExe"
}

Copy-Item -LiteralPath $builtExe -Destination $localExe -Force
Write-Host "EXE generated: $localExe"

if ($NoLaunch) {
    Write-Host "NoLaunch specified. Skipping launch."
    exit 0
}

Write-Host "Launching EXE..."
Start-Process -FilePath $localExe
