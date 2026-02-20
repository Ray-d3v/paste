param(
    [switch]$NoLaunch,
    [string]$Configuration = "Release",
    [string]$RuntimeIdentifier = "win-x64",
    [string]$OutDir = "artifacts/local/publish"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$projectPath = Join-Path $repoRoot "PasteWinUI/PasteWinUI.csproj"
$publishDir = Join-Path $repoRoot $OutDir
$exePath = Join-Path $publishDir "PasteWinUI.exe"

if (-not (Test-Path $projectPath)) {
    throw "Project not found: $projectPath"
}

Write-Host "Stopping existing PasteWinUI processes..."
Get-Process -Name "PasteWinUI" -ErrorAction SilentlyContinue | Stop-Process -Force

$dotnetCandidates = Get-CimInstance Win32_Process -Filter "Name = 'dotnet.exe'" |
    Where-Object { $_.CommandLine -match "PasteWinUI[\\/]+PasteWinUI\.csproj" }

foreach ($proc in $dotnetCandidates) {
    try {
        Stop-Process -Id $proc.ProcessId -Force -ErrorAction Stop
        Write-Host "Stopped dotnet host process: $($proc.ProcessId)"
    }
    catch {
        Write-Warning "Failed to stop dotnet process $($proc.ProcessId): $($_.Exception.Message)"
    }
}

New-Item -ItemType Directory -Path $publishDir -Force | Out-Null

Write-Host "Publishing EXE..."
dotnet publish $projectPath `
    -c $Configuration `
    -r $RuntimeIdentifier `
    -p:SelfContained=true `
    -p:WindowsAppSDKSelfContained=true `
    -p:OutDir="$publishDir/"

if (-not (Test-Path $exePath)) {
    throw "EXE was not generated: $exePath"
}

Write-Host "EXE generated: $exePath"

if ($NoLaunch) {
    Write-Host "NoLaunch specified. Skipping launch."
    exit 0
}

Write-Host "Launching EXE..."
Start-Process -FilePath $exePath
