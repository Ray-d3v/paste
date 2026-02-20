param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$GitPushArgs
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$buildScript = Join-Path $PSScriptRoot "build-local-exe.ps1"
$exePath = Join-Path $repoRoot "artifacts/local/publish/PasteWinUI.exe"

Push-Location $repoRoot
try {
    $branch = (& cmd /c "git rev-parse --abbrev-ref HEAD 2>nul")
    if ($LASTEXITCODE -ne 0 -or -not $branch) {
        throw "Not a git repository or branch could not be resolved."
    }

    Write-Host "[push-and-launch] Current branch: $branch"
    Write-Host "[push-and-launch] Running git push..."
    & git push @GitPushArgs
    if ($LASTEXITCODE -ne 0) {
        Write-Error "[push-and-launch] git push failed."
        exit $LASTEXITCODE
    }

    $shouldLaunch = $branch -eq "main" -or $branch -like "release/*"
    if (-not $shouldLaunch) {
        Write-Host "[push-and-launch] Launch skipped for branch: $branch"
        exit 0
    }

    if (-not (Test-Path $exePath)) {
        if (-not (Test-Path $buildScript)) {
            throw "Build script not found: $buildScript"
        }

        Write-Host "[push-and-launch] EXE not found. Building local EXE..."
        & powershell -ExecutionPolicy Bypass -File $buildScript -NoLaunch
        if ($LASTEXITCODE -ne 0) {
            throw "Local EXE build failed."
        }
    }

    if (-not (Test-Path $exePath)) {
        throw "EXE was not generated: $exePath"
    }

    Write-Host "[push-and-launch] Launching EXE..."
    Start-Process -FilePath $exePath
    exit 0
}
finally {
    Pop-Location
}
