Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$hooksDir = Split-Path -Parent $PSCommandPath
$scriptsDir = Split-Path -Parent $hooksDir
$repoRoot = Split-Path -Parent $scriptsDir
$buildScript = Join-Path $scriptsDir "build-local-exe.ps1"

if (-not (Test-Path $buildScript)) {
    Write-Error "Required script not found: $buildScript"
    exit 1
}

Push-Location $repoRoot
try {
    Write-Host "[pre-push] Building local EXE for verification..."
    & powershell -ExecutionPolicy Bypass -File $buildScript -NoLaunch
    if ($LASTEXITCODE -ne 0) {
        Write-Error "[pre-push] EXE build failed."
        exit 1
    }

    Write-Host "[pre-push] EXE verification succeeded."
    exit 0
}
catch {
    Write-Error "[pre-push] EXE verification failed: $($_.Exception.Message)"
    exit 1
}
finally {
    Pop-Location
}
