Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
Push-Location $repoRoot
try {
    $gitDir = (& cmd /c "git rev-parse --git-dir 2>nul")
    if ($LASTEXITCODE -ne 0 -or -not $gitDir) {
        throw "Not a git repository. Run this script from a cloned git working tree."
    }

    $gitDirPath = if ([System.IO.Path]::IsPathRooted($gitDir)) {
        $gitDir
    }
    else {
        Join-Path $repoRoot $gitDir
    }

    $hooksDir = Join-Path $gitDirPath "hooks"
    $prePushHookPath = Join-Path $hooksDir "pre-push"
    $hookScriptPath = Join-Path $repoRoot "scripts/hooks/pre-push.ps1"

    if (-not (Test-Path $hookScriptPath)) {
        throw "Hook script not found: $hookScriptPath"
    }

    New-Item -ItemType Directory -Path $hooksDir -Force | Out-Null

    $wrapper = @(
        "#!/usr/bin/env bash"
        "set -euo pipefail"
        "powershell -ExecutionPolicy Bypass -File `"$hookScriptPath`""
    ) -join "`n"

    Set-Content -Path $prePushHookPath -Value $wrapper -Encoding Ascii
    Write-Host "Installed pre-push hook: $prePushHookPath"
    Write-Host "Hook target: $hookScriptPath"
}
finally {
    Pop-Location
}
