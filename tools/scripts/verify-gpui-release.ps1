param(
    [switch]$SkipBuild,
    [switch]$SkipCargoTests,
    [switch]$SkipSmokes,
    [switch]$SkipOfflineInstaller,
    [string]$Version = "0.1.0",
    [int]$SmokeDelayMs = 1800
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$results = @()

function Invoke-Step {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,
        [Parameter(Mandatory = $true)]
        [scriptblock]$Script,
        [int]$Retries = 0
    )

    $started = Get-Date
    Write-Host "==> $Name"
    $attempt = 0
    while ($true) {
        try {
            $attempt++
            $global:LASTEXITCODE = $null
            $previousErrorActionPreference = $ErrorActionPreference
            $ErrorActionPreference = "Continue"
            try {
                $stepOutput = @(& $Script 2>&1)
                $stepExitCode = $LASTEXITCODE
            }
            finally {
                $ErrorActionPreference = $previousErrorActionPreference
            }

            if ($stepExitCode -ne $null -and $stepExitCode -ne 0) {
                $tail = ($stepOutput | Select-Object -Last 20) -join [Environment]::NewLine
                throw "$Name failed with exit code $stepExitCode$([Environment]::NewLine)$tail"
            }

            if ($stepOutput.Count -gt 0) {
                $stepOutput | ForEach-Object { Write-Host $_ }
            }

            $script:results += [pscustomobject]@{
                name = $Name
                status = "passed"
                attempts = $attempt
                duration_ms = [int]((Get-Date) - $started).TotalMilliseconds
            }
            return
        }
        catch {
            if ($attempt -le $Retries) {
                Write-Warning "$Name attempt $attempt failed: $($_.Exception.Message)"
                Get-Process -Name "PasteWinUI" -ErrorAction SilentlyContinue | Stop-Process -Force
                Start-Sleep -Milliseconds 750
                continue
            }

            $script:results += [pscustomobject]@{
                name = $Name
                status = "failed"
                attempts = $attempt
                duration_ms = [int]((Get-Date) - $started).TotalMilliseconds
                error = $_.Exception.Message
            }
            throw
        }
    }
}

try {
    if (-not $SkipBuild) {
        Invoke-Step "build-local-exe" {
            powershell -ExecutionPolicy Bypass -File (Join-Path $repoRoot "tools/scripts/build-local-exe.ps1") -NoLaunch -Profile release
        }
    }

    if (-not $SkipCargoTests) {
        Invoke-Step "cargo-test-workspace" {
            cargo test --manifest-path (Join-Path $repoRoot "src-paste-gpui/Cargo.toml") --workspace
        }
    }

    if (-not $SkipSmokes) {
        Invoke-Step "smoke-hotkey" {
            powershell -ExecutionPolicy Bypass -File (Join-Path $repoRoot "tools/scripts/smoke-gpui-hotkey.ps1")
        } -Retries 1

        Invoke-Step "smoke-slide-up" {
            powershell -ExecutionPolicy Bypass -File (Join-Path $repoRoot "tools/scripts/smoke-gpui-slide.ps1")
        } -Retries 1

        Invoke-Step "smoke-command-palette" {
            powershell -ExecutionPolicy Bypass -File (Join-Path $repoRoot "tools/scripts/smoke-gpui-command-palette.ps1") -AllowSyntheticInputUnavailable
        } -Retries 1

        Invoke-Step "smoke-text-paste" {
            powershell -ExecutionPolicy Bypass -File (Join-Path $repoRoot "tools/scripts/smoke-gpui-paste-text.ps1") -OverlayDelayMs $SmokeDelayMs -PasteDelayMs $SmokeDelayMs -AllowForegroundUnavailable
        } -Retries 1

        Invoke-Step "smoke-link-paste" {
            powershell -ExecutionPolicy Bypass -File (Join-Path $repoRoot "tools/scripts/smoke-gpui-paste-link.ps1") -OverlayDelayMs $SmokeDelayMs -PasteDelayMs $SmokeDelayMs -AllowForegroundUnavailable
        } -Retries 1

        Invoke-Step "smoke-image-paste" {
            powershell -ExecutionPolicy Bypass -File (Join-Path $repoRoot "tools/scripts/smoke-gpui-paste-image.ps1") -OverlayDelayMs $SmokeDelayMs -PasteDelayMs $SmokeDelayMs -AllowForegroundUnavailable
        } -Retries 1

        Invoke-Step "smoke-tray-show-exit" {
            powershell -ExecutionPolicy Bypass -File (Join-Path $repoRoot "tools/scripts/smoke-gpui-tray.ps1")
        } -Retries 1

        Invoke-Step "smoke-tray-rect" {
            powershell -ExecutionPolicy Bypass -File (Join-Path $repoRoot "tools/scripts/probe-gpui-tray-rect.ps1") -AllowUnavailable
        } -Retries 1

        Invoke-Step "smoke-tray-notification" {
            powershell -ExecutionPolicy Bypass -File (Join-Path $repoRoot "tools/scripts/smoke-gpui-tray-notification.ps1") -AllowPopupUnavailable
        } -Retries 1

        Invoke-Step "smoke-tray-restart" {
            powershell -ExecutionPolicy Bypass -File (Join-Path $repoRoot "tools/scripts/smoke-gpui-tray-restart.ps1")
        } -Retries 1
    }

    if (-not $SkipOfflineInstaller) {
        Invoke-Step "build-offline-installer" {
            powershell -ExecutionPolicy Bypass -File (Join-Path $repoRoot "tools/scripts/build-offline-installer.ps1") -Version $Version -SkipBuild
        }

        Invoke-Step "smoke-offline-installer" {
            powershell -ExecutionPolicy Bypass -File (Join-Path $repoRoot "tools/scripts/smoke-gpui-offline-installer.ps1") -InstallerPath (Join-Path $repoRoot "deploy/installer/output/PasteWinUI-Setup-$Version-offline.exe")
        }
    }
}
finally {
    Get-Process -Name "PasteWinUI" -ErrorAction SilentlyContinue | Stop-Process -Force
}

$summary = [pscustomobject]@{
    generated_at = (Get-Date).ToString("o")
    results = @($results)
}

$summary | ConvertTo-Json -Depth 5
