param(
    [switch]$SkipReleaseVerifier,
    [switch]$SkipLaunch,
    [switch]$KeepExistingProcess
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$releaseVerifier = Join-Path $repoRoot "tools\scripts\verify-gpui-release.ps1"
$buildScript = Join-Path $repoRoot "tools\scripts\build-local-exe.ps1"
$trayRectProbe = Join-Path $repoRoot "tools\scripts\probe-gpui-tray-rect.ps1"
$trayUiProbe = Join-Path $repoRoot "tools\scripts\probe-gpui-tray-ui.ps1"
$recordScript = Join-Path $repoRoot "tools\scripts\record-gpui-manual-acceptance.ps1"
$acceptanceVerifier = Join-Path $repoRoot "tools\scripts\verify-gpui-acceptance.ps1"
$manualDoc = Join-Path $repoRoot "docs\gpui-manual-acceptance.md"

if (-not $SkipReleaseVerifier) {
    & powershell -ExecutionPolicy Bypass -File $releaseVerifier
    if ($LASTEXITCODE -ne 0) {
        throw "Release verifier failed. ExitCode=$LASTEXITCODE"
    }
}

if (-not $SkipLaunch) {
    if (-not $KeepExistingProcess) {
        Get-Process -Name "PasteWinUI" -ErrorAction SilentlyContinue | Stop-Process -Force
        Start-Sleep -Milliseconds 250
    }

    & powershell -ExecutionPolicy Bypass -File $buildScript
    if ($LASTEXITCODE -ne 0) {
        throw "Build and launch failed. ExitCode=$LASTEXITCODE"
    }
}

$running = @(Get-Process -Name "PasteWinUI" -ErrorAction SilentlyContinue)
if ($running.Count -eq 0) {
    throw "PasteWinUI.exe is not running. Run tools/scripts/build-local-exe.ps1 before manual acceptance."
}

$trayRect = $null
try {
    $trayRectRaw = & powershell -ExecutionPolicy Bypass -File $trayRectProbe -UseExistingProcess -KeepRunning
    $trayRect = $trayRectRaw | ConvertFrom-Json
}
catch {
    Write-Warning "Tray rectangle probe failed: $($_.Exception.Message)"
}

$summary = [ordered]@{
    prepared_at = (Get-Date).ToString("o")
    release_verifier = if ($SkipReleaseVerifier) { "skipped" } else { "passed" }
    launched_by_script = -not $SkipLaunch
    paste_process_ids = @($running | ForEach-Object { $_.Id })
    tray_rect_available = $null -ne $trayRect -and $trayRect.rect_available -eq $true
    tray_rect = $trayRect
    manual_checklist = $manualDoc
    optional_tray_ui_probe_command = "powershell -ExecutionPolicy Bypass -File `"$trayUiProbe`""
    record_command = "powershell -ExecutionPolicy Bypass -File `"$recordScript`""
    final_acceptance_command = "powershell -ExecutionPolicy Bypass -File `"$acceptanceVerifier`""
}

$summary | ConvertTo-Json -Depth 5

Write-Host ""
Write-Host "Manual acceptance next steps:"
Write-Host "1. Locate the PasteWinUI notification icon. If it is hidden, open the Windows tray overflow."
Write-Host "2. Left-click the physical tray icon and confirm the overlay appears."
Write-Host "3. Right-click the physical tray icon and confirm Show, Restart, Exit."
Write-Host "4. Complete the overlay visual and default interactive installer checks in docs/gpui-manual-acceptance.md."
Write-Host "5. Record the result with:"
Write-Host "   powershell -ExecutionPolicy Bypass -File `"$recordScript`""
Write-Host "6. Run the final gate with:"
Write-Host "   powershell -ExecutionPolicy Bypass -File `"$acceptanceVerifier`""
