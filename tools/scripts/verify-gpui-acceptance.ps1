param(
    [string]$ManualAcceptancePath = "artifacts/manual-acceptance/gpui-manual-acceptance.json",
    [switch]$SkipReleaseVerifier
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$releaseVerifier = Join-Path $repoRoot "tools\scripts\verify-gpui-release.ps1"

if (-not [System.IO.Path]::IsPathRooted($ManualAcceptancePath)) {
    $ManualAcceptancePath = Join-Path $repoRoot $ManualAcceptancePath
}

if (-not $SkipReleaseVerifier) {
    & powershell -ExecutionPolicy Bypass -File $releaseVerifier
    if ($LASTEXITCODE -ne 0) {
        throw "Release verifier failed. ExitCode=$LASTEXITCODE"
    }
}

if (-not (Test-Path $ManualAcceptancePath)) {
    throw "Manual acceptance record not found: $ManualAcceptancePath"
}

$record = Get-Content -Path $ManualAcceptancePath -Raw | ConvertFrom-Json
$checks = @($record.checks)
if ($checks.Count -eq 0) {
    throw "Manual acceptance record has no checks: $ManualAcceptancePath"
}

$failedChecks = @($checks | Where-Object { $_.passed -ne $true })
if ($failedChecks.Count -gt 0) {
    $failedDetails = @($failedChecks | ForEach-Object {
        $note = if ([string]::IsNullOrWhiteSpace($_.note)) { "<no note>" } else { $_.note }
        "$($_.id) [$($_.category)]: $note"
    })
    $retryHint = ""
    if (@($failedChecks | Where-Object { $_.id -eq "tray_menu_restart" }).Count -gt 0) {
        $retryHint = " To re-check tray restart, run: powershell -ExecutionPolicy Bypass -File tools\scripts\check-gpui-restart-pid.ps1 -StartIfMissing -UpdateManualAcceptance -WaitSeconds 90, then click Restart from the physical tray menu while it waits."
    }
    throw "Manual acceptance has failed checks in ${ManualAcceptancePath}: $($failedDetails -join '; ').$retryHint"
}

if ($record.all_passed -ne $true) {
    throw "Manual acceptance record has all_passed=false but no failed check entries: $ManualAcceptancePath"
}

$requiredCheckIds = @(
    "tray_icon_visible",
    "tray_left_click_show",
    "tray_right_click_menu",
    "tray_menu_show",
    "tray_menu_restart",
    "tray_menu_exit",
    "overlay_bottom_slide",
    "overlay_escape_hide",
    "installer_default_installs",
    "installer_default_shortcuts_uninstall",
    "installer_default_launches_app"
)

$actualCheckIds = @($checks | ForEach-Object { $_.id })
$missingCheckIds = @($requiredCheckIds | Where-Object { $_ -notin $actualCheckIds })
if ($missingCheckIds.Count -gt 0) {
    $actual = if ($actualCheckIds.Count -gt 0) { $actualCheckIds -join ', ' } else { '<none>' }
    throw "Manual acceptance record is missing checks: $($missingCheckIds -join ', '). Actual checks: $actual"
}

$summary = [ordered]@{
    accepted = $true
    checked_at = (Get-Date).ToString("o")
    manual_acceptance_path = $ManualAcceptancePath
    manual_recorded_at = $record.recorded_at
    operator = $record.operator
    check_count = $checks.Count
    release_verifier = if ($SkipReleaseVerifier) { "skipped" } else { "passed" }
}

$summary | ConvertTo-Json -Depth 4
