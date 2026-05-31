param(
    [switch]$StartIfMissing,
    [int]$WaitSeconds = 30,
    [string]$OutputPath = "artifacts/manual-acceptance/gpui-restart-pid-check.json",
    [string]$ManualAcceptancePath = "artifacts/manual-acceptance/gpui-manual-acceptance.json",
    [switch]$UpdateManualAcceptance
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$buildScript = Join-Path $repoRoot "tools\scripts\build-local-exe.ps1"

if (-not [System.IO.Path]::IsPathRooted($OutputPath)) {
    $OutputPath = Join-Path $repoRoot $OutputPath
}

if (-not [System.IO.Path]::IsPathRooted($ManualAcceptancePath)) {
    $ManualAcceptancePath = Join-Path $repoRoot $ManualAcceptancePath
}

function Get-PasteProcessIds {
    @(Get-Process -Name "PasteWinUI" -ErrorAction SilentlyContinue | Sort-Object Id | ForEach-Object { $_.Id })
}

if ($StartIfMissing -and @(Get-PasteProcessIds).Count -eq 0) {
    & powershell -ExecutionPolicy Bypass -File $buildScript
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to start PasteWinUI. ExitCode=$LASTEXITCODE"
    }
    Start-Sleep -Milliseconds 750
}

$beforeIds = @(Get-PasteProcessIds)
if ($beforeIds.Count -eq 0) {
    throw "PasteWinUI.exe is not running. Start it first or pass -StartIfMissing."
}

Write-Host "Current PasteWinUI process id(s): $($beforeIds -join ', ')"
Write-Host "Now use the physical tray menu and click Restart."
Write-Host "Waiting up to $WaitSeconds second(s) for the process id to change..."

$deadline = (Get-Date).AddSeconds($WaitSeconds)
$afterIds = $beforeIds
do {
    Start-Sleep -Milliseconds 500
    $afterIds = @(Get-PasteProcessIds)
    $removed = @($beforeIds | Where-Object { $_ -notin $afterIds })
    $added = @($afterIds | Where-Object { $_ -notin $beforeIds })
    if ($removed.Count -gt 0 -and $added.Count -gt 0) {
        break
    }
} while ((Get-Date) -lt $deadline)

$status = [ordered]@{
    checked_at = (Get-Date).ToString("o")
    before_process_ids = $beforeIds
    after_process_ids = $afterIds
    removed_process_ids = @($beforeIds | Where-Object { $_ -notin $afterIds })
    added_process_ids = @($afterIds | Where-Object { $_ -notin $beforeIds })
}

$status.restarted = $status.removed_process_ids.Count -gt 0 -and $status.added_process_ids.Count -gt 0
$outputDir = Split-Path -Parent $OutputPath
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
$status | ConvertTo-Json -Depth 4 | Set-Content -Path $OutputPath -Encoding UTF8
$status | ConvertTo-Json -Depth 4

if (-not $status.restarted) {
    throw "PasteWinUI restart was not observed. Before=$($beforeIds -join ',') After=$($afterIds -join ','). See $OutputPath"
}

if ($UpdateManualAcceptance) {
    if (-not (Test-Path $ManualAcceptancePath)) {
        throw "Restart was observed, but manual acceptance record was not found: $ManualAcceptancePath"
    }

    $record = Get-Content -Path $ManualAcceptancePath -Raw | ConvertFrom-Json
    $restartCheck = @($record.checks | Where-Object { $_.id -eq "tray_menu_restart" }) | Select-Object -First 1
    if ($null -eq $restartCheck) {
        throw "Restart was observed, but tray_menu_restart was not found in $ManualAcceptancePath"
    }

    $restartCheck.passed = $true
    $restartCheck.note = "Verified by process ID change. Evidence: $OutputPath"
    $record.all_passed = @($record.checks | Where-Object { $_.passed -ne $true }).Count -eq 0
    $record.notes = (($record.notes, "tray_menu_restart verified by check-gpui-restart-pid.ps1") | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join " | "
    $record | ConvertTo-Json -Depth 5 | Set-Content -Path $ManualAcceptancePath -Encoding UTF8
    Write-Host "Updated manual acceptance record: $ManualAcceptancePath"
}

Write-Host "Restart observed. Evidence written to: $OutputPath"
