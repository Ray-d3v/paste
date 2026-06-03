param(
    [string]$OutputPath = "artifacts/manual-acceptance/gpui-manual-acceptance.json",
    [switch]$AssumePassed,
    [switch]$FailOnFailedChecks,
    [switch]$RetryFailedOnly,
    [string]$Operator = $env:USERNAME,
    [string]$Notes = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

if (-not [System.IO.Path]::IsPathRooted($OutputPath)) {
    $OutputPath = Join-Path $repoRoot $OutputPath
}

$existingChecksById = @{}
if ($RetryFailedOnly) {
    if (-not (Test-Path $OutputPath)) {
        throw "RetryFailedOnly was specified, but the acceptance record does not exist: $OutputPath"
    }

    $existingRecord = Get-Content -Path $OutputPath -Raw | ConvertFrom-Json
    foreach ($check in @($existingRecord.checks)) {
        $existingChecksById[$check.id] = $check
    }
}

function Read-ChecklistResult {
    param(
        [Parameter(Mandatory = $true)][string]$Id,
        [Parameter(Mandatory = $true)][string]$Prompt,
        [string]$Category = "manual"
    )

    if ($RetryFailedOnly -and $existingChecksById.ContainsKey($Id)) {
        $existing = $existingChecksById[$Id]
        if ($existing.passed -eq $true) {
            return [pscustomobject]@{
                id = $Id
                category = $Category
                prompt = $Prompt
                passed = $true
                note = $existing.note
            }
        }
    }

    if ($AssumePassed) {
        return [pscustomobject]@{
            id = $Id
            category = $Category
            prompt = $Prompt
            passed = $true
            note = ""
        }
    }

    while ($true) {
        $answer = Read-Host "$Prompt [y/n]"
        switch ($answer.Trim().ToLowerInvariant()) {
            "y" {
                return [pscustomobject]@{
                    id = $Id
                    category = $Category
                    prompt = $Prompt
                    passed = $true
                    note = ""
                }
            }
            "yes" {
                return [pscustomobject]@{
                    id = $Id
                    category = $Category
                    prompt = $Prompt
                    passed = $true
                    note = ""
                }
            }
            "n" {
                $note = Read-Host "Failure note for $Id"
                return [pscustomobject]@{
                    id = $Id
                    category = $Category
                    prompt = $Prompt
                    passed = $false
                    note = $note
                }
            }
            "no" {
                $note = Read-Host "Failure note for $Id"
                return [pscustomobject]@{
                    id = $Id
                    category = $Category
                    prompt = $Prompt
                    passed = $false
                    note = $note
                }
            }
            default {
                Write-Host "Please answer y or n."
            }
        }
    }
}

$checks = @(
    Read-ChecklistResult -Id "tray_icon_visible" -Category "physical_shell" -Prompt "PasteWinUI tray icon was located in the notification area or overflow"
    Read-ChecklistResult -Id "tray_left_click_show" -Category "physical_shell" -Prompt "Left-clicking the physical tray icon showed the overlay"
    Read-ChecklistResult -Id "tray_right_click_menu" -Category "physical_shell" -Prompt "Right-clicking the physical tray icon showed Show/Restart/Exit"
    Read-ChecklistResult -Id "tray_menu_show" -Category "physical_shell" -Prompt "Physical tray menu Show displayed or preserved the overlay"
    Read-ChecklistResult -Id "tray_menu_restart" -Category "physical_shell" -Prompt "Physical tray menu Restart changed the PasteWinUI process ID; no visible UI change is required"
    Read-ChecklistResult -Id "tray_menu_exit" -Category "physical_shell" -Prompt "Physical tray menu Exit stopped PasteWinUI"
    Read-ChecklistResult -Id "overlay_bottom_slide" -Category "visual_manual" -Prompt "Ctrl+Alt+V showed a bottom overlay with slide-up motion on the target display"
    Read-ChecklistResult -Id "overlay_card_types" -Category "visual_manual" -Prompt "Overlay showed horizontal cards and visually distinct Text/Link/Image/File/Code/Favorite states"
    Read-ChecklistResult -Id "overlay_preview_panels" -Category "visual_manual" -Prompt "Overlay preview, command palette, context menu, delete confirmation, and settings panel were reachable"
    Read-ChecklistResult -Id "overlay_escape_hide" -Category "visual_manual" -Prompt "Esc hid the overlay immediately on the target display"
    Read-ChecklistResult -Id "installer_default_installs" -Category "physical_installer" -Prompt "Default offline installer installed to LocalAppData Programs when run interactively"
    Read-ChecklistResult -Id "installer_default_shortcuts_uninstall" -Category "physical_installer" -Prompt "Default installer created Start Menu shortcuts and uninstall entry"
    Read-ChecklistResult -Id "installer_default_launches_app" -Category "physical_installer" -Prompt "Default installed app launched successfully"
)

$allPassed = @($checks | Where-Object { -not $_.passed }).Count -eq 0
$result = [ordered]@{
    recorded_at = (Get-Date).ToString("o")
    operator = $Operator
    all_passed = $allPassed
    notes = $Notes
    retry_failed_only = [bool]$RetryFailedOnly
    checks = $checks
}

$outputDir = Split-Path -Parent $OutputPath
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
$result | ConvertTo-Json -Depth 5 | Set-Content -Path $OutputPath -Encoding UTF8

$result | ConvertTo-Json -Depth 5

if (-not $allPassed) {
    $failedChecks = @($checks | Where-Object { -not $_.passed })
    Write-Warning "Manual acceptance was recorded with failed checks: $($failedChecks.id -join ', '). See $OutputPath"
    if (@($failedChecks | Where-Object { $_.id -eq "tray_menu_restart" }).Count -gt 0) {
        Write-Warning "To re-check tray restart with PID evidence, run: powershell -ExecutionPolicy Bypass -File tools\scripts\check-gpui-restart-pid.ps1 -StartIfMissing -UpdateManualAcceptance -WaitSeconds 90, then click Restart from the physical tray menu while it waits."
    }
    if ($FailOnFailedChecks) {
        throw "Manual acceptance did not pass all checks. See $OutputPath"
    }
}
