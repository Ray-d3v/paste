param(
    [string]$InstallerPath = "deploy/installer/output/PasteWinUI-Setup-0.1.0-offline.exe",
    [int]$InstallTimeoutSeconds = 60,
    [int]$CleanupDelayMs = 2000
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

if (-not [System.IO.Path]::IsPathRooted($InstallerPath)) {
    $InstallerPath = Join-Path $repoRoot $InstallerPath
}

if (-not (Test-Path $InstallerPath)) {
    throw "Installer not found: $InstallerPath"
}

$smokeRoot = Join-Path $repoRoot "artifacts\installer-smoke"
$installDir = Join-Path $smokeRoot "app"
$startMenuDir = Join-Path $smokeRoot "start-menu"
$desktopShortcutPath = Join-Path $smokeRoot "desktop\PasteWinUI.lnk"
$uninstallKeyName = "PasteWinUI-Smoke"
$uninstallKeyPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\$uninstallKeyName"
$runValueName = "PasteWinUI-Smoke"
$runKeyPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
$exePath = Join-Path $installDir "PasteWinUI.exe"
$startMenuShortcut = Join-Path $startMenuDir "PasteWinUI.lnk"
$uninstallShortcut = Join-Path $startMenuDir "Uninstall PasteWinUI.lnk"
$uninstallScript = Join-Path $installDir "uninstall.ps1"

Get-Process -Name "PasteWinUI" -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Milliseconds 250

Remove-Item -Path $smokeRoot -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -Path $uninstallKeyPath -Recurse -Force -ErrorAction SilentlyContinue
Remove-ItemProperty -Path $runKeyPath -Name $runValueName -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Path (Split-Path -Parent $desktopShortcutPath) -Force | Out-Null

$previousEnv = @{
    PASTEWINUI_INSTALL_DIR = $env:PASTEWINUI_INSTALL_DIR
    PASTEWINUI_START_MENU_DIR = $env:PASTEWINUI_START_MENU_DIR
    PASTEWINUI_DESKTOP_SHORTCUT_PATH = $env:PASTEWINUI_DESKTOP_SHORTCUT_PATH
    PASTEWINUI_RUN_VALUE_NAME = $env:PASTEWINUI_RUN_VALUE_NAME
    PASTEWINUI_UNINSTALL_KEY_NAME = $env:PASTEWINUI_UNINSTALL_KEY_NAME
    PASTEWINUI_CREATE_DESKTOP_SHORTCUT = $env:PASTEWINUI_CREATE_DESKTOP_SHORTCUT
    PASTEWINUI_ENABLE_AUTOSTART = $env:PASTEWINUI_ENABLE_AUTOSTART
    PASTEWINUI_START_AFTER_INSTALL = $env:PASTEWINUI_START_AFTER_INSTALL
}

function Restore-InstallerSmokeEnvironment {
    foreach ($entry in $previousEnv.GetEnumerator()) {
        Set-Item -Path "Env:\$($entry.Key)" -Value $entry.Value -ErrorAction SilentlyContinue
        if ($null -eq $entry.Value) {
            Remove-Item -Path "Env:\$($entry.Key)" -ErrorAction SilentlyContinue
        }
    }
}

try {
    $env:PASTEWINUI_INSTALL_DIR = $installDir
    $env:PASTEWINUI_START_MENU_DIR = $startMenuDir
    $env:PASTEWINUI_DESKTOP_SHORTCUT_PATH = $desktopShortcutPath
    $env:PASTEWINUI_RUN_VALUE_NAME = $runValueName
    $env:PASTEWINUI_UNINSTALL_KEY_NAME = $uninstallKeyName
    $env:PASTEWINUI_CREATE_DESKTOP_SHORTCUT = "1"
    $env:PASTEWINUI_ENABLE_AUTOSTART = "1"
    $env:PASTEWINUI_START_AFTER_INSTALL = "0"

    $installProcess = Start-Process -FilePath $InstallerPath -Wait -PassThru
    if ($installProcess.ExitCode -ne 0) {
        throw "Offline installer failed. ExitCode=$($installProcess.ExitCode)"
    }

    $startedProcess = Start-Process -FilePath $exePath -PassThru
    Start-Sleep -Milliseconds 1200
    $startedProcess.Refresh()
    $appStartedBySmoke = -not $startedProcess.HasExited

    $uninstallItem = Get-ItemProperty -Path $uninstallKeyPath -ErrorAction SilentlyContinue
    $runValueItem = Get-ItemProperty -Path $runKeyPath -Name $runValueName -ErrorAction SilentlyContinue
    $runValue = if ($null -ne $runValueItem -and $null -ne $runValueItem.PSObject.Properties[$runValueName]) {
        $runValueItem.PSObject.Properties[$runValueName].Value
    }
    else {
        $null
    }

    $status = [ordered]@{
        installer = $InstallerPath
        install_exit_code = $installProcess.ExitCode
        install_dir = $installDir
        exe_exists = Test-Path $exePath
        start_menu_shortcut_exists = Test-Path $startMenuShortcut
        uninstall_shortcut_exists = Test-Path $uninstallShortcut
        desktop_shortcut_exists = Test-Path $desktopShortcutPath
        uninstall_key_exists = $null -ne $uninstallItem
        uninstall_display_name = if ($null -ne $uninstallItem) { $uninstallItem.DisplayName } else { $null }
        uninstall_display_version = if ($null -ne $uninstallItem) { $uninstallItem.DisplayVersion } else { $null }
        uninstall_script_exists = Test-Path $uninstallScript
        run_value = $runValue
        app_start_verified_by_smoke = $appStartedBySmoke
        app_process_id = if ($appStartedBySmoke) { $startedProcess.Id } else { $null }
    }

    $status | ConvertTo-Json -Depth 4

    if (-not $status.exe_exists) { throw "Installed EXE was not found: $exePath" }
    if (-not $status.start_menu_shortcut_exists) { throw "Start Menu shortcut was not created." }
    if (-not $status.uninstall_shortcut_exists) { throw "Uninstall shortcut was not created." }
    if (-not $status.desktop_shortcut_exists) { throw "Desktop shortcut override path was not created." }
    if (-not $status.uninstall_key_exists) { throw "Uninstall registry key was not created." }
    if ($status.uninstall_display_name -ne "PasteWinUI") { throw "Unexpected uninstall DisplayName: $($status.uninstall_display_name)" }
    if (-not $status.uninstall_script_exists) { throw "Uninstall script was not created." }
    if ([string]::IsNullOrWhiteSpace($status.run_value) -or $status.run_value -notmatch [regex]::Escape($exePath)) {
        throw "Auto-start Run value was not created for smoke install."
    }
    if (-not $status.app_start_verified_by_smoke) { throw "Installed app did not start from the installed path." }

    if ($appStartedBySmoke) {
        Stop-Process -Id $startedProcess.Id -Force -ErrorAction SilentlyContinue
    }

    powershell -NoProfile -ExecutionPolicy Bypass -File $uninstallScript
    Start-Sleep -Milliseconds $CleanupDelayMs

    $cleanupStatus = [ordered]@{
        install_dir_exists_after_uninstall = Test-Path $installDir
        start_menu_dir_exists_after_uninstall = Test-Path $startMenuDir
        desktop_shortcut_exists_after_uninstall = Test-Path $desktopShortcutPath
        uninstall_key_exists_after_uninstall = Test-Path $uninstallKeyPath
        run_value_exists_after_uninstall = $null -ne (Get-ItemProperty -Path $runKeyPath -Name $runValueName -ErrorAction SilentlyContinue)
    }

    $cleanupStatus | ConvertTo-Json -Depth 3

    if ($cleanupStatus.uninstall_key_exists_after_uninstall) {
        throw "Uninstall registry key still exists after uninstall."
    }
    if ($cleanupStatus.run_value_exists_after_uninstall) {
        throw "Auto-start Run value still exists after uninstall."
    }
}
finally {
    Restore-InstallerSmokeEnvironment
    Get-Process -Name "PasteWinUI" -ErrorAction SilentlyContinue | Where-Object {
        try { $_.MainModule.FileName -eq $exePath } catch { $false }
    } | Stop-Process -Force -ErrorAction SilentlyContinue
    Remove-ItemProperty -Path $runKeyPath -Name $runValueName -ErrorAction SilentlyContinue
    Remove-Item -Path $uninstallKeyPath -Recurse -Force -ErrorAction SilentlyContinue
}
