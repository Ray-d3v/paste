param(
    [string]$ExePath = "artifacts/local-gpui/PasteWinUI.exe",
    [int]$StartupDelayMs = 1500,
    [int]$RestartDelayMs = 2500,
    [switch]$KeepRunning
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

if (-not [System.IO.Path]::IsPathRooted($ExePath)) {
    $ExePath = Join-Path $repoRoot $ExePath
}

if (-not (Test-Path $ExePath)) {
    throw "EXE not found: $ExePath"
}

$probeSource = @'
using System;
using System.Runtime.InteropServices;

public static class PasteTrayRestartSmokeProbe
{
    const uint WM_PASTE_GPUI_TEST_TRAY_COMMAND = 0xD058;
    const int TEST_TRAY_COMMAND_RESTART = 0x5202;

    [DllImport("user32.dll", SetLastError = true)]
    static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

    [DllImport("user32.dll", SetLastError = true)]
    static extern bool PostMessage(IntPtr hWnd, uint Msg, UIntPtr wParam, IntPtr lParam);

    public static IntPtr FindMessageWindow()
    {
        IntPtr hwnd = FindWindowEx(new IntPtr(-3), IntPtr.Zero, "PasteGPUIMessageWindow", "Paste GPUI Message Window");
        return hwnd != IntPtr.Zero
            ? hwnd
            : FindWindowEx(IntPtr.Zero, IntPtr.Zero, "PasteGPUIMessageWindow", "Paste GPUI Message Window");
    }

    public static bool SendTrayRestart(IntPtr hwnd)
    {
        return PostMessage(hwnd, WM_PASTE_GPUI_TEST_TRAY_COMMAND, (UIntPtr)TEST_TRAY_COMMAND_RESTART, IntPtr.Zero);
    }
}
'@

if (-not ("PasteTrayRestartSmokeProbe" -as [type])) {
    Add-Type -TypeDefinition $probeSource
}

function Get-RunningPasteProcesses {
    Get-Process -Name "PasteWinUI" -ErrorAction SilentlyContinue |
        Where-Object { -not $_.HasExited } |
        Sort-Object StartTime
}

Get-Process -Name "PasteWinUI" -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Milliseconds 250

$originalProcess = $null
$newProcess = $null

try {
    $originalProcess = Start-Process -FilePath $ExePath -PassThru
    Start-Sleep -Milliseconds $StartupDelayMs
    $originalProcess.Refresh()

    if ($originalProcess.HasExited) {
        throw "PasteWinUI exited during startup. ExitCode=$($originalProcess.ExitCode)"
    }

    $messageWindow = [PasteTrayRestartSmokeProbe]::FindMessageWindow()
    if ($messageWindow -eq [IntPtr]::Zero) {
        throw "Paste GPUI message window was not found."
    }

    $restartPosted = [PasteTrayRestartSmokeProbe]::SendTrayRestart($messageWindow)
    Start-Sleep -Milliseconds $RestartDelayMs
    $originalProcess.Refresh()

    $running = @(Get-RunningPasteProcesses)
    $newProcess = $running | Where-Object { $_.Id -ne $originalProcess.Id } | Select-Object -First 1

    $status = [ordered]@{
        exe = $ExePath
        original_process_id = $originalProcess.Id
        message_window_handle = $messageWindow.ToInt64()
        restart_posted = $restartPosted
        original_exited_after_restart = $originalProcess.HasExited
        original_exit_code = if ($originalProcess.HasExited) { $originalProcess.ExitCode } else { $null }
        new_process_id = if ($newProcess -ne $null) { $newProcess.Id } else { $null }
        running_paste_process_ids = @($running | ForEach-Object { $_.Id })
    }

    $status | ConvertTo-Json -Depth 3

    if (-not $restartPosted) {
        throw "Failed to post tray Restart command."
    }

    if (-not $originalProcess.HasExited) {
        throw "Original PasteWinUI process did not exit after tray Restart command."
    }

    if ($newProcess -eq $null) {
        throw "No replacement PasteWinUI process was found after tray Restart command."
    }
}
finally {
    if (-not $KeepRunning) {
        if ($newProcess -ne $null -and -not $newProcess.HasExited) {
            Stop-Process -Id $newProcess.Id -Force -ErrorAction SilentlyContinue
        }

        if ($originalProcess -ne $null -and -not $originalProcess.HasExited) {
            Stop-Process -Id $originalProcess.Id -Force -ErrorAction SilentlyContinue
        }
    }
}
