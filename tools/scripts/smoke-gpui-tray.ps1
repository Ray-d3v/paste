param(
    [string]$ExePath = "artifacts/local-gpui/PasteWinUI.exe",
    [int]$StartupDelayMs = 1500,
    [int]$CommandDelayMs = 1200,
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

public static class PasteTraySmokeProbe
{
    const uint WM_PASTE_GPUI_TEST_TRAY_COMMAND = 0xD058;
    const int TEST_TRAY_COMMAND_SHOW = 0x5201;
    const int TEST_TRAY_COMMAND_EXIT = 0x5203;

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

    [DllImport("user32.dll", SetLastError = true)]
    static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

    [DllImport("user32.dll", SetLastError = true)]
    static extern bool PostMessage(IntPtr hWnd, uint Msg, UIntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    static extern bool IsWindowVisible(IntPtr hWnd);

    public static IntPtr FindOverlayWindow()
    {
        return FindWindow(null, "Paste GPUI Overlay");
    }

    public static bool IsOverlayVisible()
    {
        IntPtr hwnd = FindOverlayWindow();
        return hwnd != IntPtr.Zero && IsWindowVisible(hwnd);
    }

    public static IntPtr FindMessageWindow()
    {
        IntPtr hwnd = FindWindowEx(new IntPtr(-3), IntPtr.Zero, "PasteGPUIMessageWindow", "Paste GPUI Message Window");
        return hwnd != IntPtr.Zero
            ? hwnd
            : FindWindowEx(IntPtr.Zero, IntPtr.Zero, "PasteGPUIMessageWindow", "Paste GPUI Message Window");
    }

    public static bool SendTrayShow(IntPtr hwnd)
    {
        return PostMessage(hwnd, WM_PASTE_GPUI_TEST_TRAY_COMMAND, (UIntPtr)TEST_TRAY_COMMAND_SHOW, IntPtr.Zero);
    }

    public static bool SendTrayExit(IntPtr hwnd)
    {
        return PostMessage(hwnd, WM_PASTE_GPUI_TEST_TRAY_COMMAND, (UIntPtr)TEST_TRAY_COMMAND_EXIT, IntPtr.Zero);
    }
}
'@

if (-not ("PasteTraySmokeProbe" -as [type])) {
    Add-Type -TypeDefinition $probeSource
}

Get-Process -Name "PasteWinUI" -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Milliseconds 250

$process = $null
$status = $null

try {
    $process = Start-Process -FilePath $ExePath -PassThru
    Start-Sleep -Milliseconds $StartupDelayMs
    $process.Refresh()

    if ($process.HasExited) {
        throw "PasteWinUI exited during startup. ExitCode=$($process.ExitCode)"
    }

    $messageWindow = [PasteTraySmokeProbe]::FindMessageWindow()
    if ($messageWindow -eq [IntPtr]::Zero) {
        throw "Paste GPUI message window was not found."
    }

    $overlayHandleBefore = [PasteTraySmokeProbe]::FindOverlayWindow()
    $overlayVisibleBefore = [PasteTraySmokeProbe]::IsOverlayVisible()
    $showPosted = [PasteTraySmokeProbe]::SendTrayShow($messageWindow)
    Start-Sleep -Milliseconds $CommandDelayMs
    $process.Refresh()

    $overlayHandleAfterShow = [PasteTraySmokeProbe]::FindOverlayWindow()
    $overlayVisibleAfterShow = [PasteTraySmokeProbe]::IsOverlayVisible()
    $exitPosted = [PasteTraySmokeProbe]::SendTrayExit($messageWindow)
    $exitWaited = $process.WaitForExit($CommandDelayMs)
    $process.Refresh()

    $status = [ordered]@{
        exe = $ExePath
        process_id = $process.Id
        message_window_handle = $messageWindow.ToInt64()
        overlay_window_handle_before_show = $overlayHandleBefore.ToInt64()
        overlay_visible_before_show = $overlayVisibleBefore
        tray_show_posted = $showPosted
        overlay_window_handle_after_show = $overlayHandleAfterShow.ToInt64()
        overlay_visible_after_show = $overlayVisibleAfterShow
        tray_exit_posted = $exitPosted
        process_exit_waited = $exitWaited
        process_exited_after_exit = $process.HasExited
        exit_code = if ($process.HasExited) { $process.ExitCode } else { $null }
    }

    $status | ConvertTo-Json -Depth 3

    if (-not $showPosted) {
        throw "Failed to post tray Show command."
    }

    if (-not $overlayVisibleAfterShow) {
        throw "Overlay was not visible after tray Show command."
    }

    if (-not $exitPosted) {
        throw "Failed to post tray Exit command."
    }

    if (-not $process.HasExited) {
        throw "PasteWinUI did not exit after tray Exit command."
    }
}
finally {
    if (-not $KeepRunning -and $process -ne $null -and -not $process.HasExited) {
        Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
    }
}
