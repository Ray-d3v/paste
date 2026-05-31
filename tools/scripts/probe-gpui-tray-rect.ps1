param(
    [string]$ExePath = "artifacts/local-gpui/PasteWinUI.exe",
    [int]$StartupDelayMs = 1500,
    [switch]$UseExistingProcess,
    [switch]$KeepRunning,
    [switch]$AllowUnavailable
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
using System.ComponentModel;
using System.Runtime.InteropServices;

public static class PasteTrayRectProbe
{
    const int TRAY_ICON_ID = 0x5057;
    const uint WM_PASTE_GPUI_TEST_PLATFORM_STATUS = 0xD059;
    const uint WM_PASTE_GPUI_TEST_TRAY_ERROR = 0xD05A;
    const int PLATFORM_STATUS_TRAY_ICON = 0x04;

    [StructLayout(LayoutKind.Sequential)]
    struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct NOTIFYICONIDENTIFIER
    {
        public int cbSize;
        public IntPtr hWnd;
        public uint uID;
        public Guid guidItem;
    }

    public class TrayRect
    {
        public IntPtr MessageWindow;
        public int HResult;
        public int X;
        public int Y;
        public int Width;
        public int Height;
    }

    [DllImport("user32.dll", SetLastError = true)]
    static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

    [DllImport("shell32.dll", SetLastError = true)]
    static extern int Shell_NotifyIconGetRect(ref NOTIFYICONIDENTIFIER identifier, out RECT iconLocation);

    [DllImport("user32.dll", SetLastError = true)]
    static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, UIntPtr wParam, IntPtr lParam);

    public static IntPtr FindMessageWindow()
    {
        IntPtr hwnd = FindWindowEx(new IntPtr(-3), IntPtr.Zero, "PasteGPUIMessageWindow", "Paste GPUI Message Window");
        return hwnd != IntPtr.Zero
            ? hwnd
            : FindWindowEx(IntPtr.Zero, IntPtr.Zero, "PasteGPUIMessageWindow", "Paste GPUI Message Window");
    }

    public static TrayRect GetTrayRect()
    {
        IntPtr hwnd = FindMessageWindow();
        if (hwnd == IntPtr.Zero)
        {
            throw new InvalidOperationException("Paste GPUI message window was not found.");
        }

        var identifier = new NOTIFYICONIDENTIFIER
        {
            cbSize = Marshal.SizeOf(typeof(NOTIFYICONIDENTIFIER)),
            hWnd = hwnd,
            uID = TRAY_ICON_ID,
            guidItem = Guid.Empty
        };

        RECT rect;
        int hr = Shell_NotifyIconGetRect(ref identifier, out rect);
        return new TrayRect
        {
            MessageWindow = hwnd,
            HResult = hr,
            X = rect.Left,
            Y = rect.Top,
            Width = rect.Right - rect.Left,
            Height = rect.Bottom - rect.Top
        };
    }

    public static bool IsTrayIconRegistered()
    {
        IntPtr hwnd = FindMessageWindow();
        if (hwnd == IntPtr.Zero)
        {
            throw new InvalidOperationException("Paste GPUI message window was not found.");
        }

        int status = SendMessage(hwnd, WM_PASTE_GPUI_TEST_PLATFORM_STATUS, UIntPtr.Zero, IntPtr.Zero).ToInt32();
        return (status & PLATFORM_STATUS_TRAY_ICON) == PLATFORM_STATUS_TRAY_ICON;
    }

    public static long GetTrayLastError()
    {
        IntPtr hwnd = FindMessageWindow();
        if (hwnd == IntPtr.Zero)
        {
            throw new InvalidOperationException("Paste GPUI message window was not found.");
        }

        return SendMessage(hwnd, WM_PASTE_GPUI_TEST_TRAY_ERROR, UIntPtr.Zero, IntPtr.Zero).ToInt64();
    }
}
'@

if (-not ("PasteTrayRectProbe" -as [type])) {
    Add-Type -TypeDefinition $probeSource
}

$process = $null
$startedByScript = $false

try {
    if ($UseExistingProcess) {
        $process = Get-Process -Name "PasteWinUI" -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($null -eq $process) {
            throw "UseExistingProcess was specified, but PasteWinUI.exe is not running."
        }
    }
    else {
        Get-Process -Name "PasteWinUI" -ErrorAction SilentlyContinue | Stop-Process -Force
        Start-Sleep -Milliseconds 250

        $process = Start-Process -FilePath $ExePath -PassThru
        $startedByScript = $true
        Start-Sleep -Milliseconds $StartupDelayMs
        $process.Refresh()

        if ($process.HasExited) {
            throw "PasteWinUI exited during startup. ExitCode=$($process.ExitCode)"
        }
    }

    $rect = [PasteTrayRectProbe]::GetTrayRect()
    $trayIconRegistered = [PasteTrayRectProbe]::IsTrayIconRegistered()
    $trayLastError = [PasteTrayRectProbe]::GetTrayLastError()
    $rectAvailable = $rect.HResult -eq 0 -and $rect.Width -gt 0 -and $rect.Height -gt 0
    $succeeded = $trayIconRegistered

    $status = [ordered]@{
        exe = $ExePath
        process_id = $process.Id
        started_by_script = $startedByScript
        message_window_handle = $rect.MessageWindow.ToInt64()
        tray_icon_registered = $trayIconRegistered
        tray_last_error = ("0x{0:X8}" -f $trayLastError)
        shell_notify_icon_get_rect_hresult = ("0x{0:X8}" -f $rect.HResult)
        rect_available = $rectAvailable
        x = $rect.X
        y = $rect.Y
        width = $rect.Width
        height = $rect.Height
    }

    $status | ConvertTo-Json -Depth 3

    if (-not $succeeded) {
        if ($AllowUnavailable) {
            Write-Warning "Tray icon registration/rectangle was unavailable in this automation session. Physical tray acceptance is still required."
            return
        }
        throw "PasteWinUI did not report a registered tray icon. Shell_NotifyIconGetRect HRESULT=$($status.shell_notify_icon_get_rect_hresult)"
    }
}
finally {
    if (-not $KeepRunning -and $startedByScript -and $process -ne $null -and -not $process.HasExited) {
        Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
    }
}
