param(
    [string]$ExePath = "artifacts/local-gpui/PasteWinUI.exe",
    [int]$StartupDelayMs = 1500,
    [int]$EventDelayMs = 700,
    [switch]$KeepRunning,
    [switch]$AllowPopupUnavailable
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
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

public static class PasteTrayNotificationProbe
{
    const uint WM_TRAY_ICON = 0xD057;
    const uint NIN_SELECT = 0x0400;
    const uint WM_CONTEXTMENU = 0x007B;
    const uint VK_ESCAPE = 0x1B;

    [StructLayout(LayoutKind.Sequential)]
    struct INPUT
    {
        public uint type;
        public INPUTUNION U;
    }

    [StructLayout(LayoutKind.Explicit)]
    struct INPUTUNION
    {
        [FieldOffset(0)]
        public KEYBDINPUT ki;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct KEYBDINPUT
    {
        public ushort wVk;
        public ushort wScan;
        public uint dwFlags;
        public uint time;
        public IntPtr dwExtraInfo;
    }

    public class WindowInfo
    {
        public IntPtr Hwnd;
        public string ClassName;
        public string Title;
        public int Width;
        public int Height;
        public bool Visible;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

    [DllImport("user32.dll", SetLastError = true)]
    static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

    [DllImport("user32.dll", SetLastError = true)]
    static extern bool PostMessage(IntPtr hWnd, uint Msg, UIntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", SetLastError = true)]
    static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

    const uint INPUT_KEYBOARD = 1;
    const uint KEYEVENTF_KEYUP = 0x0002;

    public static IntPtr FindMessageWindow()
    {
        IntPtr hwnd = FindWindowEx(new IntPtr(-3), IntPtr.Zero, "PasteGPUIMessageWindow", "Paste GPUI Message Window");
        return hwnd != IntPtr.Zero
            ? hwnd
            : FindWindowEx(IntPtr.Zero, IntPtr.Zero, "PasteGPUIMessageWindow", "Paste GPUI Message Window");
    }

    public static IntPtr FindOverlayWindow()
    {
        return FindWindow(null, "Paste GPUI Overlay");
    }

    public static bool IsOverlayVisible()
    {
        IntPtr hwnd = FindOverlayWindow();
        return hwnd != IntPtr.Zero && IsWindowVisible(hwnd);
    }

    public static bool PostSelect(IntPtr hwnd)
    {
        return PostMessage(hwnd, WM_TRAY_ICON, UIntPtr.Zero, (IntPtr)NIN_SELECT);
    }

    public static bool PostContextMenu(IntPtr hwnd)
    {
        return PostMessage(hwnd, WM_TRAY_ICON, UIntPtr.Zero, (IntPtr)WM_CONTEXTMENU);
    }

    public static void Escape()
    {
        INPUT[] inputs = new INPUT[]
        {
            Key((ushort)VK_ESCAPE, 0),
            Key((ushort)VK_ESCAPE, KEYEVENTF_KEYUP),
        };
        SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT)));
    }

    public static WindowInfo[] EnumeratePopupMenus()
    {
        var results = new List<WindowInfo>();
        EnumWindows((hwnd, _) =>
        {
            string className = ReadClassName(hwnd);
            if (className != "#32768")
            {
                return true;
            }

            RECT rect;
            GetWindowRect(hwnd, out rect);
            results.Add(new WindowInfo
            {
                Hwnd = hwnd,
                ClassName = className,
                Title = ReadTitle(hwnd),
                Width = rect.Right - rect.Left,
                Height = rect.Bottom - rect.Top,
                Visible = IsWindowVisible(hwnd)
            });
            return true;
        }, IntPtr.Zero);
        return results.ToArray();
    }

    static INPUT Key(ushort vk, uint flags)
    {
        return new INPUT
        {
            type = INPUT_KEYBOARD,
            U = new INPUTUNION
            {
                ki = new KEYBDINPUT
                {
                    wVk = vk,
                    wScan = 0,
                    dwFlags = flags,
                    time = 0,
                    dwExtraInfo = IntPtr.Zero
                }
            }
        };
    }

    static string ReadClassName(IntPtr hwnd)
    {
        var buffer = new StringBuilder(256);
        GetClassName(hwnd, buffer, buffer.Capacity);
        return buffer.ToString();
    }

    static string ReadTitle(IntPtr hwnd)
    {
        var buffer = new StringBuilder(512);
        GetWindowText(hwnd, buffer, buffer.Capacity);
        return buffer.ToString();
    }
}
'@

if (-not ("PasteTrayNotificationProbe" -as [type])) {
    Add-Type -TypeDefinition $probeSource
}

Get-Process -Name "PasteWinUI" -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Milliseconds 250

$process = $null

try {
    $process = Start-Process -FilePath $ExePath -PassThru
    Start-Sleep -Milliseconds $StartupDelayMs
    $process.Refresh()

    if ($process.HasExited) {
        throw "PasteWinUI exited during startup. ExitCode=$($process.ExitCode)"
    }

    $messageWindow = [PasteTrayNotificationProbe]::FindMessageWindow()
    if ($messageWindow -eq [IntPtr]::Zero) {
        throw "Paste GPUI message window was not found."
    }

    $selectPosted = [PasteTrayNotificationProbe]::PostSelect($messageWindow)
    Start-Sleep -Milliseconds $EventDelayMs
    $overlayVisibleAfterSelect = [PasteTrayNotificationProbe]::IsOverlayVisible()

    [PasteTrayNotificationProbe]::Escape()
    Start-Sleep -Milliseconds 250

    $contextPosted = [PasteTrayNotificationProbe]::PostContextMenu($messageWindow)
    Start-Sleep -Milliseconds $EventDelayMs
    $menus = @([PasteTrayNotificationProbe]::EnumeratePopupMenus())
    $visibleMenus = @($menus | Where-Object { $_.Visible -and $_.Width -gt 0 -and $_.Height -gt 0 })

    [PasteTrayNotificationProbe]::Escape()

    $status = [ordered]@{
        exe = $ExePath
        process_id = $process.Id
        message_window_handle = $messageWindow.ToInt64()
        nin_select_posted = $selectPosted
        overlay_visible_after_nin_select = $overlayVisibleAfterSelect
        wm_contextmenu_posted = $contextPosted
        popup_menu_count = $visibleMenus.Count
    }

    $status | ConvertTo-Json -Depth 3

    if (-not $selectPosted) {
        throw "Failed to post NIN_SELECT tray notification."
    }

    if (-not $overlayVisibleAfterSelect) {
        throw "NIN_SELECT did not show the overlay."
    }

    if (-not $contextPosted) {
        throw "Failed to post WM_CONTEXTMENU tray notification."
    }

    if ($visibleMenus.Count -lt 1) {
        if ($AllowPopupUnavailable) {
            Write-Warning "WM_CONTEXTMENU was posted, but no visible #32768 popup menu was observed in this automation session. Physical tray menu acceptance is still required."
            return
        }
        throw "WM_CONTEXTMENU did not open a visible #32768 popup menu."
    }
}
finally {
    if (-not $KeepRunning -and $process -ne $null -and -not $process.HasExited) {
        Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
    }
}
