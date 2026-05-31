param(
    [string]$ExePath = "artifacts/local-gpui/PasteWinUI.exe",
    [int]$StartupDelayMs = 1500,
    [int]$HotkeyDelayMs = 1200,
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

$sendInputSource = @'
using System;
using System.Runtime.InteropServices;

public static class PasteSmokeKeyboard
{
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

    [DllImport("user32.dll", SetLastError = true)]
    static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

    [DllImport("user32.dll", SetLastError = true)]
    static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

    const uint INPUT_KEYBOARD = 1;
    const uint KEYEVENTF_KEYUP = 0x0002;
    const ushort VK_CONTROL = 0x11;
    const ushort VK_MENU = 0x12;
    const ushort VK_V = 0x56;

    public static uint SendCtrlAltV()
    {
        INPUT[] inputs = new INPUT[]
        {
            Key(VK_CONTROL, 0),
            Key(VK_MENU, 0),
            Key(VK_V, 0),
            Key(VK_V, KEYEVENTF_KEYUP),
            Key(VK_MENU, KEYEVENTF_KEYUP),
            Key(VK_CONTROL, KEYEVENTF_KEYUP),
        };
        return SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT)));
    }

    public static void SendCtrlAltVLegacy()
    {
        keybd_event((byte)VK_CONTROL, 0, 0, UIntPtr.Zero);
        keybd_event((byte)VK_MENU, 0, 0, UIntPtr.Zero);
        keybd_event((byte)VK_V, 0, 0, UIntPtr.Zero);
        keybd_event((byte)VK_V, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
        keybd_event((byte)VK_MENU, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
        keybd_event((byte)VK_CONTROL, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
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
}
'@

if (-not ("PasteSmokeKeyboard" -as [type])) {
    Add-Type -TypeDefinition $sendInputSource
}

$windowProbeSource = @'
using System;
using System.Runtime.InteropServices;

public static class PasteSmokeWindowProbe
{
    const uint WM_HOTKEY = 0x0312;
    const int OVERLAY_HOTKEY_ID = 0x5056;

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

    public static bool PostHotkeyToMessageWindow()
    {
        IntPtr hwnd = FindMessageWindow();
        return hwnd != IntPtr.Zero && PostMessage(hwnd, WM_HOTKEY, (UIntPtr)OVERLAY_HOTKEY_ID, IntPtr.Zero);
    }
}
'@

if (-not ("PasteSmokeWindowProbe" -as [type])) {
    Add-Type -TypeDefinition $windowProbeSource
}

Get-Process -Name "PasteWinUI" -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Milliseconds 250

$process = Start-Process -FilePath $ExePath -PassThru
Start-Sleep -Milliseconds $StartupDelayMs
$process.Refresh()

if ($process.HasExited) {
    throw "PasteWinUI exited during startup. ExitCode=$($process.ExitCode)"
}

$overlayHandleBefore = [PasteSmokeWindowProbe]::FindOverlayWindow()
$overlayVisibleBefore = [PasteSmokeWindowProbe]::IsOverlayVisible()

$sent = [PasteSmokeKeyboard]::SendCtrlAltV()
if ($sent -ne 6) {
    [PasteSmokeKeyboard]::SendCtrlAltVLegacy()
}
Start-Sleep -Milliseconds $HotkeyDelayMs
$process.Refresh()
$overlayHandleAfter = [PasteSmokeWindowProbe]::FindOverlayWindow()
$overlayVisibleAfter = [PasteSmokeWindowProbe]::IsOverlayVisible()
$postedHotkeyMessage = $false
if (-not $overlayVisibleAfter) {
    $postedHotkeyMessage = [PasteSmokeWindowProbe]::PostHotkeyToMessageWindow()
    Start-Sleep -Milliseconds $HotkeyDelayMs
    $process.Refresh()
    $overlayHandleAfter = [PasteSmokeWindowProbe]::FindOverlayWindow()
    $overlayVisibleAfter = [PasteSmokeWindowProbe]::IsOverlayVisible()
}

$status = [ordered]@{
    exe = $ExePath
    process_id = $process.Id
    alive_after_hotkey = -not $process.HasExited
    send_input_events = $sent
    expected_send_input_events = 6
    used_legacy_keybd_event = $sent -ne 6
    posted_hotkey_message = $postedHotkeyMessage
    main_window_handle = $process.MainWindowHandle.ToInt64()
    overlay_window_handle_before_hotkey = $overlayHandleBefore.ToInt64()
    overlay_visible_before_hotkey = $overlayVisibleBefore
    overlay_window_handle_after_hotkey = $overlayHandleAfter.ToInt64()
    overlay_visible_after_hotkey = $overlayVisibleAfter
}

if (-not $KeepRunning) {
    Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
}

$status | ConvertTo-Json -Depth 3

if (-not $status.alive_after_hotkey) {
    throw "PasteWinUI exited after Ctrl+Alt+V."
}

if ($status.overlay_visible_before_hotkey) {
    throw "Overlay was visible before Ctrl+Alt+V."
}

if (-not $status.overlay_visible_after_hotkey) {
    throw "Overlay was not visible after Ctrl+Alt+V."
}
