param(
    [string]$ExePath = "artifacts/local-gpui/PasteWinUI.exe",
    [int]$StartupDelayMs = 1500,
    [int]$ActionDelayMs = 500,
    [switch]$AllowSyntheticInputUnavailable
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

$source = @'
using System;
using System.Runtime.InteropServices;

public static class PasteCommandPaletteSmoke
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

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

    [DllImport("user32.dll", SetLastError = true)]
    static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

    [DllImport("user32.dll", SetLastError = true)]
    static extern bool PostMessage(IntPtr hWnd, uint Msg, UIntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true)]
    static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

    [DllImport("user32.dll", SetLastError = true)]
    static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

    const uint INPUT_KEYBOARD = 1;
    const uint KEYEVENTF_KEYUP = 0x0002;
    const ushort VK_CONTROL = 0x11;
    const ushort VK_MENU = 0x12;
    const ushort VK_V = 0x56;
    const ushort VK_K = 0x4B;
    const ushort VK_ESCAPE = 0x1B;
    const ushort VK_DOWN = 0x28;
    const ushort VK_RETURN = 0x0D;
    const uint WM_HOTKEY = 0x0312;
    const uint WM_KEYDOWN = 0x0100;
    const uint WM_KEYUP = 0x0101;
    const int OVERLAY_HOTKEY_ID = 0x5056;

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

    public static bool PostCtrlK()
    {
        IntPtr hwnd = FindOverlayWindow();
        return hwnd != IntPtr.Zero
            && PostMessage(hwnd, WM_KEYDOWN, (UIntPtr)VK_CONTROL, IntPtr.Zero)
            && PostMessage(hwnd, WM_KEYDOWN, (UIntPtr)VK_K, IntPtr.Zero)
            && PostMessage(hwnd, WM_KEYUP, (UIntPtr)VK_K, IntPtr.Zero)
            && PostMessage(hwnd, WM_KEYUP, (UIntPtr)VK_CONTROL, IntPtr.Zero);
    }

    public static bool PostEscape()
    {
        return PostKey(VK_ESCAPE);
    }

    public static bool PostDown()
    {
        return PostKey(VK_DOWN);
    }

    public static bool PostEnter()
    {
        return PostKey(VK_RETURN);
    }

    static bool PostKey(ushort vk)
    {
        IntPtr hwnd = FindOverlayWindow();
        return hwnd != IntPtr.Zero
            && PostMessage(hwnd, WM_KEYDOWN, (UIntPtr)vk, IntPtr.Zero)
            && PostMessage(hwnd, WM_KEYUP, (UIntPtr)vk, IntPtr.Zero);
    }

    public static uint SendCtrlAltV()
    {
        return SendChord(VK_CONTROL, VK_MENU, VK_V);
    }

    public static uint SendCtrlK()
    {
        return SendChord(VK_CONTROL, 0, VK_K);
    }

    public static uint SendEscape()
    {
        INPUT[] inputs = new INPUT[] { Key(VK_ESCAPE, 0), Key(VK_ESCAPE, KEYEVENTF_KEYUP) };
        return SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT)));
    }

    public static uint SendDown()
    {
        INPUT[] inputs = new INPUT[] { Key(VK_DOWN, 0), Key(VK_DOWN, KEYEVENTF_KEYUP) };
        return SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT)));
    }

    public static uint SendEnter()
    {
        INPUT[] inputs = new INPUT[] { Key(VK_RETURN, 0), Key(VK_RETURN, KEYEVENTF_KEYUP) };
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

    public static void SendCtrlKLegacy()
    {
        keybd_event((byte)VK_CONTROL, 0, 0, UIntPtr.Zero);
        keybd_event((byte)VK_K, 0, 0, UIntPtr.Zero);
        keybd_event((byte)VK_K, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
        keybd_event((byte)VK_CONTROL, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
    }

    public static void SendEscapeLegacy()
    {
        keybd_event((byte)VK_ESCAPE, 0, 0, UIntPtr.Zero);
        keybd_event((byte)VK_ESCAPE, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
    }

    public static void SendDownLegacy()
    {
        keybd_event((byte)VK_DOWN, 0, 0, UIntPtr.Zero);
        keybd_event((byte)VK_DOWN, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
    }

    public static void SendEnterLegacy()
    {
        keybd_event((byte)VK_RETURN, 0, 0, UIntPtr.Zero);
        keybd_event((byte)VK_RETURN, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
    }

    static uint SendChord(ushort modifier1, ushort modifier2, ushort key)
    {
        INPUT[] inputs;
        if (modifier2 == 0)
        {
            inputs = new INPUT[]
            {
                Key(modifier1, 0),
                Key(key, 0),
                Key(key, KEYEVENTF_KEYUP),
                Key(modifier1, KEYEVENTF_KEYUP),
            };
        }
        else
        {
            inputs = new INPUT[]
            {
                Key(modifier1, 0),
                Key(modifier2, 0),
                Key(key, 0),
                Key(key, KEYEVENTF_KEYUP),
                Key(modifier2, KEYEVENTF_KEYUP),
                Key(modifier1, KEYEVENTF_KEYUP),
            };
        }
        return SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT)));
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

if (-not ("PasteCommandPaletteSmoke" -as [type])) {
    Add-Type -TypeDefinition $source
}

function Send-CtrlK {
    $sent = [PasteCommandPaletteSmoke]::SendCtrlK()
    if ($sent -ne 4) {
        [PasteCommandPaletteSmoke]::SendCtrlKLegacy()
        [void][PasteCommandPaletteSmoke]::PostCtrlK()
    }
    return $sent
}

function Send-Escape {
    $sent = [PasteCommandPaletteSmoke]::SendEscape()
    if ($sent -ne 2) {
        [PasteCommandPaletteSmoke]::SendEscapeLegacy()
        [void][PasteCommandPaletteSmoke]::PostEscape()
    }
    return $sent
}

function Send-Down {
    $sent = [PasteCommandPaletteSmoke]::SendDown()
    if ($sent -ne 2) {
        [PasteCommandPaletteSmoke]::SendDownLegacy()
        [void][PasteCommandPaletteSmoke]::PostDown()
    }
    return $sent
}

function Send-Enter {
    $sent = [PasteCommandPaletteSmoke]::SendEnter()
    if ($sent -ne 2) {
        [PasteCommandPaletteSmoke]::SendEnterLegacy()
        [void][PasteCommandPaletteSmoke]::PostEnter()
    }
    return $sent
}

Get-Process -Name "PasteWinUI" -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Milliseconds 250

$process = Start-Process -FilePath $ExePath -PassThru
Start-Sleep -Milliseconds $StartupDelayMs
$process.Refresh()

if ($process.HasExited) {
    throw "PasteWinUI exited during startup. ExitCode=$($process.ExitCode)"
}

$sentOpen = [PasteCommandPaletteSmoke]::SendCtrlAltV()
if ($sentOpen -ne 6) {
    [PasteCommandPaletteSmoke]::SendCtrlAltVLegacy()
}
Start-Sleep -Milliseconds $ActionDelayMs
$visibleAfterOpen = [PasteCommandPaletteSmoke]::IsOverlayVisible()
$postedHotkeyMessage = $false
if (-not $visibleAfterOpen) {
    $postedHotkeyMessage = [PasteCommandPaletteSmoke]::PostHotkeyToMessageWindow()
    Start-Sleep -Milliseconds $ActionDelayMs
    $visibleAfterOpen = [PasteCommandPaletteSmoke]::IsOverlayVisible()
}

$sentPalette = Send-CtrlK
Start-Sleep -Milliseconds $ActionDelayMs
$visibleAfterPaletteOpen = [PasteCommandPaletteSmoke]::IsOverlayVisible()

[void](Send-Escape)
Start-Sleep -Milliseconds $ActionDelayMs
$visibleAfterPaletteEscape = [PasteCommandPaletteSmoke]::IsOverlayVisible()

$sentPaletteSecond = Send-CtrlK
Start-Sleep -Milliseconds $ActionDelayMs
[void](Send-Down)
Start-Sleep -Milliseconds 100
[void](Send-Down)
Start-Sleep -Milliseconds 100
[void](Send-Down)
Start-Sleep -Milliseconds 100
[void](Send-Down)
Start-Sleep -Milliseconds 100
[void](Send-Enter)
Start-Sleep -Milliseconds $ActionDelayMs
$visibleAfterCommandEnter = [PasteCommandPaletteSmoke]::IsOverlayVisible()

[void](Send-Escape)
Start-Sleep -Milliseconds $ActionDelayMs
$visibleAfterOverlayEscape = [PasteCommandPaletteSmoke]::IsOverlayVisible()

$process.Refresh()
$aliveAfterActions = -not $process.HasExited

Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue

$status = [ordered]@{
    exe = $ExePath
    process_id = $process.Id
    alive_after_actions = $aliveAfterActions
    visible_after_open = $visibleAfterOpen
    visible_after_palette_open = $visibleAfterPaletteOpen
    visible_after_palette_escape = $visibleAfterPaletteEscape
    visible_after_command_enter = $visibleAfterCommandEnter
    visible_after_overlay_escape = $visibleAfterOverlayEscape
    send_input_open_events = $sentOpen
    posted_hotkey_message = $postedHotkeyMessage
    send_input_palette_events = $sentPalette
    send_input_palette_second_events = $sentPaletteSecond
}

$status | ConvertTo-Json -Depth 4

if (-not $aliveAfterActions) {
    throw "PasteWinUI exited during command palette smoke."
}

if (-not $visibleAfterOpen) {
    throw "Overlay was not visible after Ctrl+Alt+V."
}

if (-not $visibleAfterPaletteOpen) {
    if ($AllowSyntheticInputUnavailable -and $sentPalette -eq 0) {
        Write-Warning "Ctrl+K synthetic input was unavailable in this automation session. Command palette keyboard acceptance remains physical/manual."
        return
    }
    throw "Overlay was not visible after Ctrl+K opened the command palette."
}

if (-not $visibleAfterPaletteEscape) {
    if ($AllowSyntheticInputUnavailable -and ($sentPalette -eq 0 -or $sentPaletteSecond -eq 0)) {
        Write-Warning "Command palette key sequence could not be completed because synthetic input was unavailable in this automation session."
        return
    }
    throw "Overlay was hidden by Esc while the command palette should have closed first."
}

if (-not $visibleAfterCommandEnter) {
    throw "Overlay was hidden after command palette Enter; expected Toggle Search command to keep it visible."
}

if ($visibleAfterOverlayEscape) {
    throw "Overlay remained visible after Esc with command palette closed."
}
