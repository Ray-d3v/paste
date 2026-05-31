param(
    [string]$ExePath = "artifacts/local-gpui/PasteWinUI.exe",
    [int]$StartupDelayMs = 1500,
    [int]$SampleCount = 35,
    [int]$SampleIntervalMs = 10,
    [int]$ExpectedTravelPixels = 40
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

public static class PasteSlideProbe
{
    [StructLayout(LayoutKind.Sequential)]
    struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

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

    [DllImport("user32.dll", SetLastError = true)]
    static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

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
    const uint WM_HOTKEY = 0x0312;
    const int OVERLAY_HOTKEY_ID = 0x5056;

    public static IntPtr FindOverlayWindow()
    {
        return FindWindow(null, "Paste GPUI Overlay");
    }

    public static int[] GetOverlayRect()
    {
        IntPtr hwnd = FindOverlayWindow();
        if (hwnd == IntPtr.Zero || !IsWindowVisible(hwnd))
        {
            return null;
        }

        RECT rect;
        if (!GetWindowRect(hwnd, out rect))
        {
            return null;
        }

        return new int[] { rect.Left, rect.Top, rect.Right, rect.Bottom };
    }

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

if (-not ("PasteSlideProbe" -as [type])) {
    Add-Type -TypeDefinition $probeSource
}

Get-Process -Name "PasteWinUI" -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Milliseconds 250

$process = Start-Process -FilePath $ExePath -PassThru
Start-Sleep -Milliseconds $StartupDelayMs
$process.Refresh()

if ($process.HasExited) {
    throw "PasteWinUI exited during startup. ExitCode=$($process.ExitCode)"
}

$sent = [PasteSlideProbe]::SendCtrlAltV()
if ($sent -ne 6) {
    [PasteSlideProbe]::SendCtrlAltVLegacy()
}
$postedHotkeyMessage = $false
Start-Sleep -Milliseconds 80
if ($null -eq [PasteSlideProbe]::GetOverlayRect()) {
    $postedHotkeyMessage = [PasteSlideProbe]::PostHotkeyToMessageWindow()
}

$samples = New-Object System.Collections.Generic.List[object]
for ($i = 0; $i -lt $SampleCount; $i++) {
    $rect = [PasteSlideProbe]::GetOverlayRect()
    $visible = $null -ne $rect
    $samples.Add([pscustomobject]@{
        index = $i
        elapsed_ms = $i * $SampleIntervalMs
        visible = $visible
        top = if ($visible) { $rect[1] } else { $null }
        bottom = if ($visible) { $rect[3] } else { $null }
        height = if ($visible) { $rect[3] - $rect[1] } else { $null }
    })
    Start-Sleep -Milliseconds $SampleIntervalMs
}

$process.Refresh()
$visibleSamples = @($samples | Where-Object { $_.visible -and $_.top -ne $null })

$firstTop = if ($visibleSamples.Count -gt 0) { [int]$visibleSamples[0].top } else { $null }
$minTop = if ($visibleSamples.Count -gt 0) { [int]($visibleSamples | Measure-Object -Property top -Minimum).Minimum } else { $null }
$lastTop = if ($visibleSamples.Count -gt 0) { [int]$visibleSamples[-1].top } else { $null }
$travelPixels = if ($firstTop -ne $null -and $minTop -ne $null) { $firstTop - $minTop } else { 0 }
$endsNearMinimum = if ($lastTop -ne $null -and $minTop -ne $null) { [Math]::Abs($lastTop - $minTop) -le 8 } else { $false }
$aliveAfterHotkey = -not $process.HasExited

Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue

$status = [ordered]@{
    exe = $ExePath
    process_id = $process.Id
    alive_after_hotkey = $aliveAfterHotkey
    send_input_events = $sent
    expected_send_input_events = 6
    used_legacy_keybd_event = $sent -ne 6
    posted_hotkey_message = $postedHotkeyMessage
    visible_sample_count = $visibleSamples.Count
    first_top = $firstTop
    min_top = $minTop
    last_top = $lastTop
    travel_pixels = $travelPixels
    expected_travel_pixels = $ExpectedTravelPixels
    ends_near_minimum = $endsNearMinimum
    samples = @($samples.ToArray())
}

$status | ConvertTo-Json -Depth 5

if (-not $status.alive_after_hotkey) {
    throw "PasteWinUI exited after Ctrl+Alt+V."
}

if ($visibleSamples.Count -lt 3) {
    throw "Too few visible overlay samples captured."
}

if ($travelPixels -lt $ExpectedTravelPixels) {
    throw "Overlay did not slide upward enough. Travel=$travelPixels Expected=$ExpectedTravelPixels"
}

if (-not $endsNearMinimum) {
    throw "Overlay did not settle near its highest sampled position. Last=$lastTop Min=$minTop"
}
