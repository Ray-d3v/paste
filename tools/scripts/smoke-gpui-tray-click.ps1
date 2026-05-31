param(
    [string]$ExePath = "artifacts/local-gpui/PasteWinUI.exe",
    [int]$StartupDelayMs = 1500,
    [int]$ClickDelayMs = 700,
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
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

public static class PasteTrayClickProbe
{
    const int TRAY_ICON_ID = 0x5057;
    const uint INPUT_MOUSE = 0;
    const uint INPUT_KEYBOARD = 1;
    const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
    const uint MOUSEEVENTF_LEFTUP = 0x0004;
    const uint MOUSEEVENTF_RIGHTDOWN = 0x0008;
    const uint MOUSEEVENTF_RIGHTUP = 0x0010;
    const uint KEYEVENTF_KEYUP = 0x0002;
    const ushort VK_ESCAPE = 0x1B;

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
        public MOUSEINPUT mi;
        [FieldOffset(0)]
        public KEYBDINPUT ki;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct MOUSEINPUT
    {
        public int dx;
        public int dy;
        public uint mouseData;
        public uint dwFlags;
        public uint time;
        public IntPtr dwExtraInfo;
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

    public class TrayRect
    {
        public IntPtr MessageWindow;
        public int HResult;
        public int X;
        public int Y;
        public int Width;
        public int Height;
        public int CenterX;
        public int CenterY;
    }

    public class WindowInfo
    {
        public IntPtr Hwnd;
        public string ClassName;
        public string Title;
        public int X;
        public int Y;
        public int Width;
        public int Height;
        public bool Visible;
    }

    delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);

    [DllImport("shell32.dll", SetLastError = true)]
    static extern int Shell_NotifyIconGetRect(ref NOTIFYICONIDENTIFIER identifier, out RECT iconLocation);

    [DllImport("user32.dll", SetLastError = true)]
    static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

    [DllImport("user32.dll", SetLastError = true)]
    static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

    [DllImport("user32.dll")]
    static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll", SetLastError = true)]
    static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

    [DllImport("user32.dll", SetLastError = true)]
    static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

    [DllImport("user32.dll")]
    static extern bool SetCursorPos(int X, int Y);

    [DllImport("user32.dll")]
    static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

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
            Height = rect.Bottom - rect.Top,
            CenterX = rect.Left + ((rect.Right - rect.Left) / 2),
            CenterY = rect.Top + ((rect.Bottom - rect.Top) / 2)
        };
    }

    public static uint LeftClick(int x, int y)
    {
        return Click(x, y, MOUSEEVENTF_LEFTDOWN, MOUSEEVENTF_LEFTUP);
    }

    public static uint RightClick(int x, int y)
    {
        return Click(x, y, MOUSEEVENTF_RIGHTDOWN, MOUSEEVENTF_RIGHTUP);
    }

    public static void Escape()
    {
        INPUT[] inputs = new INPUT[]
        {
            Key(VK_ESCAPE, 0),
            Key(VK_ESCAPE, KEYEVENTF_KEYUP),
        };
        uint sent = SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT)));
        if (sent != 2)
        {
            keybd_event((byte)VK_ESCAPE, 0, 0, UIntPtr.Zero);
            keybd_event((byte)VK_ESCAPE, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
        }
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
                X = rect.Left,
                Y = rect.Top,
                Width = rect.Right - rect.Left,
                Height = rect.Bottom - rect.Top,
                Visible = IsWindowVisible(hwnd)
            });
            return true;
        }, IntPtr.Zero);
        return results.ToArray();
    }

    static uint Click(int x, int y, uint downFlag, uint upFlag)
    {
        SetCursorPos(x, y);
        INPUT[] inputs = new INPUT[]
        {
            Mouse(downFlag),
            Mouse(upFlag),
        };
        return SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT)));
    }

    static INPUT Mouse(uint flags)
    {
        return new INPUT
        {
            type = INPUT_MOUSE,
            U = new INPUTUNION
            {
                mi = new MOUSEINPUT
                {
                    dx = 0,
                    dy = 0,
                    mouseData = 0,
                    dwFlags = flags,
                    time = 0,
                    dwExtraInfo = IntPtr.Zero
                }
            }
        };
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

if (-not ("PasteTrayClickProbe" -as [type])) {
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

    $rect = [PasteTrayClickProbe]::GetTrayRect()
    if ($rect.HResult -ne 0 -or $rect.Width -le 0 -or $rect.Height -le 0) {
        throw "Shell_NotifyIconGetRect did not return a usable tray icon rectangle."
    }

    $leftSent = [PasteTrayClickProbe]::LeftClick($rect.CenterX, $rect.CenterY)
    Start-Sleep -Milliseconds $ClickDelayMs
    $leftClickShowedOverlay = [PasteTrayClickProbe]::IsOverlayVisible()

    [PasteTrayClickProbe]::Escape()
    Start-Sleep -Milliseconds 250

    $rightSent = [PasteTrayClickProbe]::RightClick($rect.CenterX, $rect.CenterY)
    Start-Sleep -Milliseconds $ClickDelayMs
    $menus = @([PasteTrayClickProbe]::EnumeratePopupMenus())
    $visibleMenus = @($menus | Where-Object { $_.Visible -and $_.Width -gt 0 -and $_.Height -gt 0 })

    [PasteTrayClickProbe]::Escape()
    Start-Sleep -Milliseconds 250

    $status = [ordered]@{
        exe = $ExePath
        process_id = $process.Id
        message_window_handle = $rect.MessageWindow.ToInt64()
        tray_rect_hresult = ("0x{0:X8}" -f $rect.HResult)
        tray_x = $rect.X
        tray_y = $rect.Y
        tray_width = $rect.Width
        tray_height = $rect.Height
        click_x = $rect.CenterX
        click_y = $rect.CenterY
        left_click_send_input_events = $leftSent
        left_click_showed_overlay = $leftClickShowedOverlay
        right_click_send_input_events = $rightSent
        popup_menu_count = $visibleMenus.Count
        popup_menus = @($visibleMenus | ForEach-Object {
            [pscustomobject]@{
                hwnd = ("0x{0:X}" -f $_.Hwnd.ToInt64())
                class_name = $_.ClassName
                title = $_.Title
                x = $_.X
                y = $_.Y
                width = $_.Width
                height = $_.Height
            }
        })
    }

    $status | ConvertTo-Json -Depth 5

    if ($leftSent -lt 2) {
        throw "Left click SendInput sent too few events: $leftSent"
    }

    if (-not $leftClickShowedOverlay) {
        throw "Left-clicking the tray icon did not show the overlay."
    }

    if ($rightSent -lt 2) {
        throw "Right click SendInput sent too few events: $rightSent"
    }

    if ($visibleMenus.Count -lt 1) {
        throw "Right-clicking the tray icon did not open a visible #32768 popup menu."
    }
}
finally {
    if (-not $KeepRunning -and $process -ne $null -and -not $process.HasExited) {
        Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
    }
}
