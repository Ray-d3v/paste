param(
    [string]$ExePath = "artifacts/local-gpui/PasteWinUI.exe",
    [int]$StartupDelayMs = 1500,
    [int]$ClipboardDelayMs = 1000,
    [int]$OverlayDelayMs = 1800,
    [int]$PasteDelayMs = 1800,
    [switch]$KeepRunning,
    [switch]$AllowForegroundUnavailable
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

$keyboardSource = @'
using System;
using System.Runtime.InteropServices;

public static class PasteImageSmokeKeyboard
{
    [DllImport("user32.dll", SetLastError = true)]
    static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

    [DllImport("user32.dll")]
    static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    static extern bool IsIconic(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

    [DllImport("user32.dll", SetLastError = true)]
    static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

    [DllImport("user32.dll", SetLastError = true)]
    static extern bool PostMessage(IntPtr hWnd, uint Msg, UIntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    static extern bool IsWindowVisible(IntPtr hWnd);

    const uint KEYEVENTF_KEYUP = 0x0002;
    const uint WM_HOTKEY = 0x0312;
    const uint WM_KEYDOWN = 0x0100;
    const uint WM_KEYUP = 0x0101;
    const int SW_RESTORE = 9;
    const int OVERLAY_HOTKEY_ID = 0x5056;
    const byte VK_CONTROL = 0x11;
    const byte VK_MENU = 0x12;
    const byte VK_V = 0x56;
    const byte VK_RETURN = 0x0D;

    public static void CtrlAltV()
    {
        Down(VK_CONTROL);
        Down(VK_MENU);
        TapHeld(VK_V);
        Up(VK_MENU);
        Up(VK_CONTROL);
    }

    public static void Enter()
    {
        Tap(VK_RETURN);
    }

    public static bool Activate(IntPtr hwnd)
    {
        if (hwnd == IntPtr.Zero)
        {
            return false;
        }

        if (IsIconic(hwnd))
        {
            ShowWindow(hwnd, SW_RESTORE);
        }

        return SetForegroundWindow(hwnd);
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

    public static bool PostHotkeyToMessageWindow()
    {
        IntPtr hwnd = FindWindowEx(new IntPtr(-3), IntPtr.Zero, "PasteGPUIMessageWindow", "Paste GPUI Message Window");
        if (hwnd == IntPtr.Zero)
        {
            hwnd = FindWindowEx(IntPtr.Zero, IntPtr.Zero, "PasteGPUIMessageWindow", "Paste GPUI Message Window");
        }
        return hwnd != IntPtr.Zero && PostMessage(hwnd, WM_HOTKEY, (UIntPtr)OVERLAY_HOTKEY_ID, IntPtr.Zero);
    }

    public static bool PostEnterToOverlay()
    {
        IntPtr hwnd = FindOverlayWindow();
        return hwnd != IntPtr.Zero
            && PostMessage(hwnd, WM_KEYDOWN, (UIntPtr)VK_RETURN, IntPtr.Zero)
            && PostMessage(hwnd, WM_KEYUP, (UIntPtr)VK_RETURN, IntPtr.Zero);
    }

    static void Tap(byte vk)
    {
        Down(vk);
        Up(vk);
    }

    static void TapHeld(byte vk)
    {
        keybd_event(vk, 0, 0, UIntPtr.Zero);
        keybd_event(vk, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
    }

    static void Down(byte vk)
    {
        keybd_event(vk, 0, 0, UIntPtr.Zero);
    }

    static void Up(byte vk)
    {
        keybd_event(vk, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
    }
}
'@

if (-not ("PasteImageSmokeKeyboard" -as [type])) {
    Add-Type -TypeDefinition $keyboardSource
}

function Activate-WindowByProcessId {
    param(
        [Parameter(Mandatory = $true)]
        [int]$ProcessId,
        [int]$Retries = 10
    )

    for ($attempt = 0; $attempt -lt $Retries; $attempt++) {
        $process = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue
        if ($process -ne $null) {
            $process.Refresh()
            if ($process.MainWindowHandle -ne [IntPtr]::Zero -and [PasteImageSmokeKeyboard]::Activate($process.MainWindowHandle)) {
                Start-Sleep -Milliseconds 250
                return $true
            }
        }

        $shell = New-Object -ComObject WScript.Shell
        if ($shell.AppActivate($ProcessId)) {
            Start-Sleep -Milliseconds 250
            return $true
        }

        Start-Sleep -Milliseconds 250
    }

    return $false
}

function Set-SmokeClipboardImage {
    param(
        [int]$Width = 16,
        [int]$Height = 12
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $bitmap = New-Object System.Drawing.Bitmap($Width, $Height, [System.Drawing.Imaging.PixelFormat]::Format24bppRgb)
    try {
        for ($y = 0; $y -lt $Height; $y++) {
            for ($x = 0; $x -lt $Width; $x++) {
                $bitmap.SetPixel($x, $y, [System.Drawing.Color]::FromArgb(220, 40, 60))
            }
        }
        [System.Windows.Forms.Clipboard]::Clear()
        [System.Windows.Forms.Clipboard]::SetImage($bitmap)
    }
    finally {
        $bitmap.Dispose()
    }
}

function Start-SmokeImageTarget {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ResultPath,
        [int]$Retries = 16
    )

    $targetScript = Join-Path ([System.IO.Path]::GetTempPath()) "paste-gpui-image-target-$([Guid]::NewGuid().ToString("N")).ps1"
    $targetSource = @'
param(
    [Parameter(Mandatory = $true)]
    [string]$ResultPath
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$source = @"
using System;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

public sealed class PasteImageSmokeForm : Form
{
    private readonly string resultPath;
    private readonly PasteImageSmokeBox box;

    public PasteImageSmokeForm(string resultPath)
    {
        this.resultPath = resultPath;
        Text = "Paste GPUI Image Smoke Target";
        Width = 420;
        Height = 280;
        StartPosition = FormStartPosition.CenterScreen;
        TopMost = false;

        box = new PasteImageSmokeBox(resultPath);
        box.Dock = DockStyle.Fill;
        Controls.Add(box);
        Shown += delegate
        {
            Activate();
            box.Focus();
        };
    }
}

public sealed class PasteImageSmokeBox : Control
{
    private const int WM_PASTE = 0x0302;
    private const uint CF_DIB = 8;
    private readonly string resultPath;

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool OpenClipboard(IntPtr hWndNewOwner);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool CloseClipboard();

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool IsClipboardFormatAvailable(uint format);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr GetClipboardData(uint uFormat);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr GlobalLock(IntPtr hMem);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool GlobalUnlock(IntPtr hMem);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern UIntPtr GlobalSize(IntPtr hMem);

    public PasteImageSmokeBox(string resultPath)
    {
        this.resultPath = resultPath;
        SetStyle(ControlStyles.Selectable, true);
        TabStop = true;
        BackColor = Color.White;
    }

    protected override void WndProc(ref Message m)
    {
        if (m.Msg == WM_PASTE)
        {
            CaptureClipboardImage();
            return;
        }

        base.WndProc(ref m);
    }

    private void CaptureClipboardImage()
    {
        string dibResult = CaptureClipboardDib();
        if (dibResult != null)
        {
            File.WriteAllText(resultPath, dibResult, Encoding.UTF8);
            return;
        }

        if (!Clipboard.ContainsImage())
        {
            File.WriteAllText(resultPath, "contains_image=false", Encoding.UTF8);
            return;
        }

        using (Image image = Clipboard.GetImage())
        using (Bitmap bitmap = new Bitmap(image))
        {
            Color pixel = bitmap.GetPixel(0, 0);
            string value = string.Format(
                "contains_image=true;width={0};height={1};pixel={2:X2}{3:X2}{4:X2}",
                bitmap.Width,
                bitmap.Height,
                pixel.R,
                pixel.G,
                pixel.B);
            File.WriteAllText(resultPath, value, Encoding.UTF8);
        }
    }

    private static string CaptureClipboardDib()
    {
        if (!OpenClipboard(IntPtr.Zero))
        {
            return null;
        }

        try
        {
            if (!IsClipboardFormatAvailable(CF_DIB))
            {
                return null;
            }

            IntPtr handle = GetClipboardData(CF_DIB);
            if (handle == IntPtr.Zero)
            {
                return null;
            }

            IntPtr ptr = GlobalLock(handle);
            if (ptr == IntPtr.Zero)
            {
                return null;
            }

            try
            {
                int size = (int)GlobalSize(handle);
                if (size < 44)
                {
                    return "contains_dib=true;invalid_size=" + size;
                }

                byte[] bytes = new byte[size];
                Marshal.Copy(ptr, bytes, 0, bytes.Length);

                int headerSize = BitConverter.ToInt32(bytes, 0);
                int width = BitConverter.ToInt32(bytes, 4);
                int rawHeight = BitConverter.ToInt32(bytes, 8);
                int height = Math.Abs(rawHeight);
                short planes = BitConverter.ToInt16(bytes, 12);
                short bitCount = BitConverter.ToInt16(bytes, 14);
                int compression = BitConverter.ToInt32(bytes, 16);

                if (planes != 1 || (compression != 0 && !(compression == 3 && bitCount == 32)) || (bitCount != 24 && bitCount != 32))
                {
                    return string.Format(
                        "contains_dib=true;width={0};height={1};bit_count={2};compression={3};unsupported=true",
                        width,
                        height,
                        bitCount,
                        compression);
                }

                int stride = ((width * bitCount + 31) / 32) * 4;
                int pixelOffset = headerSize;
                int redMask = 0x00FF0000;
                int greenMask = 0x0000FF00;
                int blueMask = 0x000000FF;
                if (compression == 3)
                {
                    if (bytes.Length < headerSize + 12)
                    {
                        return "contains_dib=true;invalid_bitfields=true";
                    }

                    redMask = BitConverter.ToInt32(bytes, headerSize);
                    greenMask = BitConverter.ToInt32(bytes, headerSize + 4);
                    blueMask = BitConverter.ToInt32(bytes, headerSize + 8);
                    pixelOffset += 12;
                }
                if (pixelOffset + stride * height > bytes.Length)
                {
                    return "contains_dib=true;invalid_pixels=true";
                }

                byte r;
                byte g;
                byte b;
                if (compression == 3)
                {
                    int pixel = BitConverter.ToInt32(bytes, pixelOffset);
                    r = ComponentFromMask(pixel, redMask);
                    g = ComponentFromMask(pixel, greenMask);
                    b = ComponentFromMask(pixel, blueMask);
                }
                else
                {
                    b = bytes[pixelOffset + 0];
                    g = bytes[pixelOffset + 1];
                    r = bytes[pixelOffset + 2];
                }
                return string.Format(
                    "contains_dib=true;width={0};height={1};bit_count={2};pixel={3:X2}{4:X2}{5:X2}",
                    width,
                    height,
                    bitCount,
                    r,
                    g,
                    b);
            }
            finally
            {
                GlobalUnlock(handle);
            }
        }
        finally
        {
            CloseClipboard();
        }
    }

    private static byte ComponentFromMask(int pixelValue, int maskValue)
    {
        uint pixel = unchecked((uint)pixelValue);
        uint mask = unchecked((uint)maskValue);
        if (mask == 0)
        {
            return 0;
        }

        int shift = 0;
        uint shifted = mask;
        while ((shifted & 1) == 0)
        {
            shifted >>= 1;
            shift++;
        }

        uint value = (pixel & mask) >> shift;
        return (byte)((value * 255 + shifted / 2) / shifted);
    }
}
"@

Add-Type -TypeDefinition $source -ReferencedAssemblies System.Windows.Forms,System.Drawing
[System.Windows.Forms.Application]::Run([PasteImageSmokeForm]::new($ResultPath))
'@
    Set-Content -LiteralPath $targetScript -Value $targetSource -Encoding UTF8

    $powershellExe = Join-Path $PSHOME "powershell.exe"
    $arguments = @(
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-STA",
        "-File",
        $targetScript,
        "-ResultPath",
        $ResultPath
    )
    $launcher = Start-Process -FilePath $powershellExe -ArgumentList $arguments -PassThru

    for ($attempt = 0; $attempt -lt $Retries; $attempt++) {
        Start-Sleep -Milliseconds 250
        $launcher.Refresh()
        if ($launcher.MainWindowHandle -ne [IntPtr]::Zero) {
            $launcher | Add-Member -NotePropertyName SmokeTargetScript -NotePropertyValue $targetScript -Force
            return $launcher
        }
    }

    $launcher.Refresh()
    $launcher | Add-Member -NotePropertyName SmokeTargetScript -NotePropertyValue $targetScript -Force
    return $launcher
}

$pasteProcess = $null
$targetProcess = $null
$targetResultPath = Join-Path ([System.IO.Path]::GetTempPath()) "paste-gpui-image-result-$([Guid]::NewGuid().ToString("N")).txt"
$capturedResult = $null
$clipboardFormatsAfterSet = @()
$pasteVerified = $false

try {
    Get-Process -Name "PasteWinUI" -ErrorAction SilentlyContinue | Stop-Process -Force
    Start-Sleep -Milliseconds 250

    $pasteProcess = Start-Process -FilePath $ExePath -PassThru
    Start-Sleep -Milliseconds $StartupDelayMs
    $pasteProcess.Refresh()

    if ($pasteProcess.HasExited) {
        throw "PasteWinUI exited during startup. ExitCode=$($pasteProcess.ExitCode)"
    }

    $targetProcess = Start-SmokeImageTarget -ResultPath $targetResultPath

    if (-not (Activate-WindowByProcessId -ProcessId $targetProcess.Id)) {
        if ($AllowForegroundUnavailable) {
            $status = [ordered]@{
                exe = $ExePath
                paste_process_id = $pasteProcess.Id
                target_process_id = $targetProcess.Id
                foreground_activation_available = $false
                paste_process_alive = -not $pasteProcess.HasExited
                target_process_alive = -not $targetProcess.HasExited
                target_result_path = $targetResultPath
            }
            $status | ConvertTo-Json -Depth 3
            Write-Warning "Unable to activate image smoke target in this automation session. Physical paste acceptance is still required."
            return
        }
        throw "Unable to activate image smoke target process $($targetProcess.Id)."
    }

    Set-SmokeClipboardImage
    $clipboardFormatsAfterSet = @([System.Windows.Forms.Clipboard]::GetDataObject().GetFormats())
    Start-Sleep -Milliseconds $ClipboardDelayMs

    if (-not (Activate-WindowByProcessId -ProcessId $targetProcess.Id)) {
        if ($AllowForegroundUnavailable) {
            $status = [ordered]@{
                exe = $ExePath
                paste_process_id = $pasteProcess.Id
                target_process_id = $targetProcess.Id
                foreground_activation_available = $false
                stage = "reactivate-before-hotkey"
                clipboard_formats_after_set = $clipboardFormatsAfterSet
                paste_process_alive = -not $pasteProcess.HasExited
                target_process_alive = -not $targetProcess.HasExited
                target_result_path = $targetResultPath
            }
            $status | ConvertTo-Json -Depth 3
            Write-Warning "Unable to reactivate image smoke target in this automation session. Physical paste acceptance is still required."
            return
        }
        throw "Unable to reactivate image smoke target before hotkey."
    }

    [PasteImageSmokeKeyboard]::CtrlAltV()
    Start-Sleep -Milliseconds $OverlayDelayMs
    if (-not [PasteImageSmokeKeyboard]::IsOverlayVisible()) {
        [void][PasteImageSmokeKeyboard]::PostHotkeyToMessageWindow()
        Start-Sleep -Milliseconds $OverlayDelayMs
    }
    [PasteImageSmokeKeyboard]::Enter()
    [void][PasteImageSmokeKeyboard]::PostEnterToOverlay()
    Start-Sleep -Milliseconds $PasteDelayMs

    if (Test-Path $targetResultPath) {
        $capturedResult = [System.IO.File]::ReadAllText($targetResultPath, [System.Text.Encoding]::UTF8)
    }

    $pasteVerified = $capturedResult -match '^contains_dib=true;width=16;height=12;bit_count=(24|32);pixel=DC283C$'

    $status = [ordered]@{
        exe = $ExePath
        paste_process_id = $pasteProcess.Id
        target_process_id = $targetProcess.Id
        captured_result = $capturedResult
        expected_result = "contains_dib=true;width=16;height=12;bit_count=(24|32);pixel=DC283C"
        clipboard_formats_after_set = $clipboardFormatsAfterSet
        paste_verified = $pasteVerified
        paste_process_alive = -not $pasteProcess.HasExited
        target_process_alive = -not $targetProcess.HasExited
        target_result_path = $targetResultPath
    }

    $status | ConvertTo-Json -Depth 3

    if (-not $pasteVerified) {
        if ($AllowForegroundUnavailable) {
            Write-Warning "Image smoke target did not receive the expected pasted image in this automation session. Physical paste acceptance is still required."
            return
        }
        throw "Image smoke target did not receive the expected pasted image."
    }
}
finally {
    if (-not $KeepRunning) {
        if ($targetProcess -ne $null -and -not $targetProcess.HasExited) {
            Stop-Process -Id $targetProcess.Id -Force -ErrorAction SilentlyContinue
        }

        if ($targetProcess -ne $null -and $targetProcess.PSObject.Properties.Name -contains "SmokeTargetScript") {
            Remove-Item -LiteralPath $targetProcess.SmokeTargetScript -Force -ErrorAction SilentlyContinue
        }

        if ($pasteProcess -ne $null -and -not $pasteProcess.HasExited) {
            Stop-Process -Id $pasteProcess.Id -Force -ErrorAction SilentlyContinue
        }

        Remove-Item -LiteralPath $targetResultPath -Force -ErrorAction SilentlyContinue
    }
}
