param(
    [string]$ExePath = "artifacts/local-gpui/PasteWinUI.exe",
    [int]$StartupDelayMs = 1500,
    [int]$ClipboardDelayMs = 1000,
    [int]$OverlayDelayMs = 1200,
    [int]$PasteDelayMs = 1200,
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

public static class PasteTextSmokeKeyboard
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
    const byte VK_A = 0x41;
    const byte VK_C = 0x43;
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

    public static void CtrlA()
    {
        Down(VK_CONTROL);
        TapHeld(VK_A);
        Up(VK_CONTROL);
    }

    public static void CtrlC()
    {
        Down(VK_CONTROL);
        TapHeld(VK_C);
        Up(VK_CONTROL);
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

if (-not ("PasteTextSmokeKeyboard" -as [type])) {
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
            if ($process.MainWindowHandle -ne [IntPtr]::Zero -and [PasteTextSmokeKeyboard]::Activate($process.MainWindowHandle)) {
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

function Start-SmokeTextTarget {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ResultPath,
        [int]$Retries = 16
    )

    $targetScript = Join-Path ([System.IO.Path]::GetTempPath()) "paste-gpui-smoke-target-$([Guid]::NewGuid().ToString("N")).ps1"
    $targetSource = @'
param(
    [Parameter(Mandatory = $true)]
    [string]$ResultPath
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = "Paste GPUI Smoke Target"
$form.Width = 640
$form.Height = 360
$form.StartPosition = "CenterScreen"
$form.TopMost = $false

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Multiline = $true
$textBox.AcceptsReturn = $true
$textBox.AcceptsTab = $true
$textBox.Dock = [System.Windows.Forms.DockStyle]::Fill
$textBox.Font = New-Object System.Drawing.Font("Consolas", 12)
$form.Controls.Add($textBox)

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 200
$timer.Add_Tick({
    [System.IO.File]::WriteAllText($ResultPath, $textBox.Text, [System.Text.Encoding]::UTF8)
})
$timer.Start()

$form.Add_Shown({
    $form.Activate()
    $textBox.Focus()
})

[System.Windows.Forms.Application]::Run($form)
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
$targetResultPath = Join-Path ([System.IO.Path]::GetTempPath()) "paste-gpui-smoke-result-$([Guid]::NewGuid().ToString("N")).txt"
$uniqueText = "paste-gpui-smoke-$([Guid]::NewGuid().ToString("N"))"
$capturedText = $null
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

    $targetProcess = Start-SmokeTextTarget -ResultPath $targetResultPath

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
            Write-Warning "Unable to activate smoke target in this automation session. Physical paste acceptance is still required."
            return
        }
        throw "Unable to activate smoke target process $($targetProcess.Id)."
    }

    Set-Clipboard -Value $uniqueText
    Start-Sleep -Milliseconds $ClipboardDelayMs

    if (-not (Activate-WindowByProcessId -ProcessId $targetProcess.Id)) {
        if ($AllowForegroundUnavailable) {
            $status = [ordered]@{
                exe = $ExePath
                paste_process_id = $pasteProcess.Id
                target_process_id = $targetProcess.Id
                foreground_activation_available = $false
                stage = "reactivate-before-hotkey"
                unique_text = $uniqueText
                paste_process_alive = -not $pasteProcess.HasExited
                target_process_alive = -not $targetProcess.HasExited
                target_result_path = $targetResultPath
            }
            $status | ConvertTo-Json -Depth 3
            Write-Warning "Unable to reactivate smoke target in this automation session. Physical paste acceptance is still required."
            return
        }
        throw "Unable to reactivate smoke target before hotkey."
    }

    [PasteTextSmokeKeyboard]::CtrlAltV()
    Start-Sleep -Milliseconds $OverlayDelayMs
    if (-not [PasteTextSmokeKeyboard]::IsOverlayVisible()) {
        [void][PasteTextSmokeKeyboard]::PostHotkeyToMessageWindow()
        Start-Sleep -Milliseconds $OverlayDelayMs
    }
    [PasteTextSmokeKeyboard]::Enter()
    [void][PasteTextSmokeKeyboard]::PostEnterToOverlay()
    Start-Sleep -Milliseconds $PasteDelayMs

    Start-Sleep -Milliseconds 500

    if (Test-Path $targetResultPath) {
        $capturedText = [System.IO.File]::ReadAllText($targetResultPath, [System.Text.Encoding]::UTF8)
    }

    $pasteVerified = $capturedText -eq $uniqueText

    $status = [ordered]@{
        exe = $ExePath
        paste_process_id = $pasteProcess.Id
        target_process_id = $targetProcess.Id
        unique_text = $uniqueText
        captured_text = $capturedText
        paste_verified = $pasteVerified
        paste_process_alive = -not $pasteProcess.HasExited
        target_process_alive = -not $targetProcess.HasExited
        target_result_path = $targetResultPath
    }

    $status | ConvertTo-Json -Depth 3

    if (-not $pasteVerified) {
        if ($AllowForegroundUnavailable) {
            Write-Warning "Smoke target content did not match in this automation session. Physical paste acceptance is still required."
            return
        }
        throw "Smoke target content did not match the selected clipboard history text."
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
