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

public static class PasteFileSmokeKeyboard
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

if (-not ("PasteFileSmokeKeyboard" -as [type])) {
    Add-Type -TypeDefinition $keyboardSource
}

function Normalize-PathForSmoke {
    param([Parameter(Mandatory = $true)][string]$Path)
    return ([System.IO.Path]::GetFullPath($Path)).TrimEnd('\').ToLowerInvariant()
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
            if ($process.MainWindowHandle -ne [IntPtr]::Zero -and [PasteFileSmokeKeyboard]::Activate($process.MainWindowHandle)) {
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

function Set-SmokeClipboardFiles {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$FilePaths
    )

    Add-Type -AssemblyName System.Windows.Forms
    $collection = New-Object System.Collections.Specialized.StringCollection
    foreach ($path in $FilePaths) {
        [void]$collection.Add($path)
    }

    [System.Windows.Forms.Clipboard]::Clear()
    [System.Windows.Forms.Clipboard]::SetFileDropList($collection)
}

function Start-SmokeFileTarget {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ResultPath,
        [int]$Retries = 16
    )

    $targetScript = Join-Path ([System.IO.Path]::GetTempPath()) "paste-gpui-file-target-$([Guid]::NewGuid().ToString("N")).ps1"
    $targetSource = @'
param(
    [Parameter(Mandatory = $true)]
    [string]$ResultPath
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$source = @"
using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

public sealed class PasteFileSmokeForm : Form
{
    private readonly PasteFileSmokeBox box;

    public PasteFileSmokeForm(string resultPath)
    {
        Text = "Paste GPUI File Smoke Target";
        Width = 520;
        Height = 240;
        StartPosition = FormStartPosition.CenterScreen;
        TopMost = false;

        box = new PasteFileSmokeBox(resultPath);
        box.Dock = DockStyle.Fill;
        Controls.Add(box);
        Shown += delegate
        {
            Activate();
            box.Focus();
        };
    }
}

public sealed class PasteFileSmokeBox : Control
{
    private const int WM_PASTE = 0x0302;
    private readonly string resultPath;

    public PasteFileSmokeBox(string resultPath)
    {
        this.resultPath = resultPath;
        SetStyle(ControlStyles.Selectable, true);
        TabStop = true;
        BackColor = System.Drawing.Color.White;
    }

    protected override void WndProc(ref Message m)
    {
        if (m.Msg == WM_PASTE)
        {
            CaptureClipboardFiles();
            return;
        }

        base.WndProc(ref m);
    }

    private void CaptureClipboardFiles()
    {
        if (!Clipboard.ContainsFileDropList())
        {
            File.WriteAllText(resultPath, "contains_file_drop_list=false", Encoding.UTF8);
            return;
        }

        string[] files = Clipboard.GetFileDropList().Cast<string>().ToArray();
        string value = "contains_file_drop_list=true;count=" + files.Length + ";paths=" + string.Join("|", files);
        File.WriteAllText(resultPath, value, Encoding.UTF8);
    }
}
"@

Add-Type -TypeDefinition $source -ReferencedAssemblies System.Windows.Forms,System.Drawing
[System.Windows.Forms.Application]::Run([PasteFileSmokeForm]::new($ResultPath))
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

function Find-HistoryEntryByFilePaths {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$FilePaths
    )

    $historyPath = Join-Path $env:LOCALAPPDATA "PasteGPUI\history.json"
    if (-not (Test-Path $historyPath)) {
        return $null
    }

    $expected = @($FilePaths | ForEach-Object { Normalize-PathForSmoke $_ })
    $document = Get-Content -LiteralPath $historyPath -Raw | ConvertFrom-Json
    foreach ($entry in @($document.entries)) {
        if ($entry.kind -ne "File") {
            continue
        }

        $actual = @($entry.file_paths | ForEach-Object { Normalize-PathForSmoke $_ })
        if ($actual.Count -ne $expected.Count) {
            continue
        }

        $allMatch = $true
        for ($index = 0; $index -lt $expected.Count; $index++) {
            if ($actual[$index] -ne $expected[$index]) {
                $allMatch = $false
                break
            }
        }

        if ($allMatch) {
            return $entry
        }
    }

    return $null
}

function Test-PastedFilesMatch {
    param(
        [string]$CapturedResult,
        [string[]]$ExpectedPaths
    )

    if ($CapturedResult -notlike "contains_file_drop_list=true;*") {
        return $false
    }

    $prefix = "contains_file_drop_list=true;count=$($ExpectedPaths.Count);paths="
    if (-not $CapturedResult.StartsWith($prefix, [System.StringComparison]::Ordinal)) {
        return $false
    }

    $capturedPaths = @($CapturedResult.Substring($prefix.Length).Split('|') | ForEach-Object { Normalize-PathForSmoke $_ })
    $expected = @($ExpectedPaths | ForEach-Object { Normalize-PathForSmoke $_ })
    if ($capturedPaths.Count -ne $expected.Count) {
        return $false
    }

    for ($index = 0; $index -lt $expected.Count; $index++) {
        if ($capturedPaths[$index] -ne $expected[$index]) {
            return $false
        }
    }

    return $true
}

$pasteProcess = $null
$targetProcess = $null
$targetResultPath = Join-Path ([System.IO.Path]::GetTempPath()) "paste-gpui-file-result-$([Guid]::NewGuid().ToString("N")).txt"
$smokeDir = Join-Path ([System.IO.Path]::GetTempPath()) "paste-gpui-file-source-$([Guid]::NewGuid().ToString("N"))"
$filePaths = @()
$capturedResult = $null
$matchingEntry = $null
$pasteVerified = $false
$historyVerified = $false

try {
    New-Item -ItemType Directory -Path $smokeDir -Force | Out-Null
    $filePaths = @(
        (Join-Path $smokeDir "paste-file-a.txt"),
        (Join-Path $smokeDir "paste-file-b.log")
    )
    Set-Content -LiteralPath $filePaths[0] -Value "paste file smoke a" -Encoding UTF8
    Set-Content -LiteralPath $filePaths[1] -Value "paste file smoke b" -Encoding UTF8

    Get-Process -Name "PasteWinUI" -ErrorAction SilentlyContinue | Stop-Process -Force
    Start-Sleep -Milliseconds 250

    $pasteProcess = Start-Process -FilePath $ExePath -PassThru
    Start-Sleep -Milliseconds $StartupDelayMs
    $pasteProcess.Refresh()

    if ($pasteProcess.HasExited) {
        throw "PasteWinUI exited during startup. ExitCode=$($pasteProcess.ExitCode)"
    }

    $targetProcess = Start-SmokeFileTarget -ResultPath $targetResultPath

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
                source_file_paths = $filePaths
            }
            $status | ConvertTo-Json -Depth 4
            Write-Warning "Unable to activate file smoke target in this automation session. Physical paste acceptance is still required."
            return
        }
        throw "Unable to activate file smoke target process $($targetProcess.Id)."
    }

    Set-SmokeClipboardFiles -FilePaths $filePaths
    Start-Sleep -Milliseconds $ClipboardDelayMs

    if (-not (Activate-WindowByProcessId -ProcessId $targetProcess.Id)) {
        $matchingEntry = Find-HistoryEntryByFilePaths -FilePaths $filePaths
        $historyVerified = $matchingEntry -ne $null
        if ($AllowForegroundUnavailable) {
            $status = [ordered]@{
                exe = $ExePath
                paste_process_id = $pasteProcess.Id
                target_process_id = $targetProcess.Id
                foreground_activation_available = $false
                stage = "reactivate-before-hotkey"
                source_file_paths = $filePaths
                matching_history_kind = if ($matchingEntry -ne $null) { $matchingEntry.kind } else { $null }
                matching_history_file_paths = if ($matchingEntry -ne $null) { @($matchingEntry.file_paths) } else { $null }
                history_verified = $historyVerified
                paste_process_alive = -not $pasteProcess.HasExited
                target_process_alive = -not $targetProcess.HasExited
                target_result_path = $targetResultPath
            }
            $status | ConvertTo-Json -Depth 4
            if (-not $historyVerified) {
                throw "History did not contain the expected File item."
            }
            Write-Warning "Unable to reactivate file smoke target in this automation session. Physical paste acceptance is still required."
            return
        }
        throw "Unable to reactivate file smoke target before hotkey."
    }

    [PasteFileSmokeKeyboard]::CtrlAltV()
    Start-Sleep -Milliseconds $OverlayDelayMs
    if (-not [PasteFileSmokeKeyboard]::IsOverlayVisible()) {
        [void][PasteFileSmokeKeyboard]::PostHotkeyToMessageWindow()
        Start-Sleep -Milliseconds $OverlayDelayMs
    }
    [PasteFileSmokeKeyboard]::Enter()
    [void][PasteFileSmokeKeyboard]::PostEnterToOverlay()
    Start-Sleep -Milliseconds $PasteDelayMs
    Start-Sleep -Milliseconds 500

    if (Test-Path $targetResultPath) {
        $capturedResult = [System.IO.File]::ReadAllText($targetResultPath, [System.Text.Encoding]::UTF8)
    }

    $matchingEntry = Find-HistoryEntryByFilePaths -FilePaths $filePaths
    $pasteVerified = Test-PastedFilesMatch -CapturedResult $capturedResult -ExpectedPaths $filePaths
    $historyVerified = $matchingEntry -ne $null

    $status = [ordered]@{
        exe = $ExePath
        paste_process_id = $pasteProcess.Id
        target_process_id = $targetProcess.Id
        source_file_paths = $filePaths
        captured_result = $capturedResult
        matching_history_kind = if ($matchingEntry -ne $null) { $matchingEntry.kind } else { $null }
        matching_history_file_paths = if ($matchingEntry -ne $null) { @($matchingEntry.file_paths) } else { $null }
        paste_verified = $pasteVerified
        history_verified = $historyVerified
        paste_process_alive = -not $pasteProcess.HasExited
        target_process_alive = -not $targetProcess.HasExited
        target_result_path = $targetResultPath
    }

    $status | ConvertTo-Json -Depth 4

    if (-not $pasteVerified) {
        if ($AllowForegroundUnavailable) {
            Write-Warning "File smoke target did not receive the expected CF_HDROP file list in this automation session. Physical paste acceptance is still required."
            return
        }
        throw "File smoke target did not receive the expected CF_HDROP file list."
    }

    if (-not $historyVerified) {
        throw "History did not contain the expected File item."
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
        Remove-Item -LiteralPath $smokeDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}
