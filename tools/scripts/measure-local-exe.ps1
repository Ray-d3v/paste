param(
    [string]$RustExePath = "artifacts/local-gpui/PasteWinUI.exe",
    [string]$WinUIExePath = "artifacts/local-winui/publish/PasteWinUI.exe",
    [int]$SampleDelayMs = 1500,
    [string]$OutputPath = "artifacts/perf/local-exe-metrics.json"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

$windowProbeSource = @'
using System;
using System.Runtime.InteropServices;

public static class PasteMeasureWindowProbe
{
    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

    [DllImport("user32.dll", SetLastError = true)]
    static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

    public static IntPtr FindWindowByTitle(string title)
    {
        return FindWindow(null, title);
    }

    public static IntPtr FindMessageWindow()
    {
        IntPtr hwnd = FindWindowEx(new IntPtr(-3), IntPtr.Zero, "PasteGPUIMessageWindow", "Paste GPUI Message Window");
        return hwnd != IntPtr.Zero
            ? hwnd
            : FindWindowEx(IntPtr.Zero, IntPtr.Zero, "PasteGPUIMessageWindow", "Paste GPUI Message Window");
    }
}
'@

if (-not ("PasteMeasureWindowProbe" -as [type])) {
    Add-Type -TypeDefinition $windowProbeSource
}

function Resolve-RepoPath {
    param([Parameter(Mandatory = $true)][string]$Path)

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return $Path
    }

    return (Join-Path $repoRoot $Path)
}

function Measure-Exe {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][int]$DelayMs
    )

    $resolvedPath = Resolve-RepoPath $Path
    if (-not (Test-Path $resolvedPath)) {
        return [ordered]@{
            name = $Name
            path = $resolvedPath
            status = "missing"
        }
    }

    Get-Process -Name "PasteWinUI" -ErrorAction SilentlyContinue | Stop-Process -Force
    Start-Sleep -Milliseconds 250

    $start = Get-Date
    $process = Start-Process -FilePath $resolvedPath -PassThru
    $startupReadyMs = $null
    $readySignal = $null
    $startupDeadline = $start.AddMilliseconds($DelayMs)
    do {
        Start-Sleep -Milliseconds 50
        $process.Refresh()
        if ($process.HasExited) {
            throw "$Name exited during startup. ExitCode=$($process.ExitCode)"
        }

        $overlayWindow = [PasteMeasureWindowProbe]::FindWindowByTitle("Paste GPUI Overlay")
        $messageWindow = [PasteMeasureWindowProbe]::FindMessageWindow()
        if ($process.MainWindowHandle -ne 0 -or $overlayWindow -ne [IntPtr]::Zero -or $messageWindow -ne [IntPtr]::Zero) {
            $startupReadyMs = [math]::Round(((Get-Date) - $start).TotalMilliseconds, 0)
            $readySignal = if ($messageWindow -ne [IntPtr]::Zero) {
                "message_window"
            }
            elseif ($overlayWindow -ne [IntPtr]::Zero) {
                "overlay_window"
            }
            else {
                "main_window"
            }
            break
        }
    } while ((Get-Date) -lt $startupDeadline)

    $remainingDelay = $DelayMs - [int]([math]::Round(((Get-Date) - $start).TotalMilliseconds, 0))
    if ($remainingDelay -gt 0) {
        Start-Sleep -Milliseconds $remainingDelay
    }
    $process.Refresh()
    $sample = Get-Date

    $overlayWindowAtSample = [PasteMeasureWindowProbe]::FindWindowByTitle("Paste GPUI Overlay")
    $messageWindowAtSample = [PasteMeasureWindowProbe]::FindMessageWindow()
    $mainWindowAppeared = $process.MainWindowHandle -ne 0
    $startupSampleMs = [math]::Round(($sample - $start).TotalMilliseconds, 0)
    $workingSetMb = [math]::Round($process.WorkingSet64 / 1MB, 2)
    $peakWorkingSetMb = [math]::Round($process.PeakWorkingSet64 / 1MB, 2)
    $fileSizeMb = [math]::Round((Get-Item $resolvedPath).Length / 1MB, 2)
    $deploymentDir = Split-Path -Parent $resolvedPath
    $deploymentSizeBytes = (Get-ChildItem -LiteralPath $deploymentDir -Recurse -File |
        Measure-Object -Property Length -Sum).Sum
    $deploymentSizeMb = [math]::Round($deploymentSizeBytes / 1MB, 2)

    Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue

    [ordered]@{
        name = $Name
        path = $resolvedPath
        status = "measured"
        sample_delay_ms = $DelayMs
        startup_sample_ms = $startupSampleMs
        startup_ready_ms = $startupReadyMs
        ready_signal = $readySignal
        main_window_appeared = $mainWindowAppeared
        overlay_window_appeared = $overlayWindowAtSample -ne [IntPtr]::Zero
        message_window_appeared = $messageWindowAtSample -ne [IntPtr]::Zero
        working_set_mb = $workingSetMb
        peak_working_set_mb = $peakWorkingSetMb
        file_size_mb = $fileSizeMb
        deployment_size_mb = $deploymentSizeMb
    }
}

$results = @(
    (Measure-Exe -Name "Rust GPUI" -Path $RustExePath -DelayMs $SampleDelayMs),
    (Measure-Exe -Name "Legacy WinUI" -Path $WinUIExePath -DelayMs $SampleDelayMs)
)

$payload = [ordered]@{
    measured_at = (Get-Date).ToString("o")
    sample_delay_ms = $SampleDelayMs
    metrics = $results
}

$resolvedOutput = Resolve-RepoPath $OutputPath
$outputDir = Split-Path -Parent $resolvedOutput
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
$payload | ConvertTo-Json -Depth 5 | Set-Content -Path $resolvedOutput -Encoding UTF8

$results | ForEach-Object {
    if ($_.status -eq "missing") {
        Write-Host "$($_.name): missing at $($_.path)"
    }
    else {
        Write-Host "$($_.name): exe=$($_.file_size_mb)MB deploy=$($_.deployment_size_mb)MB ws=$($_.working_set_mb)MB peak=$($_.peak_working_set_mb)MB ready=$($_.startup_ready_ms)ms sample=$($_.startup_sample_ms)ms"
    }
}
Write-Host "Metrics written: $resolvedOutput"
