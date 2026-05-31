param(
    [string]$NamePattern = "Paste|PasteWinUI|GPUI",
    [int]$MaxItems = 200
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

$win32Source = @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

public static class PasteTrayWin32Probe
{
    public class WindowInfo
    {
        public IntPtr Hwnd;
        public string ClassName;
        public string Title;
        public int X;
        public int Y;
        public int Width;
        public int Height;
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

    [DllImport("user32.dll")]
    static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll")]
    static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    public static List<WindowInfo> EnumerateTrayRelatedWindows()
    {
        var results = new List<WindowInfo>();
        EnumWindows((hwnd, _) =>
        {
            AddIfRelated(hwnd, results);
            EnumChildWindows(hwnd, (child, __) =>
            {
                AddIfRelated(child, results);
                return true;
            }, IntPtr.Zero);
            return true;
        }, IntPtr.Zero);
        return results;
    }

    static void AddIfRelated(IntPtr hwnd, List<WindowInfo> results)
    {
        string className = ReadClassName(hwnd);
        string title = ReadTitle(hwnd);
        string haystack = (className + " " + title).ToLowerInvariant();
        if (!(haystack.Contains("tray") ||
              haystack.Contains("notify") ||
              haystack.Contains("notification") ||
              haystack.Contains("overflow") ||
              haystack.Contains("toolbarwindow32") ||
              haystack.Contains("shell_traywnd") ||
              haystack.Contains("paste")))
        {
            return;
        }

        RECT rect;
        GetWindowRect(hwnd, out rect);
        results.Add(new WindowInfo
        {
            Hwnd = hwnd,
            ClassName = className,
            Title = title,
            X = rect.Left,
            Y = rect.Top,
            Width = rect.Right - rect.Left,
            Height = rect.Bottom - rect.Top,
        });
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

if (-not ("PasteTrayWin32Probe" -as [type])) {
    Add-Type -TypeDefinition $win32Source
}

$root = [System.Windows.Automation.AutomationElement]::RootElement
$treeScope = [System.Windows.Automation.TreeScope]::Descendants
$trueCondition = [System.Windows.Automation.Condition]::TrueCondition
$elements = $root.FindAll($treeScope, $trueCondition)

$items = New-Object System.Collections.Generic.List[object]
for ($i = 0; $i -lt $elements.Count -and $items.Count -lt $MaxItems; $i++) {
    $element = $elements.Item($i)
    $name = $element.Current.Name
    $className = $element.Current.ClassName
    $automationId = $element.Current.AutomationId
    $controlType = $element.Current.ControlType.ProgrammaticName

    $haystack = "$name $className $automationId $controlType"
    if ($haystack -notmatch "Tray|Notify|Notification|Overflow|Paste|GPUI|Toolbar|Taskbar|System") {
        continue
    }

    $rect = $element.Current.BoundingRectangle
    $items.Add([pscustomobject]@{
        name = $name
        class_name = $className
        automation_id = $automationId
        control_type = $controlType
        x = [int]$rect.X
        y = [int]$rect.Y
        width = [int]$rect.Width
        height = [int]$rect.Height
        matches_target = $haystack -match $NamePattern
    })
}

$targetItems = @($items | Where-Object { $_.matches_target })
$win32Items = @(
    [PasteTrayWin32Probe]::EnumerateTrayRelatedWindows() | ForEach-Object {
        [pscustomobject]@{
            hwnd = ("0x{0:X}" -f $_.Hwnd.ToInt64())
            class_name = $_.ClassName
            title = $_.Title
            x = $_.X
            y = $_.Y
            width = $_.Width
            height = $_.Height
            matches_target = ("$($_.ClassName) $($_.Title)" -match $NamePattern)
        }
    }
)
$win32TargetItems = @($win32Items | Where-Object { $_.matches_target })
$summary = [ordered]@{
    name_pattern = $NamePattern
    uia_target_item_count = $targetItems.Count
    uia_target_items = $targetItems
    uia_related_item_count = $items.Count
    uia_related_items = @($items.ToArray())
    win32_target_item_count = $win32TargetItems.Count
    win32_target_items = $win32TargetItems
    win32_related_item_count = $win32Items.Count
    win32_related_items = $win32Items
}

$summary | ConvertTo-Json -Depth 5
