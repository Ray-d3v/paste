using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media.Animation;

namespace BottomOverlay;

public partial class MainWindow : Window
{
    private const double HorizontalMargin = 16;
    private const double VerticalMargin = 16;
    private const int CornerRadius = 36;
    private static readonly Duration SlideDuration = new(TimeSpan.FromMilliseconds(280));
    private const uint MonitorDefaultToNearest = 2;

    private bool _isOpen;
    private bool _isAnimating;
    private double _finalTop;
    private double _hiddenTop;

    public MainWindow()
    {
        InitializeComponent();
        SourceInitialized += (_, _) =>
        {
            ApplyRoundedWindowShape();
            ApplyStrongBlur();
        };
        SizeChanged += (_, _) =>
        {
            ApplyRoundedWindowShape();
            ApplyStrongBlur();
        };
    }

    public void Toggle()
    {
        if (_isOpen)
        {
            HideOverlayAnimated();
            return;
        }

        ShowOverlayAnimated();
    }

    private void ShowOverlayAnimated()
    {
        var workArea = GetCursorMonitorWorkArea();
        var workWidth = workArea.Right - workArea.Left;
        var workHeight = workArea.Bottom - workArea.Top;

        var usableWidth = Math.Max(300, workWidth - (HorizontalMargin * 2));
        var overlayHeight = Math.Max(200, (workHeight / 3) - (VerticalMargin * 2));

        Width = usableWidth;
        Height = overlayHeight;
        Left = workArea.Left + HorizontalMargin;
        Topmost = true;

        _finalTop = workArea.Bottom - overlayHeight - VerticalMargin;
        _hiddenTop = workArea.Bottom + VerticalMargin;

        BeginAnimation(TopProperty, null);
        Top = _hiddenTop;
        Show();
        ApplyRoundedWindowShape();
        ApplyStrongBlur();

        var animation = new DoubleAnimation
        {
            From = Top,
            To = _finalTop,
            Duration = SlideDuration,
            EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut }
        };

        _isAnimating = true;
        _isOpen = true;
        animation.Completed += (_, _) => _isAnimating = false;
        BeginAnimation(TopProperty, animation);
    }

    private void HideOverlayAnimated()
    {
        if (!IsVisible && !_isAnimating)
        {
            _isOpen = false;
            return;
        }

        var from = Top;
        var to = _hiddenTop;

        BeginAnimation(TopProperty, null);

        var animation = new DoubleAnimation
        {
            From = from,
            To = to,
            Duration = SlideDuration,
            EasingFunction = new CubicEase { EasingMode = EasingMode.EaseIn }
        };

        _isAnimating = true;
        _isOpen = false;
        animation.Completed += (_, _) =>
        {
            _isAnimating = false;
            Hide();
        };

        BeginAnimation(TopProperty, animation);
    }

    private static RECT GetCursorMonitorWorkArea()
    {
        if (!GetCursorPos(out var cursor))
        {
            return new RECT
            {
                Left = 0,
                Top = 0,
                Right = (int)SystemParameters.WorkArea.Width,
                Bottom = (int)SystemParameters.WorkArea.Height
            };
        }

        var monitor = MonitorFromPoint(cursor, MonitorDefaultToNearest);
        if (monitor == IntPtr.Zero)
        {
            return new RECT
            {
                Left = 0,
                Top = 0,
                Right = (int)SystemParameters.WorkArea.Width,
                Bottom = (int)SystemParameters.WorkArea.Height
            };
        }

        var info = new MONITORINFO { cbSize = Marshal.SizeOf<MONITORINFO>() };
        if (!GetMonitorInfo(monitor, ref info))
        {
            return new RECT
            {
                Left = 0,
                Top = 0,
                Right = (int)SystemParameters.WorkArea.Width,
                Bottom = (int)SystemParameters.WorkArea.Height
            };
        }

        return info.rcWork;
    }

    [DllImport("user32.dll")]
    private static extern bool GetCursorPos(out POINT lpPoint);

    [DllImport("user32.dll")]
    private static extern IntPtr MonitorFromPoint(POINT pt, uint dwFlags);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    private static extern bool GetMonitorInfo(IntPtr hMonitor, ref MONITORINFO lpmi);

    [DllImport("gdi32.dll")]
    private static extern IntPtr CreateRoundRectRgn(int left, int top, int right, int bottom, int widthEllipse, int heightEllipse);

    [DllImport("user32.dll")]
    private static extern int SetWindowRgn(IntPtr hWnd, IntPtr hRgn, bool bRedraw);

    [DllImport("gdi32.dll")]
    private static extern bool DeleteObject(IntPtr hObject);

    [DllImport("user32.dll")]
    private static extern int SetWindowCompositionAttribute(IntPtr hwnd, ref WINDOWCOMPOSITIONATTRIBDATA data);

    private void ApplyStrongBlur()
    {
        var hwnd = new WindowInteropHelper(this).Handle;
        if (hwnd == IntPtr.Zero)
        {
            OverlayBorder.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromArgb(48, 128, 128, 128));
            return;
        }

        var accentSize = Marshal.SizeOf<ACCENTPOLICY>();
        var accentPtr = Marshal.AllocHGlobal(accentSize);

        try
        {
            var data = new WINDOWCOMPOSITIONATTRIBDATA
            {
                Attribute = WINDOWCOMPOSITIONATTRIB.WCA_ACCENT_POLICY,
                Data = accentPtr,
                SizeOfData = accentSize
            };

            var policies = new[]
            {
                new ACCENTPOLICY
                {
                    AccentState = ACCENTSTATE.ACCENT_ENABLE_ACRYLICBLURBEHIND,
                    AccentFlags = 2,
                    // ABGR packed color: semi-light gray tint over acrylic blur.
                    GradientColor = unchecked((int)0x30B0B0B0),
                    AnimationId = 0
                },
                new ACCENTPOLICY
                {
                    AccentState = ACCENTSTATE.ACCENT_ENABLE_BLURBEHIND,
                    AccentFlags = 0,
                    GradientColor = 0,
                    AnimationId = 0
                }
            };

            foreach (var policy in policies)
            {
                Marshal.StructureToPtr(policy, accentPtr, fDeleteOld: false);
                var result = SetWindowCompositionAttribute(hwnd, ref data);
                if (result != 0)
                {
                    OverlayBorder.Background = System.Windows.Media.Brushes.Transparent;
                    return;
                }
            }

            // Final fallback: a translucent gray plate (no blur available).
            var disabled = new ACCENTPOLICY
            {
                AccentState = ACCENTSTATE.ACCENT_DISABLED,
                AccentFlags = 0,
                GradientColor = 0,
                AnimationId = 0
            };
            Marshal.StructureToPtr(disabled, accentPtr, fDeleteOld: false);
            SetWindowCompositionAttribute(hwnd, ref data);
            OverlayBorder.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromArgb(48, 128, 128, 128));
        }
        finally
        {
            Marshal.FreeHGlobal(accentPtr);
        }
    }

    private void ApplyRoundedWindowShape()
    {
        if (ActualWidth <= 0 || ActualHeight <= 0)
        {
            return;
        }

        var hwnd = new WindowInteropHelper(this).Handle;
        if (hwnd == IntPtr.Zero)
        {
            return;
        }

        var region = CreateRoundRectRgn(0, 0, (int)ActualWidth + 1, (int)ActualHeight + 1, CornerRadius * 2, CornerRadius * 2);
        if (region == IntPtr.Zero)
        {
            return;
        }

        // On success, OS owns the region handle. On failure we must release it.
        var result = SetWindowRgn(hwnd, region, true);
        if (result == 0)
        {
            DeleteObject(region);
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct POINT
    {
        public int X;
        public int Y;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
    private struct MONITORINFO
    {
        public int cbSize;
        public RECT rcMonitor;
        public RECT rcWork;
        public int dwFlags;
    }

    private enum WINDOWCOMPOSITIONATTRIB
    {
        WCA_ACCENT_POLICY = 19
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct WINDOWCOMPOSITIONATTRIBDATA
    {
        public WINDOWCOMPOSITIONATTRIB Attribute;
        public IntPtr Data;
        public int SizeOfData;
    }

    private enum ACCENTSTATE
    {
        ACCENT_DISABLED = 0,
        ACCENT_ENABLE_BLURBEHIND = 3,
        ACCENT_ENABLE_ACRYLICBLURBEHIND = 4,
        ACCENT_ENABLE_HOSTBACKDROP = 5
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct ACCENTPOLICY
    {
        public ACCENTSTATE AccentState;
        public int AccentFlags;
        public int GradientColor;
        public int AnimationId;
    }
}
