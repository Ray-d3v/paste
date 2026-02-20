using Microsoft.UI.Windowing;
using Microsoft.UI;
using Microsoft.UI.Input;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Media.Imaging;
using Microsoft.UI.Xaml.Shapes;
using System.Diagnostics;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using System.Threading;
using WinRT.Interop;
using Windows.ApplicationModel.DataTransfer;
using Windows.Graphics;
using Windows.Graphics.Imaging;
using Windows.Storage;
using Windows.Storage.Streams;
using Windows.UI;

namespace PasteWinUI;

public sealed partial class MainWindow : Window
{
    private const int HotkeyId = 0x5000;
    private const uint ModAlt = 0x0001;
    private const uint ModControl = 0x0002;
    private const uint VkV = 0x56;
    private const int VkF = 0x46;
    private const int VkReturn = 0x0D;
    private const int VkLeft = 0x25;
    private const int VkRight = 0x27;
    private const int VkSpace = 0x20;
    private const int VkEscape = 0x1B;
    private const int WmHotkey = 0x0312;
    private const int WmApp = 0x8000;
    private const int WmTrayIcon = WmApp + 1;
    private const int WmKeydown = 0x0100;
    private const int WmKeyup = 0x0101;
    private const int WmSyskeydown = 0x0104;
    private const int WmSyskeyup = 0x0105;
    private const int WhKeyboardLl = 13;
    private const int WhMouseLl = 14;
    private const int WmClipboardUpdate = 0x031D;
    private const int WmPaste = 0x0302;
    private const int SwHide = 0;
    private const int SwRestore = 9;
    private const int WmLbuttondown = 0x0201;
    private const int WmLbuttonup = 0x0202;
    private const int WmRbuttondown = 0x0204;
    private const int WmMbuttondown = 0x0207;
    private const int WmXbuttondown = 0x020B;
    private const int WmGetIcon = 0x007F;
    private const int IconSmall = 0;
    private const int IconBig = 1;
    private const int IconSmall2 = 2;
    private const int GclHicon = -14;
    private const int GclHiconSm = -34;
    private const uint MonitorDefaultToNearest = 0x00000002;
    private const int VkControlKey = 0x11;
    private const int VkMenuKey = 0x12;
    private const int VkLcontrol = 0xA2;
    private const int VkRcontrol = 0xA3;
    private const int VkLmenu = 0xA4;
    private const int VkRmenu = 0xA5;
    private const int GwlExstyle = -20;
    private const uint GaRoot = 2;
    private const long WsExAppwindow = 0x00040000L;
    private const long WsExToolwindow = 0x00000080L;
    private const uint SwpNomove = 0x0002;
    private const uint SwpNosize = 0x0001;
    private const uint SwpNozorder = 0x0004;
    private const uint SwpFramechanged = 0x0020;
    private const uint InputKeyboard = 1;
    private const uint KeyeventfKeyup = 0x0002;
    private const uint KeyeventfUnicode = 0x0004;
    private const uint NimAdd = 0x00000000;
    private const uint NimDelete = 0x00000002;
    private const uint NimSetVersion = 0x00000004;
    private const uint NifMessage = 0x00000001;
    private const uint NifIcon = 0x00000002;
    private const uint NifTip = 0x00000004;
    private const uint NotifyIconVersion4 = 4;
    private const int IdiApplication = 0x7F00;

    private const int HorizontalMargin = 16;
    private const int VerticalMargin = 16;
    private const double OverlayInnerPadding = 20.0;
    private const int CornerRadius = 22;
    private const int MinWidth = 320;
    private const int MinHeight = 220;
    private const double AnimationDurationMs = 220;
    private const double WheelImpulsePxPerSec = 1400.0;
    private const double WheelInertiaDampingPerFrame = 0.93;
    private const double WheelStopVelocity = 10.0;
    private const double SelectionScrollSmoothingStrength = 18.0;
    private const double ScrollResetIdleSeconds = 30.0;
    private const double SearchAnimationDurationMs = 150.0;
    private const double CardContextPopupEdgePadding = 8.0;
    private const double CardContextPopupCardInsetX = 8.0;
    private const double CardContextPopupCardInsetY = 8.0;
    private const double VisibleCardCount = 5.5;
    private const double CardGapRatio = 0.012;
    private const double CardHeightRatio = 0.92;
    private const double DuplicateMergeWindowSeconds = 2.0;
    private const int HistorySaveDebounceMs = 450;
    private const int DwmwaBorderColor = 34;
    private const uint DwmColorNone = 0xFFFFFFFE;
    private static readonly HttpClient LinkPreviewHttpClient = CreateLinkPreviewHttpClient();
    private static readonly Regex HtmlTitleRegex = new(
        "<title[^>]*>(.*?)</title>",
        RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.Compiled);
    private static readonly Regex MetaTagRegex = new(
        "<meta\\b[^>]*>",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex MetaPropertyRegex = new(
        "(?:property|name)\\s*=\\s*[\"'](?<name>[^\"']+)[\"']",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex MetaContentRegex = new(
        "content\\s*=\\s*[\"'](?<content>[^\"']+)[\"']",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex LinkTagRegex = new(
        "<link\\b[^>]*>",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex LinkRelRegex = new(
        "rel\\s*=\\s*[\"'](?<rel>[^\"']+)[\"']",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex LinkHrefRegex = new(
        "href\\s*=\\s*[\"'](?<href>[^\"']+)[\"']",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex WhitespaceRegex = new(
        "\\s+",
        RegexOptions.Compiled);
    private readonly WndProcDelegate _wndProc;
    private readonly LowLevelKeyboardProc _keyboardProc;
    private readonly LowLevelMouseProc _mouseProc;

    private IntPtr _hwnd;
    private IntPtr _oldWndProc = IntPtr.Zero;
    private IntPtr _keyboardHook = IntPtr.Zero;
    private IntPtr _mouseHook = IntPtr.Zero;
    private IntPtr _trayIconHandle = IntPtr.Zero;
    private bool _trayIconAdded;
    private NOTIFYICONDATA _trayIconData;
    private AppWindow? _appWindow;

    private bool _isOpen;
    private bool _isAnimating;

    private int _x;
    private int _width;
    private int _height;
    private int _finalY;
    private int _hiddenY;
    private int _currentY;

    private DispatcherTimer? _animationTimer;
    private DateTimeOffset _animationStart;
    private int _fromY;
    private int _toY;
    private bool _hideWhenDone;
    private readonly List<Border> _cards = new();
    private readonly List<Grid> _cardHosts = new();
    private readonly List<Border> _selectionOutlines = new();
    private readonly List<(ClipboardEntry Entry, TextBlock Label)> _agoLabels = new();
    private readonly List<ClipboardEntry> _entries = new();
    private readonly List<ClipboardEntry> _filteredEntries = new();
    private int _selectedCardIndex = -1;
    private int _previewCardIndex = -1;
    private string _searchQuery = string.Empty;
    private uint _lastClipboardSequence;
    private DispatcherTimer? _scrollInertiaTimer;
    private DispatcherTimer? _agoRefreshTimer;
    private DateTimeOffset _scrollLastTick;
    private double _scrollVelocity;
    private DispatcherTimer? _selectionScrollTimer;
    private DateTimeOffset _selectionScrollLastTick;
    private double _selectionScrollTo;
    private DispatcherTimer? _searchAnimationTimer;
    private DateTimeOffset _searchAnimationStart;
    private bool _searchAnimationExpanding;
    private bool _searchAnimationClearQueryOnCollapse;
    private bool _searchAnimationRestoreFocusOnCollapse = true;
    private Timer? _hotkeyPollTimer;
    private bool _hotkeyChordDown;
    private bool _ctrlDown;
    private bool _altDown;
    private int _lastPolledHotkeyPressed;
    private int _hotkeyPollBusy;
    private IntPtr _lastExternalForegroundWindow = IntPtr.Zero;
    private IntPtr _overlayOpenTargetWindow = IntPtr.Zero;
    private DateTimeOffset _suppressClipboardUntilUtc = DateTimeOffset.MinValue;
    private bool _isPastingSelection;
    private DateTimeOffset _overlayClosedAtUtc = DateTimeOffset.MinValue;
    private bool _resetScrollOnNextOpen;
    private bool _pendingResetToLeftOnShow;
    private DateTimeOffset _suppressDeactivateCloseUntilUtc = DateTimeOffset.MinValue;
    private ClipboardEntry? _cardContextMenuEntry;
    private FrameworkElement? _cardContextPopupAnchor;
    private readonly string _historyFilePath = System.IO.Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "PasteWinUI",
        "clipboard-history.json");
    private CancellationTokenSource? _historySaveDebounceCts;

    public MainWindow()
    {
        InitializeComponent();

        _wndProc = WindowProc;
        _keyboardProc = KeyboardProc;
        _mouseProc = MouseProc;
        LoadHistoryFromDisk();
        BuildCards();
        OverlayPanel.SizeChanged += (_, _) =>
        {
            UpdateCardSize();
            if (CardPreviewPopup.IsOpen)
            {
                PositionPreviewBubble();
            }
            if (CardContextPopup.IsOpen && _cardContextPopupAnchor is not null)
            {
                PositionCardContextPopupForAnchor(_cardContextPopupAnchor);
            }
        };
        CardScroller.ViewChanged += (_, _) =>
        {
            if (CardPreviewPopup.IsOpen)
            {
                PositionPreviewBubble();
            }
        };
        Activated += OnWindowActivated;

        Closed += OnClosed;
    }

    public void InitializeOverlay()
    {
        _hwnd = WindowNative.GetWindowHandle(this);
        var windowId = Win32Interop.GetWindowIdFromWindow(_hwnd);
        _appWindow = AppWindow.GetFromWindowId(windowId);

        ConfigureWindowChrome();
        ConfigureDwmBorderless();
        ConfigureNoTaskbarIcon();
        ConfigureHotkeyHook();
        ConfigureHotkeyPollingFallback();
        ConfigureClipboardListener();

        UpdateBoundsForCursorMonitor();
        ApplyRoundedWindowRegion(_width, _height);
        _currentY = _hiddenY;
        _appWindow.MoveAndResize(new RectInt32(_x, _hiddenY, _width, _height));
        _appWindow.Show();
        ConfigureTrayIcon();
        ScheduleTrayIconRetry();
        _isOpen = false;
    }

    private void ToggleOverlay()
    {
        if (_isAnimating)
        {
            return;
        }

        if (_appWindow is null)
        {
            return;
        }

        if (!_trayIconAdded)
        {
            ConfigureTrayIcon();
        }

        UpdateBoundsForCursorMonitor();

        if (!_isOpen)
        {
            var foreground = GetForegroundWindow();
            var targetWindow = ResolveExternalTopLevelWindow(foreground);
            _overlayOpenTargetWindow = targetWindow;
            if (IsValidExternalWindow(targetWindow))
            {
                _lastExternalForegroundWindow = targetWindow;
            }

            var shouldApplyReset = false;
            var shouldResetByIdle = _overlayClosedAtUtc != DateTimeOffset.MinValue &&
                (DateTimeOffset.UtcNow - _overlayClosedAtUtc).TotalSeconds >= ScrollResetIdleSeconds;
            if (shouldResetByIdle || _resetScrollOnNextOpen)
            {
                RebuildCards();
                _resetScrollOnNextOpen = false;
                shouldApplyReset = true;
            }

            _appWindow.MoveAndResize(new RectInt32(_x, _hiddenY, _width, _height));
            _currentY = _hiddenY;
            _appWindow.Show();
            ActivateAndFocusInput();
            StartSlide(_hiddenY, _finalY, hideWhenDone: false);
            if (shouldApplyReset)
            {
                _pendingResetToLeftOnShow = true;
            }
            _isOpen = true;
            return;
        }

        CloseOverlay();
    }

    private void CloseOverlay()
    {
        if (!_isOpen && !_isAnimating)
        {
            return;
        }

        if (_isAnimating)
        {
            _animationTimer?.Stop();
            _isAnimating = false;
        }

        HideCardPreview();
        HideCardContextPopup();
        CollapseSearchBar(clearQuery: true, restoreFocus: false);
        StartSlide(_currentY, _hiddenY, hideWhenDone: true);
        _isOpen = false;
        _overlayClosedAtUtc = DateTimeOffset.UtcNow;
    }

    private void StartSlide(int fromY, int toY, bool hideWhenDone)
    {
        if (_appWindow is null)
        {
            return;
        }

        _fromY = fromY;
        _toY = toY;
        _hideWhenDone = hideWhenDone;
        _animationStart = DateTimeOffset.UtcNow;
        _isAnimating = true;

        _animationTimer ??= new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(16) };
        _animationTimer.Tick -= OnAnimationTick;
        _animationTimer.Tick += OnAnimationTick;
        _animationTimer.Start();
    }

    private void OnAnimationTick(object? sender, object e)
    {
        if (_appWindow is null)
        {
            return;
        }

        var elapsed = (DateTimeOffset.UtcNow - _animationStart).TotalMilliseconds;
        var t = Math.Clamp(elapsed / AnimationDurationMs, 0.0, 1.0);

        var eased = _hideWhenDone
            ? EaseInCubic(t)
            : EaseOutCubic(t);

        var y = (int)Math.Round(_fromY + ((_toY - _fromY) * eased));
        _currentY = y;
        _appWindow.Move(new PointInt32(_x, y));

        if (t < 1.0)
        {
            return;
        }

        _animationTimer?.Stop();
        _isAnimating = false;
        _currentY = _toY;

        if (_hideWhenDone)
        {
            // Keep the window alive off-screen so global hotkey remains registered.
            _appWindow.Move(new PointInt32(_x, _hiddenY));
        }
        else if (_isOpen)
        {
            // Re-assert focus so keyboard navigation is stable immediately after showing.
            ActivateAndFocusInput();
            if (_pendingResetToLeftOnShow)
            {
                _pendingResetToLeftOnShow = false;
                DispatcherQueue.TryEnqueue(() =>
                {
                    ResetCardScrollToLeft();
                    ResetSelectionToFirstCard();
                });
            }
        }
    }

    private void UpdateBoundsForCursorMonitor()
    {
        if (_appWindow is null)
        {
            return;
        }

        var monitorInfo = GetCursorMonitorInfo();
        var work = monitorInfo.rcWork;
        var monitor = monitorInfo.rcMonitor;
        var workWidth = work.Right - work.Left;
        var workHeight = work.Bottom - work.Top;

        _width = Math.Max(MinWidth, workWidth - (HorizontalMargin * 2));
        _height = Math.Max(MinHeight, (workHeight / 3) - (VerticalMargin * 2));

        _x = work.Left + HorizontalMargin;
        _finalY = work.Bottom - _height - VerticalMargin;
        _hiddenY = monitor.Bottom + VerticalMargin;

        _appWindow.MoveAndResize(new RectInt32(_x, _currentY, _width, _height));
        ApplyRoundedWindowRegion(_width, _height);
    }

    private void ConfigureWindowChrome()
    {
        if (_appWindow is null)
        {
            return;
        }

        if (_appWindow.Presenter is OverlappedPresenter presenter)
        {
            presenter.SetBorderAndTitleBar(false, false);
            presenter.IsResizable = false;
            presenter.IsMaximizable = false;
            presenter.IsMinimizable = false;
            presenter.IsAlwaysOnTop = true;
        }
    }

    private void ConfigureHotkeyHook()
    {
        _ = RegisterHotKey(_hwnd, HotkeyId, ModControl | ModAlt, VkV);
        InstallKeyboardFallbackHook();
        InstallMouseOutsideClickHook();
        _oldWndProc = SetWindowLongPtr(_hwnd, -4, Marshal.GetFunctionPointerForDelegate(_wndProc));
    }

    private void InstallKeyboardFallbackHook()
    {
        if (_keyboardHook != IntPtr.Zero)
        {
            return;
        }

        var module = GetModuleHandle(null);
        _keyboardHook = SetWindowsHookEx(WhKeyboardLl, _keyboardProc, module, 0);
    }

    private void InstallMouseOutsideClickHook()
    {
        if (_mouseHook != IntPtr.Zero)
        {
            return;
        }

        var module = GetModuleHandle(null);
        _mouseHook = SetWindowsHookExMouse(WhMouseLl, _mouseProc, module, 0);
    }

    private IntPtr KeyboardProc(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0)
        {
            var msg = wParam.ToInt32();
            var keyDown = msg == WmKeydown || msg == WmSyskeydown;
            var keyUp = msg == WmKeyup || msg == WmSyskeyup;

            if (keyDown || keyUp)
            {
                var keyData = Marshal.PtrToStructure<KBDLLHOOKSTRUCT>(lParam);
                var vk = (int)keyData.vkCode;
                var isDown = keyDown;
                var isUp = keyUp;

                if (isDown && _isOpen)
                {
                    if (vk == VkReturn)
                    {
                        DispatcherQueue.TryEnqueue(OnConfirmSelectionRequested);
                        return (IntPtr)1;
                    }

                    if (vk == VkLeft)
                    {
                        DispatcherQueue.TryEnqueue(() => MoveSelectedCard(-1));
                        return (IntPtr)1;
                    }

                    if (vk == VkRight)
                    {
                        DispatcherQueue.TryEnqueue(() => MoveSelectedCard(1));
                        return (IntPtr)1;
                    }

                    if (vk == VkEscape)
                    {
                        DispatcherQueue.TryEnqueue(() =>
                        {
                            if (IsSearchBarExpanded())
                            {
                                CollapseSearchBar(clearQuery: true);
                                return;
                            }

                            CloseOverlay();
                        });
                        return (IntPtr)1;
                    }

                    if (vk == VkF && _ctrlDown)
                    {
                        DispatcherQueue.TryEnqueue(ExpandSearchBar);
                        return (IntPtr)1;
                    }
                }

                if (vk == VkControlKey || vk == VkLcontrol || vk == VkRcontrol)
                {
                    if (isDown)
                    {
                        _ctrlDown = true;
                    }
                    if (isUp)
                    {
                        _ctrlDown = false;
                        _hotkeyChordDown = false;
                    }
                }
                else if (vk == VkMenuKey || vk == VkLmenu || vk == VkRmenu)
                {
                    if (isDown)
                    {
                        _altDown = true;
                    }
                    if (isUp)
                    {
                        _altDown = false;
                        _hotkeyChordDown = false;
                    }
                }
                else if (vk == VkV)
                {
                    if (isDown && _ctrlDown && _altDown && !_hotkeyChordDown)
                    {
                        _hotkeyChordDown = true;
                        if (!_isOpen)
                        {
                            var candidate = ResolveExternalTopLevelWindow(GetForegroundWindow());
                            if (IsValidExternalWindow(candidate))
                            {
                                _overlayOpenTargetWindow = candidate;
                                _lastExternalForegroundWindow = candidate;
                            }
                        }
                        DispatcherQueue.TryEnqueue(ToggleOverlay);
                        return (IntPtr)1;
                    }

                    if (isUp)
                    {
                        _hotkeyChordDown = false;
                    }
                }
            }
        }

        return CallNextHookEx(_keyboardHook, nCode, wParam, lParam);
    }

    private IntPtr MouseProc(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0 && (_isOpen || _isAnimating))
        {
            var msg = wParam.ToInt32();
            if (msg == WmLbuttondown || msg == WmRbuttondown || msg == WmMbuttondown || msg == WmXbuttondown)
            {
                var mouseData = Marshal.PtrToStructure<MSLLHOOKSTRUCT>(lParam);
                var isInsideOverlay = IsPointInsideOverlayWindow(mouseData.pt.X, mouseData.pt.Y);
                if (msg == WmRbuttondown && isInsideOverlay)
                {
                    var screenX = mouseData.pt.X;
                    var screenY = mouseData.pt.Y;
                    DispatcherQueue.TryEnqueue(() =>
                    {
                        if (_isOpen)
                        {
                            _ = TryShowCardContextPopupAtScreenPoint(screenX, screenY);
                        }
                    });
                    return (IntPtr)1;
                }

                if (!isInsideOverlay)
                {
                    DispatcherQueue.TryEnqueue(() =>
                    {
                        if (_isOpen || _isAnimating)
                        {
                            CloseOverlay();
                        }
                    });
                }
            }
        }

        return CallNextHookEx(_mouseHook, nCode, wParam, lParam);
    }

    private bool TryShowCardContextPopupAtScreenPoint(int screenX, int screenY)
    {
        if (_filteredEntries.Count == 0 || Root.ActualWidth <= 0 || Root.ActualHeight <= 0)
        {
            return false;
        }

        if (_hwnd == IntPtr.Zero || !GetWindowRect(_hwnd, out var windowRect))
        {
            return false;
        }

        var pointInRoot = new Windows.Foundation.Point(screenX - windowRect.Left, screenY - windowRect.Top);
        if (pointInRoot.X < 0 || pointInRoot.Y < 0 || pointInRoot.X > Root.ActualWidth || pointInRoot.Y > Root.ActualHeight)
        {
            return false;
        }

        var index = ResolveCardIndexAtRootPoint(pointInRoot);
        if (index < 0)
        {
            index = (_selectedCardIndex >= 0 && _selectedCardIndex < _filteredEntries.Count) ? _selectedCardIndex : 0;
        }

        if (index >= _cardHosts.Count)
        {
            return false;
        }

        ShowCardContextPopup(_filteredEntries[index], index, _cardHosts[index]);
        return true;
    }

    private int ResolveCardIndexAtRootPoint(Windows.Foundation.Point pointInRoot)
    {
        for (var i = 0; i < _cardHosts.Count && i < _filteredEntries.Count; i++)
        {
            var host = _cardHosts[i];
            if (host.ActualWidth <= 0 || host.ActualHeight <= 0)
            {
                continue;
            }

            var transform = host.TransformToVisual(Root);
            var topLeft = transform.TransformPoint(new Windows.Foundation.Point(0, 0));
            var right = topLeft.X + host.ActualWidth;
            var bottom = topLeft.Y + host.ActualHeight;
            if (pointInRoot.X >= topLeft.X && pointInRoot.X <= right &&
                pointInRoot.Y >= topLeft.Y && pointInRoot.Y <= bottom)
            {
                return i;
            }
        }

        return -1;
    }

    private bool IsPointInsideOverlayWindow(int screenX, int screenY)
    {
        if (_hwnd != IntPtr.Zero && GetWindowRect(_hwnd, out var windowRect))
        {
            return screenX >= windowRect.Left && screenX < windowRect.Right &&
                screenY >= windowRect.Top && screenY < windowRect.Bottom;
        }

        var left = _x;
        var top = _currentY;
        var right = left + _width;
        var bottom = top + _height;
        return screenX >= left && screenX < right && screenY >= top && screenY < bottom;
    }

    private void ConfigureClipboardListener()
    {
        if (_hwnd == IntPtr.Zero)
        {
            return;
        }

        AddClipboardFormatListener(_hwnd);
        _lastClipboardSequence = GetClipboardSequenceNumber();

        _agoRefreshTimer ??= new DispatcherTimer { Interval = TimeSpan.FromSeconds(20) };
        _agoRefreshTimer.Tick -= OnAgoRefreshTick;
        _agoRefreshTimer.Tick += OnAgoRefreshTick;
        if (!_agoRefreshTimer.IsEnabled)
        {
            _agoRefreshTimer.Start();
        }
    }

    private void ConfigureHotkeyPollingFallback()
    {
        if (_hotkeyPollTimer is null)
        {
            _hotkeyPollTimer = new Timer(_ => PollHotkeyOnBackground(), null, TimeSpan.Zero, TimeSpan.FromMilliseconds(16));
        }
        else
        {
            _hotkeyPollTimer.Change(TimeSpan.Zero, TimeSpan.FromMilliseconds(16));
        }
    }

    private void PollHotkeyOnBackground()
    {
        if (Interlocked.Exchange(ref _hotkeyPollBusy, 1) == 1)
        {
            return;
        }

        try
        {
        var ctrl = (GetAsyncKeyState(VkControlKey) & 0x8000) != 0
            || (GetAsyncKeyState(VkLcontrol) & 0x8000) != 0
            || (GetAsyncKeyState(VkRcontrol) & 0x8000) != 0;
        var alt = (GetAsyncKeyState(VkMenuKey) & 0x8000) != 0
            || (GetAsyncKeyState(VkLmenu) & 0x8000) != 0
            || (GetAsyncKeyState(VkRmenu) & 0x8000) != 0;
        var v = (GetAsyncKeyState((int)VkV) & 0x8000) != 0;

            var pressed = ctrl && alt && v ? 1 : 0;
            var previous = Volatile.Read(ref _lastPolledHotkeyPressed);

            if (pressed == 1 && previous == 0)
            {
                Volatile.Write(ref _lastPolledHotkeyPressed, 1);
                if (_hwnd != IntPtr.Zero)
                {
                    _ = PostMessage(_hwnd, WmHotkey, (IntPtr)HotkeyId, IntPtr.Zero);
                }
            }
            else if (pressed == 0 && previous == 1)
            {
                Volatile.Write(ref _lastPolledHotkeyPressed, 0);
            }
        }
        finally
        {
            Volatile.Write(ref _hotkeyPollBusy, 0);
        }
    }

    private void OnAgoRefreshTick(object? sender, object e)
    {
        var removed = PruneEntriesOlderThanOneMonth();
        if (removed > 0)
        {
            RebuildCards();
            ScheduleHistorySave();
        }

        RefreshAgoLabels();
    }

    private void RefreshAgoLabels()
    {
        foreach (var (entry, label) in _agoLabels)
        {
            label.Text = FormatAgo(entry.CopiedAtUtc);
        }
    }

    private async System.Threading.Tasks.Task HandleClipboardUpdatedAsync()
    {
        try
        {
            var seq = GetClipboardSequenceNumber();
            if (seq == _lastClipboardSequence)
            {
                return;
            }
            _lastClipboardSequence = seq;

            if (DateTimeOffset.UtcNow < _suppressClipboardUntilUtc)
            {
                return;
            }

            var package = Clipboard.GetContent();
            if (package is null)
            {
                return;
            }

            var entry = await BuildEntryFromClipboardAsync(package, DateTimeOffset.UtcNow);
            if (entry is null)
            {
                return;
            }

            if (_entries.Count > 0 && IsLikelyDuplicateClipboardEntry(_entries[0], entry))
            {
                _entries[0] = MergeClipboardEntries(_entries[0], entry);
            }
            else
            {
                _entries.Insert(0, entry);
                _resetScrollOnNextOpen = true;
                if (_entries.Count > 200)
                {
                    _entries.RemoveAt(_entries.Count - 1);
                }
            }

            _ = PruneEntriesOlderThanOneMonth();

            RebuildCards();
            RefreshAgoLabels();
            ScheduleHistorySave();
        }
        catch
        {
            // Clipboard can be transiently unavailable while owner app updates it.
        }
    }

    private void LoadHistoryFromDisk()
    {
        try
        {
            if (!File.Exists(_historyFilePath))
            {
                return;
            }

            var json = File.ReadAllText(_historyFilePath);
            if (string.IsNullOrWhiteSpace(json))
            {
                return;
            }

            var persisted = JsonSerializer.Deserialize<List<PersistedClipboardEntry>>(json);
            if (persisted is null || persisted.Count == 0)
            {
                return;
            }

            _entries.Clear();
            foreach (var item in persisted)
            {
                if (item is null || string.IsNullOrWhiteSpace(item.Kind))
                {
                    continue;
                }

                _entries.Add(new ClipboardEntry(
                    item.Kind,
                    item.SourceApp ?? "unknown",
                    item.Content ?? string.Empty,
                    item.CopiedAtUtc == default ? DateTimeOffset.UtcNow : item.CopiedAtUtc,
                    item.SourceExePath,
                    item.SourceIconPngBytes,
                    FromArgb(item.SourceHeaderColorArgb, Color.FromArgb(255, 44, 44, 44)),
                    item.ImagePngBytes,
                    item.ImageWidth,
                    item.ImageHeight,
                    item.LinkUrl,
                    item.LinkTitle,
                    item.LinkPreviewImageBytes,
                    item.LinkFaviconImageBytes,
                    item.LinkHost));
            }

            _entries.Sort((a, b) => b.CopiedAtUtc.CompareTo(a.CopiedAtUtc));
            if (PruneEntriesOlderThanOneMonth() > 0)
            {
                SaveHistoryToDiskNow();
            }
        }
        catch
        {
            // Corrupt or incompatible history should not break startup.
        }
    }

    private int PruneEntriesOlderThanOneMonth()
    {
        var cutoff = DateTimeOffset.UtcNow.AddMonths(-1);
        return _entries.RemoveAll(entry => entry.CopiedAtUtc < cutoff);
    }

    private void ScheduleHistorySave()
    {
        _historySaveDebounceCts?.Cancel();
        _historySaveDebounceCts?.Dispose();

        var cts = new CancellationTokenSource();
        _historySaveDebounceCts = cts;
        _ = SaveHistoryDebouncedAsync(cts.Token);
    }

    private async System.Threading.Tasks.Task SaveHistoryDebouncedAsync(CancellationToken token)
    {
        try
        {
            await System.Threading.Tasks.Task.Delay(HistorySaveDebounceMs, token);
        }
        catch (OperationCanceledException)
        {
            return;
        }

        SaveHistoryToDiskNow();
    }

    private void SaveHistoryToDiskNow()
    {
        try
        {
            var directory = System.IO.Path.GetDirectoryName(_historyFilePath);
            if (string.IsNullOrWhiteSpace(directory))
            {
                return;
            }

            Directory.CreateDirectory(directory);

            var persisted = new List<PersistedClipboardEntry>(_entries.Count);
            foreach (var entry in _entries)
            {
                persisted.Add(new PersistedClipboardEntry
                {
                    Kind = entry.Kind,
                    SourceApp = entry.SourceApp,
                    Content = entry.Content,
                    CopiedAtUtc = entry.CopiedAtUtc,
                    SourceExePath = entry.SourceExePath,
                    SourceIconPngBytes = entry.SourceIconPngBytes,
                    SourceHeaderColorArgb = ToArgb(entry.SourceHeaderColor),
                    ImagePngBytes = entry.ImagePngBytes,
                    ImageWidth = entry.ImageWidth,
                    ImageHeight = entry.ImageHeight,
                    LinkUrl = entry.LinkUrl,
                    LinkTitle = entry.LinkTitle,
                    LinkPreviewImageBytes = entry.LinkPreviewImageBytes,
                    LinkFaviconImageBytes = entry.LinkFaviconImageBytes,
                    LinkHost = entry.LinkHost
                });
            }

            var json = JsonSerializer.Serialize(persisted);
            File.WriteAllText(_historyFilePath, json);
        }
        catch
        {
            // Failing to persist should not break main workflow.
        }
    }

    private static uint ToArgb(Color color)
    {
        return ((uint)color.A << 24) | ((uint)color.R << 16) | ((uint)color.G << 8) | color.B;
    }

    private static Color FromArgb(uint argb, Color fallback)
    {
        if (argb == 0)
        {
            return fallback;
        }

        var a = (byte)((argb >> 24) & 0xFF);
        var r = (byte)((argb >> 16) & 0xFF);
        var g = (byte)((argb >> 8) & 0xFF);
        var b = (byte)(argb & 0xFF);
        return Color.FromArgb(a, r, g, b);
    }

    private async System.Threading.Tasks.Task<ClipboardEntry?> BuildEntryFromClipboardAsync(
        DataPackageView package,
        DateTimeOffset copiedAtUtc)
    {
        var (processName, sourceExePath, sourceIconPngBytes) = GetForegroundProcessInfo();
        var sourceApp = FormatSourceAppName(processName);
        var sourceHeaderColor = ResolveSourceHeaderColor(processName, sourceApp, sourceIconPngBytes);

        if (package.Contains(StandardDataFormats.StorageItems))
        {
            var items = await package.GetStorageItemsAsync();
            var names = new List<string>();
            foreach (var item in items)
            {
                names.Add(item.Name);
                if (names.Count >= 5)
                {
                    break;
                }
            }

            var content = names.Count > 0 ? string.Join(", ", names) : "[Files]";
            return new ClipboardEntry("File", sourceApp, content, copiedAtUtc, sourceExePath, sourceIconPngBytes, sourceHeaderColor);
        }

        if (package.Contains(StandardDataFormats.Bitmap))
        {
            var (imagePngBytes, imageWidth, imageHeight) = await TryReadClipboardBitmapAsync(package);
            return new ClipboardEntry(
                "Image",
                sourceApp,
                "[Image copied]",
                copiedAtUtc,
                sourceExePath,
                sourceIconPngBytes,
                sourceHeaderColor,
                imagePngBytes,
                imageWidth,
                imageHeight);
        }

        if (package.Contains(StandardDataFormats.WebLink))
        {
            var uri = await package.GetWebLinkAsync();
            var content = uri?.ToString() ?? "[Link]";
            var preview = uri is null
                ? (Title: (string?)null, PreviewImageBytes: (byte[]?)null, FaviconImageBytes: (byte[]?)null, Host: (string?)null)
                : await TryFetchLinkPreviewAsync(uri);
            return new ClipboardEntry(
                "Link",
                sourceApp,
                content,
                copiedAtUtc,
                sourceExePath,
                sourceIconPngBytes,
                sourceHeaderColor,
                linkUrl: content,
                linkTitle: preview.Title,
                linkPreviewImageBytes: preview.PreviewImageBytes,
                linkFaviconImageBytes: preview.FaviconImageBytes,
                linkHost: preview.Host);
        }

        if (package.Contains(StandardDataFormats.Text))
        {
            var text = (await package.GetTextAsync())?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(text))
            {
                return null;
            }

            var isLink = Uri.TryCreate(text, UriKind.Absolute, out var detectedUri);
            if (isLink && detectedUri is not null)
            {
                var preview = await TryFetchLinkPreviewAsync(detectedUri);
                return new ClipboardEntry(
                    "Link",
                    sourceApp,
                    text,
                    copiedAtUtc,
                    sourceExePath,
                    sourceIconPngBytes,
                    sourceHeaderColor,
                    linkUrl: text,
                    linkTitle: preview.Title,
                    linkPreviewImageBytes: preview.PreviewImageBytes,
                    linkFaviconImageBytes: preview.FaviconImageBytes,
                    linkHost: preview.Host);
            }

            return new ClipboardEntry("Text", sourceApp, text, copiedAtUtc, sourceExePath, sourceIconPngBytes, sourceHeaderColor);
        }

        return null;
    }

    private static async System.Threading.Tasks.Task<(byte[]? PngBytes, int Width, int Height)> TryReadClipboardBitmapAsync(
        DataPackageView package)
    {
        try
        {
            var bitmapReference = await package.GetBitmapAsync();
            if (bitmapReference is null)
            {
                return (null, 0, 0);
            }

            using var sourceStream = await bitmapReference.OpenReadAsync();
            var decoder = await BitmapDecoder.CreateAsync(sourceStream);
            using var softwareBitmap = await decoder.GetSoftwareBitmapAsync(
                BitmapPixelFormat.Bgra8,
                BitmapAlphaMode.Premultiplied);

            using var pngStream = new InMemoryRandomAccessStream();
            var encoder = await BitmapEncoder.CreateAsync(BitmapEncoder.PngEncoderId, pngStream);
            encoder.SetSoftwareBitmap(softwareBitmap);
            await encoder.FlushAsync();

            var pngBytes = await ReadStreamBytesAsync(pngStream);
            if (pngBytes is null || pngBytes.Length == 0)
            {
                return (null, 0, 0);
            }

            return (pngBytes, (int)decoder.PixelWidth, (int)decoder.PixelHeight);
        }
        catch
        {
            return (null, 0, 0);
        }
    }

    private static async System.Threading.Tasks.Task<byte[]?> ReadStreamBytesAsync(IRandomAccessStream stream)
    {
        try
        {
            stream.Seek(0);
            var size = checked((uint)stream.Size);
            if (size == 0)
            {
                return Array.Empty<byte>();
            }

            using var inputStream = stream.GetInputStreamAt(0);
            using var reader = new DataReader(inputStream);
            var loaded = await reader.LoadAsync(size);
            if (loaded < size)
            {
                return null;
            }

            var bytes = new byte[size];
            reader.ReadBytes(bytes);
            return bytes;
        }
        catch
        {
            return null;
        }
    }

    private static (string ProcessName, string? ExePath, byte[]? IconPngBytes) GetForegroundProcessInfo()
    {
        try
        {
            var hwnd = GetForegroundWindow();
            if (hwnd == IntPtr.Zero)
            {
                return ("unknown", null, null);
            }

            _ = GetWindowThreadProcessId(hwnd, out var pid);
            if (pid == 0)
            {
                return ("unknown", null, null);
            }

            using var process = Process.GetProcessById((int)pid);
            var iconPngBytes = TryGetWindowIconPngBytes(hwnd);
            string? exePath = null;
            try
            {
                exePath = process.MainModule?.FileName;
            }
            catch
            {
                // Access denied for some processes; keep fallback badge.
            }

            return (process.ProcessName.ToLowerInvariant(), exePath, iconPngBytes);
        }
        catch
        {
            return ("unknown", null, null);
        }
    }

    private static bool IsLikelyDuplicateClipboardEntry(ClipboardEntry existing, ClipboardEntry incoming)
    {
        var timeDiff = (incoming.CopiedAtUtc - existing.CopiedAtUtc).Duration();
        if (timeDiff > TimeSpan.FromSeconds(DuplicateMergeWindowSeconds))
        {
            return false;
        }

        if (!string.Equals(existing.Kind, incoming.Kind, StringComparison.Ordinal))
        {
            return false;
        }

        if (string.Equals(incoming.Kind, "Link", StringComparison.Ordinal))
        {
            var left = NormalizeLinkForComparison(existing.LinkUrl ?? existing.Content);
            var right = NormalizeLinkForComparison(incoming.LinkUrl ?? incoming.Content);
            return string.Equals(left, right, StringComparison.OrdinalIgnoreCase);
        }

        if (string.Equals(incoming.Kind, "Image", StringComparison.Ordinal))
        {
            return existing.ImageWidth == incoming.ImageWidth &&
                   existing.ImageHeight == incoming.ImageHeight &&
                   existing.ImagePngBytes is { Length: > 0 } &&
                   incoming.ImagePngBytes is { Length: > 0 };
        }

        return string.Equals(existing.Content, incoming.Content, StringComparison.Ordinal);
    }

    private static ClipboardEntry MergeClipboardEntries(ClipboardEntry existing, ClipboardEntry incoming)
    {
        var content = ChoosePreferredContent(incoming.Content, existing.Content);
        var sourceExePath = incoming.SourceExePath ?? existing.SourceExePath;
        var sourceIconPngBytes = incoming.SourceIconPngBytes ?? existing.SourceIconPngBytes;
        var sourceHeaderColor = incoming.SourceHeaderColor;
        var imagePngBytes = incoming.ImagePngBytes ?? existing.ImagePngBytes;
        var imageWidth = incoming.ImageWidth > 0 ? incoming.ImageWidth : existing.ImageWidth;
        var imageHeight = incoming.ImageHeight > 0 ? incoming.ImageHeight : existing.ImageHeight;
        var linkUrl = ChoosePreferredText(incoming.LinkUrl, existing.LinkUrl);
        var linkTitle = ChoosePreferredText(incoming.LinkTitle, existing.LinkTitle);
        var linkPreviewImageBytes = incoming.LinkPreviewImageBytes ?? existing.LinkPreviewImageBytes;
        var linkFaviconImageBytes = incoming.LinkFaviconImageBytes ?? existing.LinkFaviconImageBytes;
        var linkHost = ChoosePreferredText(incoming.LinkHost, existing.LinkHost);

        return new ClipboardEntry(
            incoming.Kind,
            incoming.SourceApp,
            content,
            incoming.CopiedAtUtc > existing.CopiedAtUtc ? incoming.CopiedAtUtc : existing.CopiedAtUtc,
            sourceExePath,
            sourceIconPngBytes,
            sourceHeaderColor,
            imagePngBytes,
            imageWidth,
            imageHeight,
            linkUrl,
            linkTitle,
            linkPreviewImageBytes,
            linkFaviconImageBytes,
            linkHost);
    }

    private static string ChoosePreferredContent(string preferred, string fallback)
    {
        if (string.IsNullOrWhiteSpace(preferred))
        {
            return fallback;
        }

        if (preferred is "[Link]" or "[Image copied]" && !string.IsNullOrWhiteSpace(fallback))
        {
            return fallback;
        }

        return preferred;
    }

    private static string? ChoosePreferredText(string? preferred, string? fallback)
    {
        return string.IsNullOrWhiteSpace(preferred) ? fallback : preferred;
    }

    private static string NormalizeLinkForComparison(string? rawLink)
    {
        if (string.IsNullOrWhiteSpace(rawLink))
        {
            return string.Empty;
        }

        var trimmed = rawLink.Trim();
        if (!Uri.TryCreate(trimmed, UriKind.Absolute, out var uri))
        {
            return trimmed;
        }

        var builder = new UriBuilder(uri)
        {
            Fragment = string.Empty
        };

        if (builder.Path.Length > 1)
        {
            builder.Path = builder.Path.TrimEnd('/');
        }

        var normalized = builder.Uri.AbsoluteUri;
        return normalized.Length > 1 ? normalized.TrimEnd('/') : normalized;
    }

    private static string FormatSourceAppName(string processName)
    {
        return processName switch
        {
            "code" => "vscode",
            "msedge" => "edge",
            _ => processName
        };
    }

    private static HttpClient CreateLinkPreviewHttpClient()
    {
        var handler = new HttpClientHandler
        {
            AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
        };

        var client = new HttpClient(handler)
        {
            Timeout = TimeSpan.FromSeconds(2.5)
        };
        client.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent", "PasteWinUI/1.0");
        return client;
    }

    private static async System.Threading.Tasks.Task<(string? Title, byte[]? PreviewImageBytes, byte[]? FaviconImageBytes, string? Host)> TryFetchLinkPreviewAsync(Uri uri)
    {
        if (!uri.IsAbsoluteUri)
        {
            return (null, null, null, null);
        }

        if (!string.Equals(uri.Scheme, Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase))
        {
            return (null, null, null, null);
        }

        try
        {
            using var request = new HttpRequestMessage(HttpMethod.Get, uri);
            using var response = await LinkPreviewHttpClient.SendAsync(request, HttpCompletionOption.ResponseHeadersRead);
            if (!response.IsSuccessStatusCode)
            {
                return (null, null, null, uri.Host);
            }

            var mediaType = response.Content.Headers.ContentType?.MediaType;
            if (!string.IsNullOrWhiteSpace(mediaType) &&
                !mediaType.Contains("html", StringComparison.OrdinalIgnoreCase))
            {
                return (null, null, null, uri.Host);
            }

            var html = await response.Content.ReadAsStringAsync();
            var title = ExtractHtmlTitle(html);

            byte[]? previewImageBytes = null;
            var previewImageUri = TryExtractPreviewImageUri(uri, html);
            if (previewImageUri is not null)
            {
                previewImageBytes = await TryDownloadLinkPreviewImageAsync(previewImageUri);
            }

            byte[]? faviconImageBytes = null;
            if (previewImageBytes is null)
            {
                var faviconUri = TryExtractFaviconUri(uri, html) ?? TryCreateDefaultFaviconUri(uri);
                if (faviconUri is not null)
                {
                    faviconImageBytes = await TryDownloadLinkPreviewImageAsync(faviconUri);
                }
            }

            return (title, previewImageBytes, faviconImageBytes, uri.Host);
        }
        catch
        {
            return (null, null, null, uri.Host);
        }
    }

    private static string? ExtractHtmlTitle(string html)
    {
        var match = HtmlTitleRegex.Match(html);
        if (!match.Success)
        {
            return null;
        }

        var decoded = WebUtility.HtmlDecode(match.Groups[1].Value);
        var collapsed = WhitespaceRegex.Replace(decoded, " ").Trim();
        return string.IsNullOrWhiteSpace(collapsed) ? null : collapsed;
    }

    private static Uri? TryExtractPreviewImageUri(Uri pageUri, string html)
    {
        var tags = MetaTagRegex.Matches(html);
        foreach (Match tag in tags)
        {
            var propertyMatch = MetaPropertyRegex.Match(tag.Value);
            if (!propertyMatch.Success)
            {
                continue;
            }

            var propertyName = propertyMatch.Groups["name"].Value.Trim().ToLowerInvariant();
            if (propertyName is not ("og:image" or "og:image:url" or "twitter:image"))
            {
                continue;
            }

            var contentMatch = MetaContentRegex.Match(tag.Value);
            if (!contentMatch.Success)
            {
                continue;
            }

            var raw = WebUtility.HtmlDecode(contentMatch.Groups["content"].Value).Trim();
            if (string.IsNullOrWhiteSpace(raw))
            {
                continue;
            }

            if (Uri.TryCreate(raw, UriKind.Absolute, out var absolute))
            {
                return absolute;
            }

            if (Uri.TryCreate(pageUri, raw, out var resolved))
            {
                return resolved;
            }
        }

        return null;
    }

    private static Uri? TryExtractFaviconUri(Uri pageUri, string html)
    {
        var tags = LinkTagRegex.Matches(html);
        foreach (Match tag in tags)
        {
            var relMatch = LinkRelRegex.Match(tag.Value);
            if (!relMatch.Success)
            {
                continue;
            }

            var relValue = relMatch.Groups["rel"].Value.ToLowerInvariant();
            if (!relValue.Contains("icon", StringComparison.Ordinal))
            {
                continue;
            }

            var hrefMatch = LinkHrefRegex.Match(tag.Value);
            if (!hrefMatch.Success)
            {
                continue;
            }

            var raw = WebUtility.HtmlDecode(hrefMatch.Groups["href"].Value).Trim();
            var resolved = TryResolveUri(pageUri, raw);
            if (resolved is not null)
            {
                return resolved;
            }
        }

        return null;
    }

    private static Uri? TryCreateDefaultFaviconUri(Uri pageUri)
    {
        if (string.IsNullOrWhiteSpace(pageUri.Host))
        {
            return null;
        }

        return new Uri($"{pageUri.Scheme}://{pageUri.Host}/favicon.ico");
    }

    private static Uri? TryResolveUri(Uri baseUri, string raw)
    {
        if (string.IsNullOrWhiteSpace(raw))
        {
            return null;
        }

        if (Uri.TryCreate(raw, UriKind.Absolute, out var absolute))
        {
            return absolute;
        }

        if (Uri.TryCreate(baseUri, raw, out var relative))
        {
            return relative;
        }

        return null;
    }

    private static async System.Threading.Tasks.Task<byte[]?> TryDownloadLinkPreviewImageAsync(Uri imageUri)
    {
        try
        {
            if (!imageUri.IsAbsoluteUri)
            {
                return null;
            }

            if (!string.Equals(imageUri.Scheme, Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(imageUri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            using var request = new HttpRequestMessage(HttpMethod.Get, imageUri);
            using var response = await LinkPreviewHttpClient.SendAsync(request, HttpCompletionOption.ResponseHeadersRead);
            if (!response.IsSuccessStatusCode)
            {
                return null;
            }

            var mediaType = response.Content.Headers.ContentType?.MediaType;
            if (!string.IsNullOrWhiteSpace(mediaType) &&
                !mediaType.StartsWith("image/", StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            if (response.Content.Headers.ContentLength is long contentLength &&
                contentLength > (6L * 1024L * 1024L))
            {
                return null;
            }

            var bytes = await response.Content.ReadAsByteArrayAsync();
            if (bytes.Length == 0 || bytes.Length > (6 * 1024 * 1024))
            {
                return null;
            }

            return bytes;
        }
        catch
        {
            return null;
        }
    }

    private static Color ResolveSourceHeaderColor(string processName, string sourceApp, byte[]? sourceIconPngBytes)
    {
        var iconColor = TryExtractRepresentativeColor(sourceIconPngBytes);
        if (iconColor.HasValue)
        {
            return ToHeaderBarColor(iconColor.Value);
        }

        return GetKnownSourceColor(processName, sourceApp);
    }

    private static Color GetKnownSourceColor(string processName, string sourceApp)
    {
        var key = (sourceApp ?? processName).ToLowerInvariant();
        return key switch
        {
            "chrome" => Color.FromArgb(255, 66, 133, 244),
            "msedge" => Color.FromArgb(255, 15, 163, 236),
            "edge" => Color.FromArgb(255, 15, 163, 236),
            "code" => Color.FromArgb(255, 0, 122, 204),
            "vscode" => Color.FromArgb(255, 0, 122, 204),
            "firefox" => Color.FromArgb(255, 255, 113, 57),
            "explorer" => Color.FromArgb(255, 240, 188, 66),
            "notion" => Color.FromArgb(255, 82, 82, 82),
            "slack" => Color.FromArgb(255, 74, 21, 75),
            "teams" => Color.FromArgb(255, 98, 100, 167),
            _ => Color.FromArgb(255, 44, 44, 44)
        };
    }

    private static Color? TryExtractRepresentativeColor(byte[]? pngBytes)
    {
        if (pngBytes is null || pngBytes.Length == 0)
        {
            return null;
        }

        try
        {
            using var stream = new MemoryStream(pngBytes, writable: false);
            using var bitmap = new System.Drawing.Bitmap(stream);

            if (bitmap.Width <= 0 || bitmap.Height <= 0)
            {
                return null;
            }

            var stepX = Math.Max(1, bitmap.Width / 24);
            var stepY = Math.Max(1, bitmap.Height / 24);

            double weightedR = 0;
            double weightedG = 0;
            double weightedB = 0;
            double weightedTotal = 0;

            for (var y = 0; y < bitmap.Height; y += stepY)
            {
                for (var x = 0; x < bitmap.Width; x += stepX)
                {
                    var pixel = bitmap.GetPixel(x, y);
                    if (pixel.A < 48)
                    {
                        continue;
                    }

                    var max = Math.Max(pixel.R, Math.Max(pixel.G, pixel.B));
                    var min = Math.Min(pixel.R, Math.Min(pixel.G, pixel.B));
                    var saturation = max == 0 ? 0.0 : (max - min) / (double)max;
                    var luminance = ((0.2126 * pixel.R) + (0.7152 * pixel.G) + (0.0722 * pixel.B)) / 255.0;

                    if (luminance < 0.10 || luminance > 0.95)
                    {
                        continue;
                    }

                    var weight = 0.35 + (saturation * 1.65);
                    weightedR += pixel.R * weight;
                    weightedG += pixel.G * weight;
                    weightedB += pixel.B * weight;
                    weightedTotal += weight;
                }
            }

            if (weightedTotal <= 0.0)
            {
                return null;
            }

            var r = (byte)Math.Clamp((int)Math.Round(weightedR / weightedTotal), 0, 255);
            var g = (byte)Math.Clamp((int)Math.Round(weightedG / weightedTotal), 0, 255);
            var b = (byte)Math.Clamp((int)Math.Round(weightedB / weightedTotal), 0, 255);
            return Color.FromArgb(255, r, g, b);
        }
        catch
        {
            return null;
        }
    }

    private static Color ToHeaderBarColor(Color color)
    {
        const double darkBase = 20.0;
        const double mix = 0.74;

        var r = (color.R * mix) + (darkBase * (1.0 - mix));
        var g = (color.G * mix) + (darkBase * (1.0 - mix));
        var b = (color.B * mix) + (darkBase * (1.0 - mix));

        var max = Math.Max(r, Math.Max(g, b));
        if (max > 0 && max < 78)
        {
            var boost = 78.0 / max;
            r *= boost;
            g *= boost;
            b *= boost;
        }

        return Color.FromArgb(
            255,
            (byte)Math.Clamp((int)Math.Round(r), 0, 255),
            (byte)Math.Clamp((int)Math.Round(g), 0, 255),
            (byte)Math.Clamp((int)Math.Round(b), 0, 255));
    }

    private static string FormatAgo(DateTimeOffset copiedAtUtc)
    {
        var span = DateTimeOffset.UtcNow - copiedAtUtc;
        if (span.TotalMinutes < 1)
        {
            return "just now";
        }
        if (span.TotalHours < 1)
        {
            return $"{Math.Max(1, (int)Math.Floor(span.TotalMinutes))} minutes ago";
        }
        if (span.TotalDays < 1)
        {
            return $"{Math.Max(1, (int)Math.Floor(span.TotalHours))} hours ago";
        }
        if (span.TotalDays < 2)
        {
            return "yesterday";
        }
        return $"{Math.Max(2, (int)Math.Floor(span.TotalDays))} days ago";
    }

    private void ActivateAndFocusInput()
    {
        Activate();
        _ = BringWindowToTop(_hwnd);
        SetForegroundWindow(_hwnd);
        _ = SetActiveWindow(_hwnd);
        _ = Root.Focus(FocusState.Programmatic);
        _ = OverlayPanel.Focus(FocusState.Programmatic);

        DispatcherQueue.TryEnqueue(() =>
        {
            _ = BringWindowToTop(_hwnd);
            SetForegroundWindow(_hwnd);
            _ = SetActiveWindow(_hwnd);
            if (IsSearchBarExpanded())
            {
                _ = SearchTextBox.Focus(FocusState.Programmatic);
                SearchTextBox.Select(SearchTextBox.Text.Length, 0);
            }
            else
            {
                _ = Root.Focus(FocusState.Programmatic);
            }
        });
    }

    private IntPtr WindowProc(IntPtr hwnd, uint msg, IntPtr wParam, IntPtr lParam)
    {
        if (msg == WmTrayIcon)
        {
            var trayMsg = unchecked((int)lParam.ToInt64()) & 0xFFFF;
            if (trayMsg == WmLbuttonup)
            {
                DispatcherQueue.TryEnqueue(ToggleOverlay);
            }
            return IntPtr.Zero;
        }

        if (msg == WmHotkey && wParam.ToInt32() == HotkeyId)
        {
            DispatcherQueue.TryEnqueue(ToggleOverlay);
            return IntPtr.Zero;
        }
        if (msg == WmKeydown || msg == WmSyskeydown)
        {
            if (!_isOpen)
            {
                return CallWindowProc(_oldWndProc, hwnd, msg, wParam, lParam);
            }

            var vk = wParam.ToInt32();
            if (vk == VkReturn)
            {
                DispatcherQueue.TryEnqueue(OnConfirmSelectionRequested);
                return IntPtr.Zero;
            }

            if (vk == VkLeft)
            {
                DispatcherQueue.TryEnqueue(() => MoveSelectedCard(-1));
                return IntPtr.Zero;
            }

            if (vk == VkRight)
            {
                DispatcherQueue.TryEnqueue(() => MoveSelectedCard(1));
                return IntPtr.Zero;
            }

            if (vk == VkEscape)
            {
                DispatcherQueue.TryEnqueue(() =>
                {
                    if (IsSearchBarExpanded())
                    {
                        CollapseSearchBar(clearQuery: true);
                        return;
                    }

                    CloseOverlay();
                });
                return IntPtr.Zero;
            }

            if (vk == VkSpace && !IsSearchInputActive())
            {
                DispatcherQueue.TryEnqueue(ToggleSelectedCardPreview);
                return IntPtr.Zero;
            }

            if (vk == VkF && IsControlPressed())
            {
                DispatcherQueue.TryEnqueue(ExpandSearchBar);
                return IntPtr.Zero;
            }
        }
        if (msg == WmClipboardUpdate)
        {
            DispatcherQueue.TryEnqueue(async () => await HandleClipboardUpdatedAsync());
            return IntPtr.Zero;
        }

        return CallWindowProc(_oldWndProc, hwnd, msg, wParam, lParam);
    }

    private void OnClosed(object sender, WindowEventArgs args)
    {
        RemoveTrayIcon();

        if (_hwnd != IntPtr.Zero)
        {
            UnregisterHotKey(_hwnd, HotkeyId);
            RemoveClipboardFormatListener(_hwnd);
        }

        if (_keyboardHook != IntPtr.Zero)
        {
            _ = UnhookWindowsHookEx(_keyboardHook);
            _keyboardHook = IntPtr.Zero;
        }

        if (_mouseHook != IntPtr.Zero)
        {
            _ = UnhookWindowsHookEx(_mouseHook);
            _mouseHook = IntPtr.Zero;
        }

        if (_hwnd != IntPtr.Zero && _oldWndProc != IntPtr.Zero)
        {
            _ = SetWindowLongPtr(_hwnd, -4, _oldWndProc);
            _oldWndProc = IntPtr.Zero;
        }

        _agoRefreshTimer?.Stop();
        _selectionScrollTimer?.Stop();
        _scrollInertiaTimer?.Stop();
        _hotkeyPollTimer?.Dispose();
        _hotkeyPollTimer = null;
        _historySaveDebounceCts?.Cancel();
        _historySaveDebounceCts?.Dispose();
        _historySaveDebounceCts = null;
        _searchAnimationTimer?.Stop();
        _searchAnimationTimer = null;
        SaveHistoryToDiskNow();
    }

    private void OnWindowActivated(object sender, WindowActivatedEventArgs args)
    {
        if (args.WindowActivationState == WindowActivationState.Deactivated)
        {
            if (DateTimeOffset.UtcNow < _suppressDeactivateCloseUntilUtc)
            {
                return;
            }

            if (CardContextPopup.IsOpen)
            {
                return;
            }

            CloseOverlay();
        }
    }

    private void Root_PointerPressed(object sender, PointerRoutedEventArgs e)
    {
        if (IsEventInsideOverlayPanel(e.OriginalSource))
        {
            return;
        }

        CloseOverlay();
    }

    private void OverlayPanel_PointerPressed(object sender, PointerRoutedEventArgs e)
    {
        var point = e.GetCurrentPoint(OverlayPanel);
        if (!point.Properties.IsLeftButtonPressed)
        {
            return;
        }

        HideCardContextPopup();
        ActivateAndFocusInput();
        e.Handled = true;
    }

    private void Root_KeyDown(object sender, KeyRoutedEventArgs e)
    {
        if (e.Key == Windows.System.VirtualKey.Space)
        {
            if (IsSearchInputActive())
            {
                return;
            }

            ToggleSelectedCardPreview();
            e.Handled = true;
            return;
        }

        if (e.Key == Windows.System.VirtualKey.Escape && CardContextPopup.IsOpen)
        {
            HideCardContextPopup();
            e.Handled = true;
            return;
        }

        if (e.Key == Windows.System.VirtualKey.F && IsControlPressed())
        {
            ExpandSearchBar();
            e.Handled = true;
            return;
        }

        if (e.Key == Windows.System.VirtualKey.Escape && IsSearchBarExpanded())
        {
            CollapseSearchBar(clearQuery: true);
            e.Handled = true;
            return;
        }

        if (e.Key == Windows.System.VirtualKey.Enter)
        {
            OnConfirmSelectionRequested();
            e.Handled = true;
            return;
        }

        if (e.Key == Windows.System.VirtualKey.Escape)
        {
            CloseOverlay();
            e.Handled = true;
            return;
        }

        if (e.Key == Windows.System.VirtualKey.Left)
        {
            MoveSelectedCard(-1);
            e.Handled = true;
            return;
        }

        if (e.Key == Windows.System.VirtualKey.Right)
        {
            MoveSelectedCard(1);
            e.Handled = true;
        }
    }

    private void OnConfirmSelectionRequested()
    {
        if (_isPastingSelection)
        {
            return;
        }

        HideCardPreview();
        _ = PasteSelectedEntryAsync();
    }

    private void ToggleSelectedCardPreview()
    {
        if (_selectedCardIndex < 0 || _selectedCardIndex >= _filteredEntries.Count)
        {
            return;
        }

        if (CardPreviewPopup.IsOpen && _previewCardIndex == _selectedCardIndex)
        {
            HideCardPreview();
            return;
        }

        var selectedEntry = _filteredEntries[_selectedCardIndex];
        CardPreviewTitleText.Text = BuildPreviewTitle(selectedEntry);
        CardPreviewBodyText.Text = BuildPreviewBody(selectedEntry);
        CardPreviewImage.Source = null;

        var previewImageBytes = ResolvePreviewImageBytes(selectedEntry);
        if (previewImageBytes is { Length: > 0 })
        {
            CardPreviewImage.Visibility = Visibility.Visible;
            _ = TrySetClipboardImagePreviewAsync(previewImageBytes, CardPreviewImage);
        }
        else
        {
            CardPreviewImage.Visibility = Visibility.Collapsed;
        }

        _previewCardIndex = _selectedCardIndex;
        CardPreviewPopup.IsOpen = true;
        PositionPreviewBubble();
        DispatcherQueue.TryEnqueue(PositionPreviewBubble);
    }

    private void HideCardPreview()
    {
        if (!CardPreviewPopup.IsOpen)
        {
            _previewCardIndex = -1;
            return;
        }

        CardPreviewPopup.IsOpen = false;
        _previewCardIndex = -1;
    }

    private void ShowCardContextPopup(ClipboardEntry entry, int cardIndex, FrameworkElement anchor)
    {
        SetSelectedCard(cardIndex);
        _cardContextMenuEntry = entry;
        _cardContextPopupAnchor = anchor;
        _suppressDeactivateCloseUntilUtc = DateTimeOffset.UtcNow.AddMilliseconds(500);

        CardContextPopup.IsOpen = false;
        PositionCardContextPopupForAnchor(anchor);
        CardContextPopup.IsOpen = true;
        DispatcherQueue.TryEnqueue(() =>
        {
            if (_cardContextPopupAnchor is not null)
            {
                PositionCardContextPopupForAnchor(_cardContextPopupAnchor);
            }
        });
    }

    private void PositionCardContextPopupForAnchor(FrameworkElement anchor)
    {
        if (Root.ActualWidth <= 0 || Root.ActualHeight <= 0 || anchor.ActualWidth <= 0 || anchor.ActualHeight <= 0)
        {
            return;
        }

        CardContextPopupPanel.Measure(new Windows.Foundation.Size(double.PositiveInfinity, double.PositiveInfinity));
        var popupWidth = Math.Max(120.0, CardContextPopupPanel.DesiredSize.Width);
        var popupHeight = Math.Max(40.0, CardContextPopupPanel.DesiredSize.Height);

        var transform = anchor.TransformToVisual(Root);
        var topLeft = transform.TransformPoint(new Windows.Foundation.Point(0, 0));

        var maxX = Math.Max(CardContextPopupEdgePadding, Root.ActualWidth - popupWidth - CardContextPopupEdgePadding);
        var maxY = Math.Max(CardContextPopupEdgePadding, Root.ActualHeight - popupHeight - CardContextPopupEdgePadding);
        var x = topLeft.X + anchor.ActualWidth - popupWidth - CardContextPopupCardInsetX;
        var y = topLeft.Y + CardContextPopupCardInsetY;
        x = Math.Clamp(x, CardContextPopupEdgePadding, maxX);
        y = Math.Clamp(y, CardContextPopupEdgePadding, maxY);

        CardContextPopup.HorizontalOffset = x;
        CardContextPopup.VerticalOffset = y;
    }

    private void HideCardContextPopup()
    {
        if (!CardContextPopup.IsOpen)
        {
            _cardContextMenuEntry = null;
            _cardContextPopupAnchor = null;
            return;
        }

        CardContextPopup.IsOpen = false;
        _cardContextMenuEntry = null;
        _cardContextPopupAnchor = null;
    }

    private void CardContextDeleteButton_Click(object sender, RoutedEventArgs e)
    {
        var target = _cardContextMenuEntry;
        HideCardContextPopup();
        if (target is null)
        {
            return;
        }

        DeleteEntry(target);
    }

    private void CardContextPopup_Closed(object sender, object e)
    {
        _cardContextMenuEntry = null;
        _cardContextPopupAnchor = null;
    }

    private void PositionPreviewBubble()
    {
        if (!CardPreviewPopup.IsOpen || _selectedCardIndex < 0 || _selectedCardIndex >= _cardHosts.Count)
        {
            return;
        }

        var host = _cardHosts[_selectedCardIndex];
        if (host.ActualWidth <= 0 || host.ActualHeight <= 0 || Root.ActualWidth <= 0 || Root.ActualHeight <= 0)
        {
            return;
        }

        CardPreviewBubble.Measure(new Windows.Foundation.Size(double.PositiveInfinity, double.PositiveInfinity));
        var bubbleWidth = Math.Max(220.0, CardPreviewBubble.DesiredSize.Width);
        var bubbleHeight = Math.Max(120.0, CardPreviewBubble.DesiredSize.Height);

        var transform = host.TransformToVisual(Root);
        var hostTopLeft = transform.TransformPoint(new Windows.Foundation.Point(0, 0));
        var anchorCenterX = hostTopLeft.X + (host.ActualWidth / 2.0);
        var targetX = anchorCenterX - (bubbleWidth / 2.0);
        var targetY = hostTopLeft.Y - bubbleHeight - 10.0;

        const double edgePadding = 8.0;
        var maxX = Math.Max(edgePadding, Root.ActualWidth - bubbleWidth - edgePadding);
        targetX = Math.Clamp(targetX, edgePadding, maxX);
        targetY = Math.Max(edgePadding, targetY);

        CardPreviewPopup.HorizontalOffset = targetX;
        CardPreviewPopup.VerticalOffset = targetY;
    }

    private static string BuildPreviewTitle(ClipboardEntry entry)
    {
        if (entry.Kind == "Link")
        {
            if (!string.IsNullOrWhiteSpace(entry.LinkTitle))
            {
                return entry.LinkTitle!;
            }

            if (!string.IsNullOrWhiteSpace(entry.LinkHost))
            {
                return entry.LinkHost!;
            }
        }

        if (entry.Kind == "Image")
        {
            if (entry.ImageWidth > 0 && entry.ImageHeight > 0)
            {
                return $"Image {entry.ImageWidth:N0}x{entry.ImageHeight:N0}";
            }

            return "Image";
        }

        return entry.Kind;
    }

    private static string BuildPreviewBody(ClipboardEntry entry)
    {
        if (entry.Kind == "Link")
        {
            var link = string.IsNullOrWhiteSpace(entry.LinkUrl) ? entry.Content : entry.LinkUrl!;
            if (!string.IsNullOrWhiteSpace(entry.Content) && !string.Equals(entry.Content, link, StringComparison.Ordinal))
            {
                return $"{link}\n\n{entry.Content}";
            }

            return link;
        }

        if (entry.Kind == "Image")
        {
            return $"Copied {FormatAgo(entry.CopiedAtUtc)}";
        }

        return entry.Content;
    }

    private static byte[]? ResolvePreviewImageBytes(ClipboardEntry entry)
    {
        if (entry.Kind == "Image")
        {
            return entry.ImagePngBytes;
        }

        if (entry.Kind == "Link")
        {
            if (entry.LinkPreviewImageBytes is { Length: > 0 })
            {
                return entry.LinkPreviewImageBytes;
            }

            if (entry.LinkFaviconImageBytes is { Length: > 0 })
            {
                return entry.LinkFaviconImageBytes;
            }
        }

        return null;
    }

    private async System.Threading.Tasks.Task PasteSelectedEntryAsync()
    {
        if (_isPastingSelection)
        {
            return;
        }

        if (_selectedCardIndex < 0 || _selectedCardIndex >= _filteredEntries.Count)
        {
            return;
        }

        _isPastingSelection = true;
        try
        {
            var selectedEntry = _filteredEntries[_selectedCardIndex];
            var clipboardApplied = await TryApplyEntryToClipboardAsync(selectedEntry);
            if (!clipboardApplied)
            {
                return;
            }

            var targetWindow = ResolvePasteTargetWindow();
            HideOverlayForPaste();

            await System.Threading.Tasks.Task.Delay(80);
            if (!IsValidExternalWindow(targetWindow))
            {
                targetWindow = ResolveExternalTopLevelWindow(GetForegroundWindow());
                if (!IsValidExternalWindow(targetWindow))
                {
                    _ = SendCtrlVKeystroke();
                    return;
                }
            }

            _ = BringWindowToTop(targetWindow);
            InjectEntryIntoTarget(targetWindow);
        }
        finally
        {
            _isPastingSelection = false;
        }
    }

    private void HideOverlayForPaste()
    {
        if (_isAnimating)
        {
            _animationTimer?.Stop();
            _isAnimating = false;
        }

        HideCardPreview();
        HideCardContextPopup();
        CollapseSearchBar(clearQuery: true, restoreFocus: false);

        _isOpen = false;
        _overlayClosedAtUtc = DateTimeOffset.UtcNow;
        _currentY = _hiddenY;
        _appWindow?.Move(new PointInt32(_x, _hiddenY));
        _appWindow?.Hide();
        if (_hwnd != IntPtr.Zero)
        {
            _ = ShowWindow(_hwnd, SwHide);
        }
    }

    private IntPtr ResolvePasteTargetWindow()
    {
        if (IsValidExternalWindow(_overlayOpenTargetWindow))
        {
            return _overlayOpenTargetWindow;
        }

        if (IsValidExternalWindow(_lastExternalForegroundWindow))
        {
            return _lastExternalForegroundWindow;
        }

        var foreground = GetForegroundWindow();
        var targetWindow = ResolveExternalTopLevelWindow(foreground);
        return IsValidExternalWindow(targetWindow) ? targetWindow : IntPtr.Zero;
    }

    private bool IsValidExternalWindow(IntPtr hwnd)
    {
        return hwnd != IntPtr.Zero && hwnd != _hwnd && IsWindow(hwnd);
    }

    private IntPtr ResolveExternalTopLevelWindow(IntPtr hwnd)
    {
        if (!IsValidExternalWindow(hwnd))
        {
            return IntPtr.Zero;
        }

        var root = GetAncestor(hwnd, GaRoot);
        if (IsValidExternalWindow(root))
        {
            return root;
        }

        return hwnd;
    }

    private async System.Threading.Tasks.Task<bool> TryApplyEntryToClipboardAsync(ClipboardEntry entry)
    {
        for (var attempt = 0; attempt < 5; attempt++)
        {
            try
            {
                var package = new DataPackage
                {
                    RequestedOperation = DataPackageOperation.Copy
                };

                if (string.Equals(entry.Kind, "Image", StringComparison.Ordinal) &&
                    entry.ImagePngBytes is { Length: > 0 })
                {
                    var streamRef = await CreateBitmapReferenceFromBytesAsync(entry.ImagePngBytes);
                    if (streamRef is null)
                    {
                        return false;
                    }

                    package.SetBitmap(streamRef);
                }
                else if (string.Equals(entry.Kind, "Link", StringComparison.Ordinal))
                {
                    var linkText = string.IsNullOrWhiteSpace(entry.LinkUrl) ? entry.Content : entry.LinkUrl!;
                    if (Uri.TryCreate(linkText, UriKind.Absolute, out var uri))
                    {
                        package.SetWebLink(uri);
                    }
                    package.SetText(linkText);
                }
                else
                {
                    package.SetText(entry.Content);
                }

                _suppressClipboardUntilUtc = DateTimeOffset.UtcNow.AddMilliseconds(900);
                Clipboard.SetContent(package);
                Clipboard.Flush();
                return true;
            }
            catch
            {
                if (attempt == 4)
                {
                    return false;
                }

                await System.Threading.Tasks.Task.Delay(25);
            }
        }

        return false;
    }

    private static async System.Threading.Tasks.Task<RandomAccessStreamReference?> CreateBitmapReferenceFromBytesAsync(byte[] pngBytes)
    {
        if (pngBytes.Length == 0)
        {
            return null;
        }

        try
        {
            var stream = new InMemoryRandomAccessStream();
            using var writer = new DataWriter(stream);
            writer.WriteBytes(pngBytes);
            await writer.StoreAsync();
            await writer.FlushAsync();
            writer.DetachStream();
            stream.Seek(0);
            return RandomAccessStreamReference.CreateFromStream(stream);
        }
        catch
        {
            return null;
        }
    }

    private bool SendCtrlVKeystroke()
    {
        var inputs = new INPUT[]
        {
            new INPUT { type = InputKeyboard, U = new InputUnion { ki = new KEYBDINPUT { wVk = (ushort)VkControlKey, dwFlags = 0 } } },
            new INPUT { type = InputKeyboard, U = new InputUnion { ki = new KEYBDINPUT { wVk = (ushort)VkV, dwFlags = 0 } } },
            new INPUT { type = InputKeyboard, U = new InputUnion { ki = new KEYBDINPUT { wVk = (ushort)VkV, dwFlags = KeyeventfKeyup } } },
            new INPUT { type = InputKeyboard, U = new InputUnion { ki = new KEYBDINPUT { wVk = (ushort)VkControlKey, dwFlags = KeyeventfKeyup } } }
        };

        var sent = SendInput((uint)inputs.Length, inputs, Marshal.SizeOf<INPUT>());
        return sent == inputs.Length;
    }

    private void InjectEntryIntoTarget(IntPtr targetWindow)
    {
        var resolvedTarget = ResolveExternalTopLevelWindow(targetWindow);
        if (TryInjectWithAttachedThreads(resolvedTarget))
        {
            return;
        }

        var activated = IsValidExternalWindow(resolvedTarget) && TryActivateExternalWindow(resolvedTarget);
        if (activated)
        {
            if (SendCtrlVKeystroke())
            {
                return;
            }

            if (SendShiftInsertKeystroke())
            {
                return;
            }

            _ = TryPostPasteMessageToTarget(resolvedTarget);
            return;
        }

        if (SendCtrlVKeystroke())
        {
            return;
        }

        if (SendShiftInsertKeystroke())
        {
            return;
        }

        _ = TryPostPasteMessageToTarget(resolvedTarget);
    }

    private bool TryInjectWithAttachedThreads(IntPtr targetWindow)
    {
        if (!IsValidExternalWindow(targetWindow))
        {
            return false;
        }

        var foreground = GetForegroundWindow();
        var currentThreadId = GetCurrentThreadId();
        var targetThreadId = GetWindowThreadProcessId(targetWindow, out _);
        var foregroundThreadId = foreground == IntPtr.Zero ? 0 : GetWindowThreadProcessId(foreground, out _);
        var attachedToTarget = false;
        var attachedToForeground = false;

        try
        {
            if (IsIconic(targetWindow))
            {
                _ = ShowWindow(targetWindow, SwRestore);
            }

            _ = AllowSetForegroundWindow(-1);

            if (targetThreadId != 0 && targetThreadId != currentThreadId)
            {
                attachedToTarget = AttachThreadInput(currentThreadId, targetThreadId, true);
            }

            if (foregroundThreadId != 0 &&
                foregroundThreadId != currentThreadId &&
                foregroundThreadId != targetThreadId)
            {
                attachedToForeground = AttachThreadInput(currentThreadId, foregroundThreadId, true);
            }

            _ = BringWindowToTop(targetWindow);
            _ = SetForegroundWindow(targetWindow);
            _ = SetActiveWindow(targetWindow);

            var guiThreadInfo = new GUITHREADINFO
            {
                cbSize = Marshal.SizeOf<GUITHREADINFO>()
            };
            if (targetThreadId != 0 && GetGUIThreadInfo(targetThreadId, ref guiThreadInfo))
            {
                if (IsWindow(guiThreadInfo.hwndFocus))
                {
                    _ = SetFocus(guiThreadInfo.hwndFocus);
                }
            }

            if (SendCtrlVKeystroke())
            {
                return true;
            }

            if (SendShiftInsertKeystroke())
            {
                return true;
            }

            _ = TryPostPasteMessageToTarget(targetWindow);
            return true;
        }
        catch
        {
            return false;
        }
        finally
        {
            if (attachedToForeground)
            {
                _ = AttachThreadInput(currentThreadId, foregroundThreadId, false);
            }

            if (attachedToTarget)
            {
                _ = AttachThreadInput(currentThreadId, targetThreadId, false);
            }
        }
    }

    private bool TryInjectTextEntryDirectly(ClipboardEntry entry)
    {
        if (string.Equals(entry.Kind, "Image", StringComparison.Ordinal))
        {
            return false;
        }

        var text = string.Equals(entry.Kind, "Link", StringComparison.Ordinal)
            ? (string.IsNullOrWhiteSpace(entry.LinkUrl) ? entry.Content : entry.LinkUrl!)
            : entry.Content;
        if (string.IsNullOrEmpty(text))
        {
            return false;
        }

        // Normalize LF-only lines for apps expecting CRLF on Enter-like text input.
        text = text.Replace("\r\n", "\n", StringComparison.Ordinal).Replace("\n", "\r\n", StringComparison.Ordinal);
        return SendUnicodeTextInput(text);
    }

    private bool SendUnicodeTextInput(string text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return false;
        }

        var inputs = new List<INPUT>(text.Length * 2);
        foreach (var ch in text)
        {
            inputs.Add(new INPUT
            {
                type = InputKeyboard,
                U = new InputUnion
                {
                    ki = new KEYBDINPUT
                    {
                        wVk = 0,
                        wScan = ch,
                        dwFlags = KeyeventfUnicode
                    }
                }
            });
            inputs.Add(new INPUT
            {
                type = InputKeyboard,
                U = new InputUnion
                {
                    ki = new KEYBDINPUT
                    {
                        wVk = 0,
                        wScan = ch,
                        dwFlags = KeyeventfUnicode | KeyeventfKeyup
                    }
                }
            });
        }

        var sent = SendInput((uint)inputs.Count, inputs.ToArray(), Marshal.SizeOf<INPUT>());
        return sent == inputs.Count;
    }

    private bool SendShiftInsertKeystroke()
    {
        var inputs = new INPUT[]
        {
            new INPUT { type = InputKeyboard, U = new InputUnion { ki = new KEYBDINPUT { wVk = 0x10, dwFlags = 0 } } },
            new INPUT { type = InputKeyboard, U = new InputUnion { ki = new KEYBDINPUT { wVk = 0x2D, dwFlags = 0 } } },
            new INPUT { type = InputKeyboard, U = new InputUnion { ki = new KEYBDINPUT { wVk = 0x2D, dwFlags = KeyeventfKeyup } } },
            new INPUT { type = InputKeyboard, U = new InputUnion { ki = new KEYBDINPUT { wVk = 0x10, dwFlags = KeyeventfKeyup } } }
        };

        var sent = SendInput((uint)inputs.Length, inputs, Marshal.SizeOf<INPUT>());
        return sent == inputs.Length;
    }

    private bool TryActivateExternalWindow(IntPtr targetWindow)
    {
        if (!IsValidExternalWindow(targetWindow))
        {
            return false;
        }

        var foreground = GetForegroundWindow();
        if (foreground == targetWindow)
        {
            return true;
        }

        var currentThreadId = GetCurrentThreadId();
        var targetThreadId = GetWindowThreadProcessId(targetWindow, out _);
        var foregroundThreadId = foreground == IntPtr.Zero ? 0 : GetWindowThreadProcessId(foreground, out _);
        var attachedToTarget = false;
        var attachedToForeground = false;

        try
        {
            if (IsIconic(targetWindow))
            {
                _ = ShowWindow(targetWindow, SwRestore);
            }

            _ = AllowSetForegroundWindow(-1);

            if (targetThreadId != 0 && targetThreadId != currentThreadId)
            {
                attachedToTarget = AttachThreadInput(currentThreadId, targetThreadId, true);
            }

            if (foregroundThreadId != 0 &&
                foregroundThreadId != currentThreadId &&
                foregroundThreadId != targetThreadId)
            {
                attachedToForeground = AttachThreadInput(currentThreadId, foregroundThreadId, true);
            }

            _ = BringWindowToTop(targetWindow);
            _ = SetForegroundWindow(targetWindow);
            _ = SetActiveWindow(targetWindow);
        }
        finally
        {
            if (attachedToForeground)
            {
                _ = AttachThreadInput(currentThreadId, foregroundThreadId, false);
            }

            if (attachedToTarget)
            {
                _ = AttachThreadInput(currentThreadId, targetThreadId, false);
            }
        }

        if (GetForegroundWindow() == targetWindow)
        {
            return true;
        }

        for (var attempt = 0; attempt < 6; attempt++)
        {
            _ = BringWindowToTop(targetWindow);
            _ = SetForegroundWindow(targetWindow);
            _ = SetActiveWindow(targetWindow);
            Thread.Sleep(20);
            if (GetForegroundWindow() == targetWindow)
            {
                return true;
            }
        }

        return false;
    }

    private bool TryPostPasteMessageToTarget(IntPtr targetWindow)
    {
        var resolvedTarget = ResolveExternalTopLevelWindow(targetWindow);
        if (!IsValidExternalWindow(resolvedTarget))
        {
            return false;
        }

        var targetThreadId = GetWindowThreadProcessId(resolvedTarget, out _);
        var guiThreadInfo = new GUITHREADINFO
        {
            cbSize = Marshal.SizeOf<GUITHREADINFO>()
        };

        IntPtr destination = resolvedTarget;
        if (targetThreadId != 0 && GetGUIThreadInfo(targetThreadId, ref guiThreadInfo))
        {
            if (IsWindow(guiThreadInfo.hwndFocus))
            {
                destination = guiThreadInfo.hwndFocus;
            }
        }

        return PostMessage(destination, WmPaste, IntPtr.Zero, IntPtr.Zero);
    }

    private void Root_PointerWheelChanged(object sender, PointerRoutedEventArgs e)
    {
        if (!_isOpen)
        {
            return;
        }

        var delta = e.GetCurrentPoint(Root).Properties.MouseWheelDelta;
        if (delta == 0)
        {
            return;
        }

        _scrollVelocity += (delta / 120.0) * WheelImpulsePxPerSec;
        EnsureScrollInertiaTimer();
        e.Handled = true;
    }

    private void EnsureScrollInertiaTimer()
    {
        _scrollInertiaTimer ??= new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(16) };
        _scrollInertiaTimer.Tick -= OnScrollInertiaTick;
        _scrollInertiaTimer.Tick += OnScrollInertiaTick;

        _scrollLastTick = DateTimeOffset.UtcNow;
        if (!_scrollInertiaTimer.IsEnabled)
        {
            _scrollInertiaTimer.Start();
        }
    }

    private void OnScrollInertiaTick(object? sender, object e)
    {
        if (!_isOpen)
        {
            _scrollInertiaTimer?.Stop();
            _scrollVelocity = 0;
            return;
        }

        var now = DateTimeOffset.UtcNow;
        var dt = Math.Max(0.001, (now - _scrollLastTick).TotalSeconds);
        _scrollLastTick = now;

        if (Math.Abs(_scrollVelocity) < WheelStopVelocity)
        {
            _scrollVelocity = 0;
            _scrollInertiaTimer?.Stop();
            return;
        }

        var deltaOffset = -_scrollVelocity * dt;
        var target = CardScroller.HorizontalOffset + deltaOffset;
        var clamped = Math.Clamp(target, 0, CardScroller.ScrollableWidth);
        CardScroller.ChangeView(clamped, null, null, true);

        // Exponential decay tuned for wheel-like momentum.
        var frameScale = Math.Pow(WheelInertiaDampingPerFrame, dt * 60.0);
        _scrollVelocity *= frameScale;

        // If we hit an edge, bleed speed faster.
        if (Math.Abs(clamped - target) > 0.001)
        {
            _scrollVelocity *= 0.6;
        }
    }

    private bool IsEventInsideOverlayPanel(object? source)
    {
        var current = source as DependencyObject;
        while (current is not null)
        {
            if (ReferenceEquals(current, OverlayPanel))
            {
                return true;
            }

            current = Microsoft.UI.Xaml.Media.VisualTreeHelper.GetParent(current);
        }

        return false;
    }

    private static double EaseOutCubic(double t)
    {
        var inv = 1.0 - t;
        return 1.0 - (inv * inv * inv);
    }

    private static double EaseInCubic(double t)
    {
        return t * t * t;
    }

    private static MONITORINFO GetCursorMonitorInfo()
    {
        static MONITORINFO CreateFallback()
        {
            var fallback = new MONITORINFO { cbSize = Marshal.SizeOf<MONITORINFO>() };
            fallback.rcMonitor = new RECT { Left = 0, Top = 0, Right = 1920, Bottom = 1080 };
            fallback.rcWork = fallback.rcMonitor;
            return fallback;
        }

        if (!GetCursorPos(out var cursor))
        {
            return CreateFallback();
        }

        var monitor = MonitorFromPoint(cursor, MonitorDefaultToNearest);
        if (monitor == IntPtr.Zero)
        {
            return CreateFallback();
        }

        var info = new MONITORINFO { cbSize = Marshal.SizeOf<MONITORINFO>() };
        if (!GetMonitorInfo(monitor, ref info))
        {
            return CreateFallback();
        }

        return info;
    }

    private delegate IntPtr WndProcDelegate(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);
    private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
    private delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hmod, uint dwThreadId);

    [DllImport("user32.dll", EntryPoint = "SetWindowsHookExW", SetLastError = true)]
    private static extern IntPtr SetWindowsHookExMouse(int idHook, LowLevelMouseProc lpfn, IntPtr hmod, uint dwThreadId);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll")]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern IntPtr GetModuleHandle(string? lpModuleName);

    [DllImport("kernel32.dll")]
    private static extern uint GetCurrentThreadId();

    [DllImport("user32.dll")]
    private static extern short GetAsyncKeyState(int vKey);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool AddClipboardFormatListener(IntPtr hwnd);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool RemoveClipboardFormatListener(IntPtr hwnd);

    [DllImport("user32.dll")]
    private static extern uint GetClipboardSequenceNumber();

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    private static extern bool IsWindow(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

    [DllImport("user32.dll")]
    private static extern IntPtr GetAncestor(IntPtr hWnd, uint gaFlags);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    [DllImport("user32.dll")]
    private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern IntPtr CopyIcon(IntPtr hIcon);

    [DllImport("user32.dll")]
    private static extern bool DestroyIcon(IntPtr hIcon);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr LoadIcon(IntPtr hInstance, IntPtr lpIconName);

    [DllImport("user32.dll", EntryPoint = "GetClassLongPtrW", SetLastError = true)]
    private static extern IntPtr GetClassLongPtr64(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll", EntryPoint = "GetClassLongW", SetLastError = true)]
    private static extern uint GetClassLong32(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    private static extern bool IsIconic(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool AllowSetForegroundWindow(int dwProcessId);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

    [DllImport("user32.dll")]
    private static extern bool BringWindowToTop(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern IntPtr SetActiveWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern IntPtr SetFocus(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool GetGUIThreadInfo(uint idThread, ref GUITHREADINFO lpgui);

    [DllImport("user32.dll", EntryPoint = "SetWindowLongPtrW")]
    private static extern IntPtr SetWindowLongPtr(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

    [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW")]
    private static extern IntPtr GetWindowLongPtr(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll")]
    private static extern bool PostMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern IntPtr CallWindowProc(IntPtr lpPrevWndFunc, IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, uint uFlags);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    private static extern bool GetCursorPos(out POINT lpPoint);

    [DllImport("user32.dll")]
    private static extern IntPtr MonitorFromPoint(POINT pt, uint dwFlags);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    private static extern bool GetMonitorInfo(IntPtr hMonitor, ref MONITORINFO lpmi);

    [DllImport("dwmapi.dll")]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int dwAttribute, ref uint pvAttribute, int cbAttribute);

    [DllImport("gdi32.dll")]
    private static extern IntPtr CreateRoundRectRgn(int left, int top, int right, int bottom, int widthEllipse, int heightEllipse);

    [DllImport("user32.dll")]
    private static extern int SetWindowRgn(IntPtr hWnd, IntPtr hRgn, bool bRedraw);

    [DllImport("gdi32.dll")]
    private static extern bool DeleteObject(IntPtr hObject);

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern bool Shell_NotifyIcon(uint dwMessage, ref NOTIFYICONDATA lpData);

    private void ApplyRoundedWindowRegion(int width, int height)
    {
        if (_hwnd == IntPtr.Zero || width <= 0 || height <= 0)
        {
            return;
        }

        var region = CreateRoundRectRgn(0, 0, width + 1, height + 1, CornerRadius * 2, CornerRadius * 2);
        if (region == IntPtr.Zero)
        {
            return;
        }

        var result = SetWindowRgn(_hwnd, region, true);
        if (result == 0)
        {
            DeleteObject(region);
        }
    }

    private void ConfigureDwmBorderless()
    {
        if (_hwnd == IntPtr.Zero)
        {
            return;
        }

        var none = DwmColorNone;
        _ = DwmSetWindowAttribute(_hwnd, DwmwaBorderColor, ref none, sizeof(uint));
    }

    private void ConfigureNoTaskbarIcon()
    {
        if (_hwnd == IntPtr.Zero)
        {
            return;
        }

        var exStylePtr = GetWindowLongPtr(_hwnd, GwlExstyle);
        var exStyle = exStylePtr.ToInt64();
        exStyle &= ~WsExAppwindow;
        exStyle |= WsExToolwindow;

        _ = SetWindowLongPtr(_hwnd, GwlExstyle, new IntPtr(exStyle));
        _ = SetWindowPos(_hwnd, IntPtr.Zero, 0, 0, 0, 0, SwpNomove | SwpNosize | SwpNozorder | SwpFramechanged);
    }

    private void ConfigureTrayIcon()
    {
        if (_hwnd == IntPtr.Zero || _trayIconAdded)
        {
            return;
        }

        var trayIconHandle = TryCreateTrayIconHandle();
        if (trayIconHandle == IntPtr.Zero)
        {
            return;
        }

        var data = new NOTIFYICONDATA
        {
            cbSize = (uint)Marshal.SizeOf<NOTIFYICONDATA>(),
            hWnd = _hwnd,
            uID = 1,
            uFlags = NifMessage | NifIcon | NifTip,
            uCallbackMessage = WmTrayIcon,
            hIcon = trayIconHandle,
            szTip = "PasteWinUI",
            szInfo = string.Empty,
            szInfoTitle = string.Empty
        };

        if (!Shell_NotifyIcon(NimAdd, ref data))
        {
            // Fallback for shell versions that expect older NOTIFYICONDATA size.
            data.cbSize = GetNotifyIconLegacySize();
            if (!Shell_NotifyIcon(NimAdd, ref data))
            {
                _ = DestroyIcon(trayIconHandle);
                return;
            }
        }

        data.uVersion = NotifyIconVersion4;
        _ = Shell_NotifyIcon(NimSetVersion, ref data);

        _trayIconHandle = trayIconHandle;
        _trayIconData = data;
        _trayIconAdded = true;
    }

    private void ScheduleTrayIconRetry()
    {
        _ = System.Threading.Tasks.Task.Run(async () =>
        {
            for (var i = 0; i < 3; i++)
            {
                await System.Threading.Tasks.Task.Delay(1200 * (i + 1));
                DispatcherQueue.TryEnqueue(() =>
                {
                    if (!_trayIconAdded && _hwnd != IntPtr.Zero)
                    {
                        ConfigureTrayIcon();
                    }
                });

                if (_trayIconAdded)
                {
                    break;
                }
            }
        });
    }

    private static uint GetNotifyIconLegacySize()
    {
        return (uint)Marshal.OffsetOf<NOTIFYICONDATA>(nameof(NOTIFYICONDATA.guidItem));
    }

    private IntPtr TryCreateTrayIconHandle()
    {
        try
        {
            var hIcon = SendMessage(_hwnd, WmGetIcon, (IntPtr)IconSmall2, IntPtr.Zero);
            if (hIcon == IntPtr.Zero)
            {
                hIcon = SendMessage(_hwnd, WmGetIcon, (IntPtr)IconSmall, IntPtr.Zero);
            }
            if (hIcon == IntPtr.Zero)
            {
                hIcon = SendMessage(_hwnd, WmGetIcon, (IntPtr)IconBig, IntPtr.Zero);
            }
            if (hIcon == IntPtr.Zero)
            {
                hIcon = GetClassLongPtr(_hwnd, GclHiconSm);
            }
            if (hIcon == IntPtr.Zero)
            {
                hIcon = GetClassLongPtr(_hwnd, GclHicon);
            }
            if (hIcon == IntPtr.Zero)
            {
                hIcon = System.Drawing.SystemIcons.Application.Handle;
            }
            if (hIcon == IntPtr.Zero)
            {
                hIcon = LoadIcon(IntPtr.Zero, (IntPtr)IdiApplication);
            }

            return hIcon == IntPtr.Zero ? IntPtr.Zero : CopyIcon(hIcon);
        }
        catch
        {
            return IntPtr.Zero;
        }
    }

    private void RemoveTrayIcon()
    {
        if (_trayIconAdded)
        {
            var data = _trayIconData;
            _ = Shell_NotifyIcon(NimDelete, ref data);
            _trayIconAdded = false;
            _trayIconData = default;
        }

        if (_trayIconHandle != IntPtr.Zero)
        {
            _ = DestroyIcon(_trayIconHandle);
            _trayIconHandle = IntPtr.Zero;
        }
    }

    private void BuildCards()
    {
        RebuildCards();
    }

    private void SearchToggleButton_Click(object sender, RoutedEventArgs e)
    {
        ExpandSearchBar();
    }

    private void SearchCloseButton_Click(object sender, RoutedEventArgs e)
    {
        CollapseSearchBar(clearQuery: true);
    }

    private void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        var nextQuery = SearchTextBox.Text?.Trim() ?? string.Empty;
        if (string.Equals(_searchQuery, nextQuery, StringComparison.Ordinal))
        {
            return;
        }

        _searchQuery = nextQuery;
        RebuildCards();
    }

    private void SearchTextBox_KeyDown(object sender, KeyRoutedEventArgs e)
    {
        if (e.Key == Windows.System.VirtualKey.F && IsControlPressed())
        {
            ExpandSearchBar();
            e.Handled = true;
            return;
        }

        if (e.Key != Windows.System.VirtualKey.Escape)
        {
            return;
        }

        CollapseSearchBar(clearQuery: true);
        e.Handled = true;
    }

    private void ExpandSearchBar()
    {
        if (IsSearchBarExpanded() || (_searchAnimationTimer?.IsEnabled == true && _searchAnimationExpanding))
        {
            SearchTextBox.IsEnabled = true;
            _ = SearchTextBox.Focus(FocusState.Programmatic);
            SearchTextBox.Select(SearchTextBox.Text.Length, 0);
            return;
        }

        SearchTextBox.IsEnabled = true;
        SearchToggleButton.Visibility = Visibility.Collapsed;
        SearchBarShell.Visibility = Visibility.Visible;
        SetSearchAnimationProgress(0.0);
        StartSearchAnimation(expanding: true, clearQueryOnCollapse: false);
    }

    private bool IsSearchBarExpanded()
    {
        return SearchBarShell.Visibility == Visibility.Visible;
    }

    private bool IsSearchInputActive()
    {
        return IsSearchBarExpanded() && SearchTextBox.FocusState != FocusState.Unfocused;
    }

    private static bool IsControlPressed()
    {
        return (GetAsyncKeyState(VkControlKey) & 0x8000) != 0 ||
            (GetAsyncKeyState(VkLcontrol) & 0x8000) != 0 ||
            (GetAsyncKeyState(VkRcontrol) & 0x8000) != 0;
    }

    private void StartSearchAnimation(bool expanding, bool clearQueryOnCollapse, bool restoreFocusOnCollapse = true)
    {
        _searchAnimationExpanding = expanding;
        _searchAnimationClearQueryOnCollapse = !expanding && clearQueryOnCollapse;
        _searchAnimationRestoreFocusOnCollapse = !expanding && restoreFocusOnCollapse;
        _searchAnimationStart = DateTimeOffset.UtcNow;
        _searchAnimationTimer ??= new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(16) };
        _searchAnimationTimer.Tick -= OnSearchAnimationTick;
        _searchAnimationTimer.Tick += OnSearchAnimationTick;
        _searchAnimationTimer.Start();
    }

    private void OnSearchAnimationTick(object? sender, object e)
    {
        var elapsed = (DateTimeOffset.UtcNow - _searchAnimationStart).TotalMilliseconds;
        var t = Math.Clamp(elapsed / SearchAnimationDurationMs, 0.0, 1.0);
        var progress = _searchAnimationExpanding
            ? EaseOutCubic(t)
            : 1.0 - EaseInCubic(t);
        SetSearchAnimationProgress(progress);

        if (t < 1.0)
        {
            return;
        }

        _searchAnimationTimer?.Stop();
        if (_searchAnimationExpanding)
        {
            SetSearchAnimationProgress(1.0);
            _ = SearchTextBox.Focus(FocusState.Programmatic);
            SearchTextBox.Select(SearchTextBox.Text.Length, 0);
            return;
        }

        SetSearchAnimationProgress(0.0);
        SearchBarShell.Visibility = Visibility.Collapsed;
        SearchTextBox.IsEnabled = false;
        SearchToggleButton.Visibility = Visibility.Visible;
        if (_searchAnimationRestoreFocusOnCollapse)
        {
            ActivateAndFocusInput();
        }

        if (!_searchAnimationClearQueryOnCollapse)
        {
            return;
        }

        if (_searchQuery.Length == 0 && string.IsNullOrEmpty(SearchTextBox.Text))
        {
            return;
        }

        _searchQuery = string.Empty;
        SearchTextBox.Text = string.Empty;
        RebuildCards();
    }

    private void SetSearchAnimationProgress(double progress)
    {
        SearchBarShell.Opacity = 0.18 + (progress * 0.82);
        SearchBarScaleTransform.ScaleX = 0.65 + (progress * 0.35);
        SearchBarScaleTransform.ScaleY = 0.96 + (progress * 0.04);
        SearchCloseButton.Opacity = Math.Clamp(progress * 1.2, 0.0, 1.0);
        SearchToggleButton.Opacity = Math.Clamp(1.0 - (progress * 1.35), 0.0, 1.0);
    }

    private void CollapseSearchBar(bool clearQuery, bool restoreFocus = true)
    {
        if (!IsSearchBarExpanded() && !(_searchAnimationTimer?.IsEnabled == true && _searchAnimationExpanding))
        {
            if (clearQuery)
            {
                if (_searchQuery.Length == 0 && string.IsNullOrEmpty(SearchTextBox.Text))
                {
                    SearchTextBox.IsEnabled = false;
                    if (restoreFocus)
                    {
                        ActivateAndFocusInput();
                    }
                    return;
                }

                _searchQuery = string.Empty;
                SearchTextBox.Text = string.Empty;
                RebuildCards();
            }
            SearchTextBox.IsEnabled = false;
            if (restoreFocus)
            {
                ActivateAndFocusInput();
            }
            return;
        }

        StartSearchAnimation(expanding: false, clearQueryOnCollapse: clearQuery, restoreFocusOnCollapse: restoreFocus);
    }

    private void UpdateFilteredEntries()
    {
        _filteredEntries.Clear();

        if (string.IsNullOrWhiteSpace(_searchQuery))
        {
            _filteredEntries.AddRange(_entries);
            return;
        }

        foreach (var entry in _entries)
        {
            if (EntryMatchesSearch(entry, _searchQuery))
            {
                _filteredEntries.Add(entry);
            }
        }
    }

    private static bool EntryMatchesSearch(ClipboardEntry entry, string query)
    {
        var comparison = StringComparison.OrdinalIgnoreCase;
        if (entry.Kind.Contains(query, comparison) ||
            entry.SourceApp.Contains(query, comparison) ||
            entry.Content.Contains(query, comparison))
        {
            return true;
        }

        return (!string.IsNullOrWhiteSpace(entry.LinkTitle) && entry.LinkTitle.Contains(query, comparison)) ||
            (!string.IsNullOrWhiteSpace(entry.LinkUrl) && entry.LinkUrl.Contains(query, comparison)) ||
            (!string.IsNullOrWhiteSpace(entry.LinkHost) && entry.LinkHost.Contains(query, comparison));
    }

    private void RebuildCards()
    {
        UpdateFilteredEntries();
        HideCardPreview();
        HideCardContextPopup();

        CardRow.Children.Clear();
        _cards.Clear();
        _cardHosts.Clear();
        _selectionOutlines.Clear();
        _agoLabels.Clear();
        _selectedCardIndex = -1;

        for (var i = 0; i < _filteredEntries.Count; i++)
        {
            var entry = _filteredEntries[i];
            AddCardFromEntry(entry);
        }

        if (_cards.Count > 0)
        {
            SetSelectedCard(0);
        }
        UpdateCardSize();
    }

    private void ResetCardScrollToLeft()
    {
        _selectionScrollTimer?.Stop();
        _scrollInertiaTimer?.Stop();
        _scrollVelocity = 0;
        CardScroller.ChangeView(0, null, null, true);
    }

    private void ResetSelectionToFirstCard()
    {
        if (_cards.Count == 0)
        {
            return;
        }

        _selectedCardIndex = -1;
        SetSelectedCard(0);
    }

    private void DeleteEntry(ClipboardEntry entry)
    {
        if (!_entries.Remove(entry))
        {
            return;
        }

        HideCardPreview();
        HideCardContextPopup();
        RebuildCards();
        RefreshAgoLabels();
        ScheduleHistorySave();
    }

    private void AddCardFromEntry(ClipboardEntry entry)
    {
        var cardIndex = _cards.Count;
        var isLinkCard = entry.Kind == "Link";
        var cardSurfaceColor = Color.FromArgb(238, 22, 27, 32);
        var cardSurfaceHoverColor = Color.FromArgb(248, 28, 34, 41);
        var cardBorderDefaultColor = Color.FromArgb(56, 255, 255, 255);
        var cardBorderHoverColor = Color.FromArgb(140, 255, 180, 105);
        var cardSurfaceBrush = new SolidColorBrush(cardSurfaceColor);
        var cardBorderBrush = new SolidColorBrush(cardBorderDefaultColor);

        var card = new Border
        {
            CornerRadius = new CornerRadius(18),
            Background = cardSurfaceBrush,
            BorderBrush = cardBorderBrush,
            BorderThickness = new Thickness(1),
            HorizontalAlignment = HorizontalAlignment.Stretch,
            VerticalAlignment = VerticalAlignment.Stretch
        };

        var cardHost = new Grid
        {
            Width = 120,
            Height = 120
        };
        Windows.Foundation.TypedEventHandler<UIElement, ContextRequestedEventArgs> onCardContextRequested = (_, e) =>
        {
            ShowCardContextPopup(entry, cardIndex, cardHost);
            e.Handled = true;
        };
        cardHost.AddHandler(UIElement.ContextRequestedEvent, onCardContextRequested, true);
        card.AddHandler(UIElement.ContextRequestedEvent, onCardContextRequested, true);
        cardHost.PointerEntered += (_, _) =>
        {
            cardBorderBrush.Color = cardBorderHoverColor;
            cardSurfaceBrush.Color = cardSurfaceHoverColor;
        };
        cardHost.PointerExited += (_, _) =>
        {
            cardBorderBrush.Color = cardBorderDefaultColor;
            cardSurfaceBrush.Color = cardSurfaceColor;
        };

        cardHost.PointerPressed += (_, e) =>
        {
            var point = e.GetCurrentPoint(cardHost);
            if (point.Properties.IsRightButtonPressed)
            {
                ShowCardContextPopup(entry, cardIndex, cardHost);
                e.Handled = true;
                return;
            }

            if (!point.Properties.IsLeftButtonPressed)
            {
                return;
            }

            SetSelectedCard(cardIndex);
            HideCardContextPopup();
            ActivateAndFocusInput();
            e.Handled = true;
        };

        var selectionOutline = new Border
        {
            CornerRadius = new CornerRadius(18),
            BorderBrush = new SolidColorBrush(Color.FromArgb(255, 255, 180, 105)),
            BorderThickness = new Thickness(2),
            Margin = new Thickness(0),
            Opacity = 0.0,
            IsHitTestVisible = false
        };

        var cardLayout = new Grid();
        cardLayout.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        cardLayout.RowDefinitions.Add(new RowDefinition { Height = new GridLength(isLinkCard ? 3 : 4, GridUnitType.Star) });
        if (isLinkCard)
        {
            cardLayout.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        }

        var topBar = new Border
        {
            Background = new SolidColorBrush(entry.SourceHeaderColor),
            CornerRadius = new CornerRadius(17, 17, 0, 0)
        };

        var topBarGrid = new Grid();
        topBarGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        topBarGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        var headerKindFontSize = isLinkCard ? 11.5 : 11.0;
        var headerAgoFontSize = isLinkCard ? 9.5 : 9.0;
        var headerTextOffset = isLinkCard ? 2.0 : 0.0;
        var topBarContent = new StackPanel
        {
            Orientation = Orientation.Vertical,
            Spacing = 1,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = isLinkCard ? new Thickness(12, 5, 8, 5) : new Thickness(10, 5, 10, 5)
        };
        topBarContent.Children.Add(new TextBlock
        {
            Text = entry.Kind,
            FontSize = headerKindFontSize,
            FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
            Foreground = new SolidColorBrush(Color.FromArgb(230, 255, 255, 255)),
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(headerTextOffset, 0, 0, 0)
        });
        var agoLabel = new TextBlock
        {
            Text = FormatAgo(entry.CopiedAtUtc),
            FontSize = headerAgoFontSize,
            Foreground = new SolidColorBrush(Color.FromArgb(180, 255, 255, 255)),
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(headerTextOffset, 0, 0, 0)
        };
        topBarContent.Children.Add(agoLabel);
        _agoLabels.Add((entry, agoLabel));
        topBarGrid.Children.Add(topBarContent);

        var iconBadge = CreateAppIconBadge(entry.SourceApp, entry.SourceExePath, entry.SourceIconPngBytes);
        Grid.SetColumn(iconBadge, 1);
        topBarGrid.Children.Add(iconBadge);

        topBar.Child = topBarGrid;

        var body = new Border
        {
            Background = cardSurfaceBrush
        };

        var isImageCard = entry.Kind == "Image" && entry.ImagePngBytes is { Length: > 0 };
        var isLinkPreviewCard = isLinkCard && entry.LinkPreviewImageBytes is { Length: > 0 };
        var isLinkFaviconCard = isLinkCard && entry.LinkFaviconImageBytes is { Length: > 0 };
        var usesBodyFooter = !isImageCard && !isLinkCard;
        var bodyGrid = new Grid();
        var textMargin = (isImageCard || isLinkCard) ? new Thickness(0) : new Thickness(10, 8, 10, 8);
        var bodyContent = new Grid
        {
            Margin = textMargin
        };
        bodyContent.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        bodyContent.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        bodyGrid.Children.Add(bodyContent);

        var textArea = new Grid
        {
            VerticalAlignment = VerticalAlignment.Top
        };
        Grid.SetRow(textArea, 0);
        bodyContent.Children.Add(textArea);

        TextBlock? bodyText = null;
        Rectangle? fadeOverlay = null;
        Image? previewImage = null;
        Border? imageSizeBadge = null;
        Border? linkFallbackVisual = null;

        if (isImageCard || isLinkPreviewCard)
        {
            previewImage = new Image
            {
                Stretch = Stretch.UniformToFill,
                HorizontalAlignment = HorizontalAlignment.Stretch,
                VerticalAlignment = VerticalAlignment.Stretch
            };
            textArea.Children.Add(previewImage);
            var imageBytes = isImageCard ? entry.ImagePngBytes : entry.LinkPreviewImageBytes;
            _ = TrySetClipboardImagePreviewAsync(imageBytes, previewImage);
        }
        else if (isLinkCard)
        {
            var host = ResolveLinkHost(entry);
            linkFallbackVisual = CreateLinkFallbackVisual(host, !isLinkFaviconCard);
            textArea.Children.Add(linkFallbackVisual);

            if (isLinkFaviconCard)
            {
                previewImage = new Image
                {
                    Stretch = Stretch.Uniform,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center,
                    IsHitTestVisible = false
                };
                textArea.Children.Add(previewImage);
                _ = TrySetClipboardImagePreviewAsync(entry.LinkFaviconImageBytes, previewImage);
            }
        }
        else
        {
            bodyText = new TextBlock
            {
                Text = entry.Content,
                TextWrapping = TextWrapping.WrapWholeWords,
                TextTrimming = TextTrimming.None,
                FontSize = 13,
                LineHeight = 20,
                Foreground = new SolidColorBrush(Color.FromArgb(220, 245, 245, 245)),
                Opacity = 1.0,
                VerticalAlignment = VerticalAlignment.Top
            };
            textArea.Children.Add(bodyText);

            fadeOverlay = new Rectangle
            {
                VerticalAlignment = VerticalAlignment.Bottom,
                IsHitTestVisible = false
            };
            textArea.Children.Add(fadeOverlay);
        }

        var footerText = entry.Kind == "Image"
            ? (entry.ImageWidth > 0 && entry.ImageHeight > 0
                ? $"{entry.ImageWidth:N0}×{entry.ImageHeight:N0}"
                : "画像")
            : $"{entry.Content.Length:N0}文字";
        const double footerLabelHeight = 18.0;

        var footerLabel = new TextBlock
        {
            Text = footerText,
            FontSize = 11,
            Foreground = new SolidColorBrush(Color.FromArgb(224, 236, 240, 245)),
            HorizontalAlignment = HorizontalAlignment.Center,
            TextAlignment = TextAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
            Height = double.NaN,
            LineHeight = 14,
            FontWeight = Microsoft.UI.Text.FontWeights.SemiBold
        };

        if (isImageCard)
        {
            imageSizeBadge = new Border
            {
                Background = new SolidColorBrush(Color.FromArgb(176, 20, 26, 32)),
                CornerRadius = new CornerRadius(9),
                Padding = new Thickness(8, 2, 8, 2),
                MinHeight = footerLabelHeight,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Bottom,
                Margin = new Thickness(0, 0, 0, 10),
                IsHitTestVisible = false,
                Child = footerLabel
            };
            textArea.Children.Add(imageSizeBadge);
        }
        else if (usesBodyFooter)
        {
            footerLabel.Height = footerLabelHeight;
            footerLabel.LineHeight = footerLabelHeight;
            Grid.SetRow(footerLabel, 1);
            bodyContent.Children.Add(footerLabel);
        }

        bodyGrid.SizeChanged += (_, _) =>
        {
            var contentHeight = Math.Max(0, bodyGrid.ActualHeight - textMargin.Top - textMargin.Bottom);
            bodyContent.Height = contentHeight;
            var footerHeight = usesBodyFooter ? footerLabelHeight : 0.0;
            var textAreaHeight = Math.Max(0, contentHeight - footerHeight);

            textArea.Height = textAreaHeight;
            if (usesBodyFooter)
            {
                footerLabel.Height = footerHeight;
            }
            bodyContent.RowDefinitions[1].Height = new GridLength(footerHeight);

            if (previewImage is not null)
            {
                previewImage.Height = isLinkFaviconCard ? Math.Clamp(textAreaHeight * 0.45, 42.0, 84.0) : textAreaHeight;
                if (isLinkFaviconCard)
                {
                    previewImage.Width = previewImage.Height;
                }
            }

            if (imageSizeBadge is not null)
            {
                imageSizeBadge.Margin = new Thickness(0, 0, 0, Math.Max(8.0, textAreaHeight * 0.06));
            }

            if (bodyText is null || fadeOverlay is null)
            {
                return;
            }

            bodyText.Height = textAreaHeight;
            var fadeHeight = Math.Min(textAreaHeight, bodyText.LineHeight * 3.3);
            fadeOverlay.Height = fadeHeight;
            fadeOverlay.Fill = new LinearGradientBrush
            {
                StartPoint = new Windows.Foundation.Point(0, 0),
                EndPoint = new Windows.Foundation.Point(0, 1),
                GradientStops =
                {
                    new GradientStop { Offset = 0.00, Color = Color.FromArgb(0, cardSurfaceColor.R, cardSurfaceColor.G, cardSurfaceColor.B) },
                    new GradientStop { Offset = 0.25, Color = Color.FromArgb(0, cardSurfaceColor.R, cardSurfaceColor.G, cardSurfaceColor.B) },
                    new GradientStop { Offset = 0.80, Color = Color.FromArgb(210, cardSurfaceColor.R, cardSurfaceColor.G, cardSurfaceColor.B) },
                    new GradientStop { Offset = 1.00, Color = Color.FromArgb(255, cardSurfaceColor.R, cardSurfaceColor.G, cardSurfaceColor.B) }
                }
            };
        };
        body.Child = bodyGrid;
        Grid.SetRow(body, 1);

        cardLayout.Children.Add(topBar);
        cardLayout.Children.Add(body);
        if (isLinkCard)
        {
            var linkFooter = CreateLinkFooter(entry);
            Grid.SetRow(linkFooter, 2);
            cardLayout.Children.Add(linkFooter);
        }
        card.Child = cardLayout;

        cardHost.Children.Add(card);
        cardHost.Children.Add(selectionOutline);

        CardRow.Children.Add(cardHost);
        _cardHosts.Add(cardHost);
        _selectionOutlines.Add(selectionOutline);
        _cards.Add(card);
    }

    private static Border CreateLinkFooter(ClipboardEntry entry)
    {
        var linkUrl = string.IsNullOrWhiteSpace(entry.LinkUrl) ? entry.Content : entry.LinkUrl!;
        var title = ResolveLinkTitleForFooter(entry, linkUrl);

        var titleText = new TextBlock
        {
            Text = title,
            FontSize = 10.5,
            FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
            Foreground = new SolidColorBrush(Color.FromArgb(238, 245, 247, 250)),
            TextWrapping = TextWrapping.NoWrap,
            TextTrimming = TextTrimming.CharacterEllipsis,
            Margin = new Thickness(1, 0, 0, 0)
        };

        var linkText = new TextBlock
        {
            Text = linkUrl,
            FontSize = 9.0,
            Foreground = new SolidColorBrush(Color.FromArgb(208, 218, 225, 234)),
            TextWrapping = TextWrapping.NoWrap,
            TextTrimming = TextTrimming.CharacterEllipsis,
            Margin = new Thickness(1, 0, 0, 0)
        };

        var content = new StackPanel
        {
            Orientation = Orientation.Vertical,
            Spacing = 0,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(12, 4, 8, 4)
        };
        content.Children.Add(titleText);
        content.Children.Add(linkText);

        return new Border
        {
            Background = new SolidColorBrush(Color.FromArgb(238, 23, 29, 35)),
            CornerRadius = new CornerRadius(0, 0, 17, 17),
            Child = content
        };
    }

    private static Border CreateLinkFallbackVisual(string host, bool showInitial)
    {
        var backgroundColor = GetLinkPlaceholderColor(host);
        var border = new Border
        {
            Background = new SolidColorBrush(backgroundColor),
            CornerRadius = new CornerRadius(0)
        };

        var layout = new Grid();

        var hostLabel = new TextBlock
        {
            Text = host,
            FontSize = 10,
            Foreground = new SolidColorBrush(Color.FromArgb(180, 255, 255, 255)),
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Bottom,
            Margin = new Thickness(8, 0, 8, 8),
            TextTrimming = TextTrimming.CharacterEllipsis,
            TextWrapping = TextWrapping.NoWrap
        };
        layout.Children.Add(hostLabel);

        if (showInitial)
        {
            var initial = GetLinkInitial(host);
            var initialLabel = new TextBlock
            {
                Text = initial,
                FontSize = 46,
                FontWeight = Microsoft.UI.Text.FontWeights.Bold,
                Foreground = new SolidColorBrush(Color.FromArgb(238, 255, 255, 255)),
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, -4, 0, 0)
            };
            layout.Children.Add(initialLabel);
        }

        border.Child = layout;
        return border;
    }

    private static Color GetLinkPlaceholderColor(string host)
    {
        var hash = GetStableHash(host.ToLowerInvariant());
        var palette = new[]
        {
            Color.FromArgb(255, 47, 79, 109),
            Color.FromArgb(255, 77, 64, 122),
            Color.FromArgb(255, 44, 88, 74),
            Color.FromArgb(255, 88, 63, 63),
            Color.FromArgb(255, 62, 72, 94),
            Color.FromArgb(255, 70, 70, 70)
        };
        return palette[Math.Abs(hash % palette.Length)];
    }

    private static int GetStableHash(string input)
    {
        unchecked
        {
            var hash = (int)2166136261;
            foreach (var ch in input)
            {
                hash ^= ch;
                hash *= 16777619;
            }
            return hash;
        }
    }

    private static string GetLinkInitial(string host)
    {
        foreach (var ch in host)
        {
            if (char.IsLetterOrDigit(ch))
            {
                return char.ToUpperInvariant(ch).ToString();
            }
        }
        return "L";
    }

    private static string ResolveLinkHost(ClipboardEntry entry)
    {
        if (!string.IsNullOrWhiteSpace(entry.LinkHost))
        {
            return entry.LinkHost!;
        }

        var linkUrl = string.IsNullOrWhiteSpace(entry.LinkUrl) ? entry.Content : entry.LinkUrl!;
        if (Uri.TryCreate(linkUrl, UriKind.Absolute, out var uri) && !string.IsNullOrWhiteSpace(uri.Host))
        {
            return uri.Host;
        }

        return "link";
    }

    private static string ResolveLinkTitleForFooter(ClipboardEntry entry, string linkUrl)
    {
        if (!string.IsNullOrWhiteSpace(entry.LinkTitle))
        {
            return entry.LinkTitle!;
        }

        if (Uri.TryCreate(linkUrl, UriKind.Absolute, out var uri) && !string.IsNullOrWhiteSpace(uri.Host))
        {
            return uri.Host;
        }

        return "Link";
    }

    private void MoveSelectedCard(int delta)
    {
        if (_cards.Count == 0)
        {
            return;
        }

        var next = _selectedCardIndex < 0 ? 0 : _selectedCardIndex + delta;
        next = Math.Clamp(next, 0, _cards.Count - 1);
        SetSelectedCard(next);
        EnsureSelectedCardVisibleByStep(delta);
    }

    private void SetSelectedCard(int index)
    {
        if (index < 0 || index >= _cards.Count)
        {
            return;
        }

        if (_selectedCardIndex != index)
        {
            HideCardPreview();
        }

        _selectedCardIndex = index;

        for (var i = 0; i < _cards.Count; i++)
        {
            var isSelected = i == _selectedCardIndex;
            _selectionOutlines[i].Opacity = isSelected ? 1.0 : 0.0;
        }
    }

    private void EnsureSelectedCardVisibleByStep(int direction)
    {
        if (_selectedCardIndex < 0 || _selectedCardIndex >= _cardHosts.Count)
        {
            return;
        }

        var host = _cardHosts[_selectedCardIndex];
        var cardWidth = host.Width > 0 ? host.Width : host.ActualWidth;
        if (cardWidth <= 0 || CardScroller.ActualWidth <= 0)
        {
            return;
        }

        var step = cardWidth + CardRow.Spacing;
        var cardLeft = _selectedCardIndex * step;
        var cardRight = cardLeft + cardWidth;

        var viewportLeft = CardScroller.HorizontalOffset;
        var viewportRight = viewportLeft + CardScroller.ActualWidth;

        double? target = null;
        if (direction > 0 && cardRight > viewportRight)
        {
            // Move enough to bring the selected card fully into view.
            target = cardRight - CardScroller.ActualWidth;
        }
        else if (direction < 0 && cardLeft < viewportLeft)
        {
            target = cardLeft;
        }

        if (!target.HasValue)
        {
            return;
        }

        var clamped = Math.Clamp(target.Value, 0, CardScroller.ScrollableWidth);
        StartSelectionScrollAnimation(clamped);
    }

    private void StartSelectionScrollAnimation(double target)
    {
        if (Math.Abs(target - CardScroller.HorizontalOffset) < 0.5)
        {
            CardScroller.ChangeView(target, null, null, true);
            return;
        }

        _selectionScrollTo = target;
        _selectionScrollLastTick = DateTimeOffset.UtcNow;

        _selectionScrollTimer ??= new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(10) };
        _selectionScrollTimer.Tick -= OnSelectionScrollTick;
        _selectionScrollTimer.Tick += OnSelectionScrollTick;
        if (!_selectionScrollTimer.IsEnabled)
        {
            _selectionScrollTimer.Start();
        }
    }

    private void OnSelectionScrollTick(object? sender, object e)
    {
        var now = DateTimeOffset.UtcNow;
        var dt = Math.Max(0.001, (now - _selectionScrollLastTick).TotalSeconds);
        _selectionScrollLastTick = now;

        var current = CardScroller.HorizontalOffset;
        var alpha = 1.0 - Math.Exp(-SelectionScrollSmoothingStrength * dt);
        var next = current + ((_selectionScrollTo - current) * alpha);
        next = Math.Clamp(next, 0, CardScroller.ScrollableWidth);
        CardScroller.ChangeView(next, null, null, true);

        if (Math.Abs(_selectionScrollTo - next) > 0.5)
        {
            return;
        }

        _selectionScrollTimer?.Stop();
        CardScroller.ChangeView(_selectionScrollTo, null, null, true);
    }

    private FrameworkElement CreateAppIconBadge(string appName, string? sourceExePath, byte[]? sourceIconPngBytes)
    {
        var badge = new Border
        {
            Width = 60,
            Height = 60,
            CornerRadius = new CornerRadius(16),
            Background = new SolidColorBrush(Color.FromArgb(0, 0, 0, 0)),
            HorizontalAlignment = HorizontalAlignment.Right,
            VerticalAlignment = VerticalAlignment.Stretch,
            Margin = new Thickness(0, -10, -10, -10),
            Opacity = 0.0
        };

        _ = TrySetAppIconAsync(sourceIconPngBytes, sourceExePath, badge);
        return badge;
    }

    private async System.Threading.Tasks.Task TrySetAppIconAsync(byte[]? windowIconPngBytes, string? exePath, Border badge)
    {
        BitmapImage? imageSource = null;

        if (windowIconPngBytes is not null && windowIconPngBytes.Length > 0)
        {
            imageSource = await LoadPngBytesAsync(windowIconPngBytes);
        }

        if (imageSource is null && !string.IsNullOrWhiteSpace(exePath) && File.Exists(exePath))
        {
            imageSource = await LoadExecutableIconAsync(exePath);
        }

        if (imageSource is null)
        {
            return;
        }

        badge.Opacity = 1.0;
        badge.Child = new Border
        {
            CornerRadius = new CornerRadius(16),
            Clip = null,
            Child = new Image
            {
                Source = imageSource,
                Stretch = Stretch.UniformToFill,
                HorizontalAlignment = HorizontalAlignment.Stretch,
                VerticalAlignment = VerticalAlignment.Stretch
            }
        };
    }

    private static async System.Threading.Tasks.Task TrySetClipboardImagePreviewAsync(byte[]? imagePngBytes, Image previewImage)
    {
        if (imagePngBytes is null || imagePngBytes.Length == 0)
        {
            return;
        }

        var imageSource = await LoadPngBytesAsync(imagePngBytes);
        if (imageSource is null)
        {
            return;
        }

        previewImage.Source = imageSource;
    }

    private static async System.Threading.Tasks.Task<BitmapImage?> LoadPngBytesAsync(byte[] pngBytes)
    {
        try
        {
            using var memory = new MemoryStream(pngBytes, writable: false);
            using var ras = memory.AsRandomAccessStream();
            var image = new BitmapImage();
            await image.SetSourceAsync(ras);
            return image;
        }
        catch
        {
            return null;
        }
    }

    private static async System.Threading.Tasks.Task<BitmapImage?> LoadExecutableIconAsync(string exePath)
    {
        try
        {
            using var icon = System.Drawing.Icon.ExtractAssociatedIcon(exePath);
            if (icon is null)
            {
                return null;
            }

            using var bitmap = icon.ToBitmap();
            using var memory = new MemoryStream();
            bitmap.Save(memory, System.Drawing.Imaging.ImageFormat.Png);
            memory.Position = 0;
            using var ras = memory.AsRandomAccessStream();

            var image = new BitmapImage();
            await image.SetSourceAsync(ras);
            return image;
        }
        catch
        {
            return null;
        }
    }

    private static byte[]? TryGetWindowIconPngBytes(IntPtr hwnd)
    {
        try
        {
            var hIcon = SendMessage(hwnd, WmGetIcon, (IntPtr)IconBig, IntPtr.Zero);
            if (hIcon == IntPtr.Zero)
            {
                hIcon = SendMessage(hwnd, WmGetIcon, (IntPtr)IconSmall2, IntPtr.Zero);
            }
            if (hIcon == IntPtr.Zero)
            {
                hIcon = SendMessage(hwnd, WmGetIcon, (IntPtr)IconSmall, IntPtr.Zero);
            }
            if (hIcon == IntPtr.Zero)
            {
                hIcon = GetClassLongPtr(hwnd, GclHicon);
            }
            if (hIcon == IntPtr.Zero)
            {
                hIcon = GetClassLongPtr(hwnd, GclHiconSm);
            }
            if (hIcon == IntPtr.Zero)
            {
                return null;
            }

            var copiedIcon = CopyIcon(hIcon);
            if (copiedIcon == IntPtr.Zero)
            {
                return null;
            }

            using var icon = (System.Drawing.Icon)System.Drawing.Icon.FromHandle(copiedIcon).Clone();
            _ = DestroyIcon(copiedIcon);
            using var bitmap = icon.ToBitmap();
            using var memory = new MemoryStream();
            bitmap.Save(memory, System.Drawing.Imaging.ImageFormat.Png);
            return memory.ToArray();
        }
        catch
        {
            return null;
        }
    }

    private static IntPtr GetClassLongPtr(IntPtr hWnd, int nIndex)
    {
        return IntPtr.Size == 8
            ? GetClassLongPtr64(hWnd, nIndex)
            : new IntPtr((long)GetClassLong32(hWnd, nIndex));
    }

    private void UpdateCardSize()
    {
        if (_cards.Count == 0 || OverlayPanel.ActualWidth <= 0)
        {
            return;
        }

        var available = CardScroller.ActualWidth > 0
            ? CardScroller.ActualWidth
            : OverlayPanel.ActualWidth - (OverlayInnerPadding * 2);
        var availableHeight = CardScroller.ActualHeight > 0
            ? CardScroller.ActualHeight
            : OverlayPanel.ActualHeight - (OverlayInnerPadding * 2);
        var gap = Math.Max(8.0, available * CardGapRatio);
        CardRow.Spacing = gap;

        var targetByWidth = (available - (gap * (VisibleCardCount - 1.0))) / VisibleCardCount;
        var targetByHeight = availableHeight * CardHeightRatio;
        var cardSize = Math.Clamp(Math.Min(targetByWidth, targetByHeight), 52.0, 280.0);

        foreach (var host in _cardHosts)
        {
            host.Width = cardSize;
            host.Height = cardSize;
        }
    }

    private sealed class ClipboardEntry
    {
        public ClipboardEntry(
            string kind,
            string sourceApp,
            string content,
            DateTimeOffset copiedAtUtc,
            string? sourceExePath,
            byte[]? sourceIconPngBytes,
            Color sourceHeaderColor,
            byte[]? imagePngBytes = null,
            int imageWidth = 0,
            int imageHeight = 0,
            string? linkUrl = null,
            string? linkTitle = null,
            byte[]? linkPreviewImageBytes = null,
            byte[]? linkFaviconImageBytes = null,
            string? linkHost = null)
        {
            Kind = kind;
            SourceApp = sourceApp;
            Content = content;
            CopiedAtUtc = copiedAtUtc;
            SourceExePath = sourceExePath;
            SourceIconPngBytes = sourceIconPngBytes;
            SourceHeaderColor = sourceHeaderColor;
            ImagePngBytes = imagePngBytes;
            ImageWidth = imageWidth;
            ImageHeight = imageHeight;
            LinkUrl = linkUrl;
            LinkTitle = linkTitle;
            LinkPreviewImageBytes = linkPreviewImageBytes;
            LinkFaviconImageBytes = linkFaviconImageBytes;
            LinkHost = linkHost;
        }

        public string Kind { get; }
        public string SourceApp { get; }
        public string Content { get; }
        public DateTimeOffset CopiedAtUtc { get; }
        public string? SourceExePath { get; }
        public byte[]? SourceIconPngBytes { get; }
        public Color SourceHeaderColor { get; }
        public byte[]? ImagePngBytes { get; }
        public int ImageWidth { get; }
        public int ImageHeight { get; }
        public string? LinkUrl { get; }
        public string? LinkTitle { get; }
        public byte[]? LinkPreviewImageBytes { get; }
        public byte[]? LinkFaviconImageBytes { get; }
        public string? LinkHost { get; }
    }

    private sealed class PersistedClipboardEntry
    {
        public string? Kind { get; set; }
        public string? SourceApp { get; set; }
        public string? Content { get; set; }
        public DateTimeOffset CopiedAtUtc { get; set; }
        public string? SourceExePath { get; set; }
        public byte[]? SourceIconPngBytes { get; set; }
        public uint SourceHeaderColorArgb { get; set; }
        public byte[]? ImagePngBytes { get; set; }
        public int ImageWidth { get; set; }
        public int ImageHeight { get; set; }
        public string? LinkUrl { get; set; }
        public string? LinkTitle { get; set; }
        public byte[]? LinkPreviewImageBytes { get; set; }
        public byte[]? LinkFaviconImageBytes { get; set; }
        public string? LinkHost { get; set; }
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct NOTIFYICONDATA
    {
        public uint cbSize;
        public IntPtr hWnd;
        public uint uID;
        public uint uFlags;
        public uint uCallbackMessage;
        public IntPtr hIcon;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
        public string szTip;
        public uint dwState;
        public uint dwStateMask;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]
        public string szInfo;
        public uint uVersion;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 64)]
        public string szInfoTitle;
        public uint dwInfoFlags;
        public Guid guidItem;
        public IntPtr hBalloonIcon;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct INPUT
    {
        public uint type;
        public InputUnion U;
    }

    [StructLayout(LayoutKind.Explicit)]
    private struct InputUnion
    {
        [FieldOffset(0)]
        public MOUSEINPUT mi;

        [FieldOffset(0)]
        public KEYBDINPUT ki;

        [FieldOffset(0)]
        public HARDWAREINPUT hi;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MOUSEINPUT
    {
        public int dx;
        public int dy;
        public uint mouseData;
        public uint dwFlags;
        public uint time;
        public IntPtr dwExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct KEYBDINPUT
    {
        public ushort wVk;
        public ushort wScan;
        public uint dwFlags;
        public uint time;
        public IntPtr dwExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct HARDWAREINPUT
    {
        public uint uMsg;
        public ushort wParamL;
        public ushort wParamH;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct KBDLLHOOKSTRUCT
    {
        public uint vkCode;
        public uint scanCode;
        public uint flags;
        public uint time;
        public IntPtr dwExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MSLLHOOKSTRUCT
    {
        public POINT pt;
        public uint mouseData;
        public uint flags;
        public uint time;
        public IntPtr dwExtraInfo;
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

    [StructLayout(LayoutKind.Sequential)]
    private struct GUITHREADINFO
    {
        public int cbSize;
        public uint flags;
        public IntPtr hwndActive;
        public IntPtr hwndFocus;
        public IntPtr hwndCapture;
        public IntPtr hwndMenuOwner;
        public IntPtr hwndMoveSize;
        public IntPtr hwndCaret;
        public RECT rcCaret;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
    private struct MONITORINFO
    {
        public int cbSize;
        public RECT rcMonitor;
        public RECT rcWork;
        public int dwFlags;
    }

}

