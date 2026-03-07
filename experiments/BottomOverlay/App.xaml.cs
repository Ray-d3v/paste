using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

namespace BottomOverlay;

public partial class App : Application
{
    private const int HotkeyId = 1;
    private const uint ModAlt = 0x0001;
    private const uint ModControl = 0x0002;
    private const uint VkV = 0x56;
    private const int WmHotkey = 0x0312;

    private MainWindow? _overlayWindow;
    private HwndSource? _overlaySource;
    private IntPtr _overlayHandle = IntPtr.Zero;

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        _overlayWindow = new MainWindow();

        var helper = new WindowInteropHelper(_overlayWindow);
        _overlayHandle = helper.EnsureHandle();

        _overlaySource = HwndSource.FromHwnd(_overlayHandle);
        _overlaySource?.AddHook(WndProc);

        var success = RegisterHotKey(_overlayHandle, HotkeyId, ModControl | ModAlt, VkV);
        if (!success)
        {
            var error = Marshal.GetLastWin32Error();
            var message = new Win32Exception(error).Message;
            MessageBox.Show(
                $"Ctrl+Alt+V のホットキー登録に失敗しました。\nWin32 Error: {error} ({message})\n\n他のアプリが同じキーを使用している可能性があります。",
                "BottomOverlay",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            Shutdown();
        }
    }

    protected override void OnExit(ExitEventArgs e)
    {
        if (_overlayHandle != IntPtr.Zero)
        {
            UnregisterHotKey(_overlayHandle, HotkeyId);
        }

        if (_overlaySource is not null)
        {
            _overlaySource.RemoveHook(WndProc);
        }

        _overlayWindow?.Close();

        base.OnExit(e);
    }

    private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        if (msg == WmHotkey && wParam.ToInt32() == HotkeyId)
        {
            _overlayWindow?.Toggle();
            handled = true;
        }

        return IntPtr.Zero;
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);
}
