using Microsoft.UI.Xaml;
using System.Diagnostics;

namespace PasteWinUI;

public partial class App : Application
{
    private MainWindow? _window;

    public App()
    {
        InitializeComponent();
    }

    protected override void OnLaunched(LaunchActivatedEventArgs args)
    {
        EnsureOnlyLatestInstanceRunning();

        _window = new MainWindow();
        _window.Activate();
        _window.InitializeOverlay();
    }

    private static void EnsureOnlyLatestInstanceRunning()
    {
        Process? current = null;
        try
        {
            current = Process.GetCurrentProcess();
            var processName = current.ProcessName;
            var siblings = Process.GetProcessesByName(processName);
            foreach (var process in siblings)
            {
                try
                {
                    if (process.Id == current.Id)
                    {
                        continue;
                    }

                    process.Kill(true);
                    process.WaitForExit(3000);
                }
                catch
                {
                    // Ignore inaccessible or already-terminating processes.
                }
                finally
                {
                    process.Dispose();
                }
            }
        }
        catch
        {
            // Best effort only.
        }
        finally
        {
            current?.Dispose();
        }
    }
}
