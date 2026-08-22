using System.Windows;
using MediaPipelineTray.Models;
using MediaPipelineTray.Services;
using MediaPipelineTray.Tray;
using MediaPipelineTray.ViewModels;

namespace MediaPipelineTray;

public partial class App : Application
{
    private TrayIcon? _tray;
    private ShellViewModel? _shell;
    private MainWindow? _window;
    private MainViewModel? _viewModel;
    private WatcherService? _watcher;
    private ActivityState _lastNotifiedState = ActivityState.Idle;

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        Theme.Apply(Resources);

        var paths = PipelinePaths.Discover();
        _watcher = new WatcherService(paths);

        _shell = new ShellViewModel(paths, _watcher);
        _viewModel = _shell.Activity;
        _viewModel.TrayStateChanged += OnTrayStateChanged;

        _window = new MainWindow(_shell);

        _tray = new TrayIcon();
        _tray.OpenRequested += (_, _) => ShowWindow();
        _tray.PauseRequested += (_, _) => _viewModel.TogglePauseAll();
        _tray.QuitRequested += (_, _) => Quit();

        _shell.Start();
        RefreshTray();

        // Windows 11 hides new tray icons in the overflow, which makes a tray-first app look
        // like it never started. The registry entry only exists once the icon has been shown,
        // so ask again shortly after startup; a first run promotes itself for the next one.
        StartupService.ApplyFirstRunDefault();
        StartupService.PromoteTrayIcon();
        _ = PromoteShortlyAsync();

        // Launching straight to the tray would leave a first-time user with no idea the app
        // started, so the window opens once and hides from then on.
        ShowWindow();
    }

    private static async Task PromoteShortlyAsync()
    {
        await Task.Delay(TimeSpan.FromSeconds(3)).ConfigureAwait(false);
        StartupService.ApplyFirstRunDefault();
        StartupService.PromoteTrayIcon();
    }

    private void ShowWindow()
    {
        if (_window is null)
        {
            return;
        }

        _window.Show();

        if (_window.WindowState == WindowState.Minimized)
        {
            _window.WindowState = WindowState.Normal;
        }

        _window.Activate();
    }

    private void OnTrayStateChanged(object? sender, ActivityState state)
    {
        RefreshTray();

        // Notify only on the edge into a failure, so a lingering red icon does not keep
        // reannouncing itself every tick.
        if (state == ActivityState.Failed && _lastNotifiedState != ActivityState.Failed)
        {
            var failure = _viewModel?.Failures.FirstOrDefault();
            _tray?.Notify(
                "Processing failed",
                failure is null ? "A job failed." : $"{failure.Lane}: {failure.Detail}",
                isFailure: true);
        }

        _lastNotifiedState = state;
    }

    private void RefreshTray()
    {
        if (_tray is null || _viewModel is null)
        {
            return;
        }

        var lines = new List<string>();

        foreach (var job in _viewModel.Running.Take(5))
        {
            lines.Add($"{job.Lane}   {job.Counts}");
        }

        foreach (var failure in _viewModel.Failures.Take(3))
        {
            lines.Add($"{failure.Lane}   failed");
        }

        if (lines.Count == 0)
        {
            lines.Add(_viewModel.WatcherRunning ? "Nothing running" : "Watcher not running");
        }

        _tray.BuildMenu(_viewModel.PausedAll, lines);
        _tray.SetState(_viewModel.TrayState, $"Media Pipeline  ·  {_viewModel.StatusText}");

        // The window carries the same glyph in the same state, so the taskbar button and the
        // tray icon always agree.
        if (_window is not null)
        {
            _window.Icon = TrayGlyph.RenderForWindow(_viewModel.TrayState, Theme.IsDark);
        }
    }

    private void Quit()
    {
        // Quitting the tray app never stops the watcher: they are separate processes, and the
        // watcher is meant to keep running at logon whether this window is open or not.
        if (_window is not null)
        {
            _window.AllowClose = true;
            _window.Close();
        }

        _shell?.Dispose();
        _tray?.Dispose();

        Shutdown();
    }

    protected override void OnExit(ExitEventArgs e)
    {
        _shell?.Dispose();
        _tray?.Dispose();
        base.OnExit(e);
    }
}
