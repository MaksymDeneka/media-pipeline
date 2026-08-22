using System.Windows;
using MediaPipelineTray.Models;
using MediaPipelineTray.Services;
using MediaPipelineTray.Tray;
using MediaPipelineTray.ViewModels;

namespace MediaPipelineTray;

public partial class App : Application
{
    private TrayIcon? _tray;
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

        var events = new EventStreamReader(paths);
        var monitor = new PipelineMonitor(_watcher, events);

        _viewModel = new MainViewModel(monitor, _watcher, paths);
        _viewModel.TrayStateChanged += OnTrayStateChanged;

        _window = new MainWindow(_viewModel, _watcher);

        _tray = new TrayIcon();
        _tray.OpenRequested += (_, _) => ShowWindow();
        _tray.PauseRequested += (_, _) => _viewModel.TogglePauseAll();
        _tray.QuitRequested += (_, _) => Quit();

        _viewModel.Start();
        RefreshTray();

        // Launching straight to the tray would leave a first-time user with no idea the app
        // started, so the window opens once and hides from then on.
        ShowWindow();
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
        _tray.SetState(_viewModel.TrayState, $"Media Pipeline — {_viewModel.StatusText}");
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

        _viewModel?.Dispose();
        _tray?.Dispose();

        Shutdown();
    }

    protected override void OnExit(ExitEventArgs e)
    {
        _viewModel?.Dispose();
        _tray?.Dispose();
        base.OnExit(e);
    }
}
