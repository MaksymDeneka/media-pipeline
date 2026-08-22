using System.ComponentModel;
using System.Windows;
using MediaPipelineTray.Services;
using MediaPipelineTray.ViewModels;

namespace MediaPipelineTray;

public partial class MainWindow : Window
{
    private readonly MainViewModel _viewModel;
    private readonly WatcherService _watcher;

    /// <summary>
    /// Set by the tray before a real exit. Closing the window normally hides it instead, so the
    /// app keeps watching in the background; Quit lives only in the tray menu.
    /// </summary>
    public bool AllowClose { get; set; }

    public MainWindow(MainViewModel viewModel, WatcherService watcher)
    {
        _viewModel = viewModel;
        _watcher = watcher;

        InitializeComponent();
        DataContext = viewModel;
    }

    protected override void OnClosing(CancelEventArgs e)
    {
        if (!AllowClose)
        {
            e.Cancel = true;
            Hide();
            return;
        }

        base.OnClosing(e);
    }

    private void OnPauseAll(object sender, RoutedEventArgs e) => _viewModel.TogglePauseAll();

    private async void OnRestart(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement button)
        {
            return;
        }

        button.IsEnabled = false;

        try
        {
            var restarted = await _watcher.RestartAsync(TimeSpan.FromSeconds(90));

            if (!restarted)
            {
                MessageBox.Show(
                    this,
                    "The watcher did not stop within 90 seconds. It finishes the file it is on "
                    + "before exiting, so a long encode can take a while. Nothing was killed.",
                    "Restart timed out",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }
        }
        finally
        {
            button.IsEnabled = true;
        }
    }

    private void OnOpenLogs(object sender, RoutedEventArgs e) => _viewModel.OpenLogsFolder();

    private void OnPauseLane(object sender, RoutedEventArgs e)
    {
        if (sender is FrameworkElement { DataContext: LaneRow lane })
        {
            _viewModel.PauseLane(lane);
        }
    }

    private void OnOpenFailedFolder(object sender, RoutedEventArgs e)
    {
        if (sender is FrameworkElement { DataContext: FailureRow failure })
        {
            _viewModel.OpenLaneFolder(failure.Preset, failure.Workspace);
        }
    }

    private void OnRequeue(object sender, RoutedEventArgs e)
    {
        var moved = _viewModel.RequeueFailures();

        MessageBox.Show(
            this,
            moved == 0
                ? "Nothing to requeue. The failed folder is empty."
                : $"Moved {moved} file(s) back to the input folder.",
            "Requeue",
            MessageBoxButton.OK,
            MessageBoxImage.Information);
    }

    private void OnDismissFailures(object sender, RoutedEventArgs e) => _viewModel.DismissFailures();
}
