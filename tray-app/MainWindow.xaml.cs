using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using MediaPipelineTray.ViewModels;

namespace MediaPipelineTray;

public partial class MainWindow : Window
{
    private readonly ShellViewModel _shell;

    /// <summary>
    /// Set by the tray before a real exit. Closing normally hides the window instead, so the app
    /// keeps watching in the background; Quit lives only in the tray menu.
    /// </summary>
    public bool AllowClose { get; set; }

    public MainWindow(ShellViewModel shell)
    {
        _shell = shell;

        InitializeComponent();
        DataContext = shell;
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

    private void OnTabChanged(object sender, SelectionChangedEventArgs e)
    {
        if (e.OriginalSource is TabControl { SelectedItem: TabItem { Header: string header } })
        {
            _shell.OnTabOpened(header);
        }
    }

    private void OnPauseAll(object sender, RoutedEventArgs e) => _shell.Activity.TogglePauseAll();

    private void OnOpenLogs(object sender, RoutedEventArgs e) => _shell.Activity.OpenLogsFolder();

    private void OnStart(object sender, RoutedEventArgs e) => _shell.Watcher.Start();

    private async void OnStop(object sender, RoutedEventArgs e)
    {
        var answer = MessageBox.Show(
            this,
            "Stop the watcher?\n\n"
            + "It finishes the file it is working on first, so this can take a moment on a long "
            + "encode. Nothing in the queue is lost.",
            "Stop watcher",
            MessageBoxButton.OKCancel,
            MessageBoxImage.Question);

        if (answer != MessageBoxResult.OK)
        {
            return;
        }

        await WithBusyButton(sender, async () =>
        {
            var stopped = await _shell.Watcher.StopAsync(TimeSpan.FromSeconds(120));

            if (!stopped)
            {
                MessageBox.Show(
                    this,
                    "The watcher has not stopped yet. It is still finishing the current file and "
                    + "will exit on its own. Nothing was killed.",
                    "Still stopping",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
        });
    }

    private async void OnRestart(object sender, RoutedEventArgs e) =>
        await WithBusyButton(sender, async () =>
        {
            var restarted = await _shell.Watcher.RestartAsync(TimeSpan.FromSeconds(120));

            if (!restarted)
            {
                MessageBox.Show(
                    this,
                    "The watcher did not come back within two minutes. It finishes the file it is "
                    + "on before exiting, so a long encode can take a while. Nothing was killed.",
                    "Restart timed out",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }
        });

    private static async Task WithBusyButton(object sender, Func<Task> work)
    {
        if (sender is not FrameworkElement button)
        {
            await work();
            return;
        }

        button.IsEnabled = false;

        try
        {
            await work();
        }
        finally
        {
            button.IsEnabled = true;
        }
    }
}
