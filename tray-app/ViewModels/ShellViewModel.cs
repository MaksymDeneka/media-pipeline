using MediaPipelineTray.Services;

namespace MediaPipelineTray.ViewModels;

/// <summary>Holds the four tab view models and the shared services behind them.</summary>
public sealed class ShellViewModel : IDisposable
{
    public ShellViewModel(PipelinePaths paths, WatcherService watcher)
    {
        Paths = paths;
        Watcher = watcher;

        var events = new EventStreamReader(paths);
        var monitor = new PipelineMonitor(watcher, events);

        Activity = new MainViewModel(monitor, watcher, paths);
        Presets = new PresetsViewModel(paths, watcher);
        Settings = new SettingsViewModel(paths, watcher);
        Uploads = new UploadsViewModel(new UploadService(paths), paths);
    }

    public PipelinePaths Paths { get; }
    public WatcherService Watcher { get; }

    public MainViewModel Activity { get; }
    public PresetsViewModel Presets { get; }
    public SettingsViewModel Settings { get; }
    public UploadsViewModel Uploads { get; }

    public void Start()
    {
        Activity.Start();
        Presets.Load();
        Settings.Load();
    }

    /// <summary>
    /// Re-reads config.ini when a config tab is opened, so the view never shows something the
    /// file no longer says. Skipped while edits are pending, which would discard them.
    /// </summary>
    public void OnTabOpened(string header)
    {
        switch (header)
        {
            case "Presets" when !Presets.IsDirty:
                Presets.Load();
                break;

            case "Settings" when !Settings.IsDirty:
                Settings.Load();
                break;
        }
    }

    public void Dispose()
    {
        Activity.Dispose();
        Uploads.Dispose();
    }
}
