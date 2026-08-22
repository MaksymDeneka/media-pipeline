using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Windows.Threading;
using MediaPipelineTray.Models;
using MediaPipelineTray.Services;

namespace MediaPipelineTray.ViewModels;

public abstract class Observable : INotifyPropertyChanged
{
    public event PropertyChangedEventHandler? PropertyChanged;

    protected void Raise([CallerMemberName] string? name = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

    protected bool Set<T>(ref T field, T value, [CallerMemberName] string? name = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value))
        {
            return false;
        }

        field = value;
        Raise(name);
        return true;
    }
}

/// <summary>
/// One row in the Running list, covering a whole lane rather than a single file. A preset
/// working through several files shows one bar, because that is the unit of work a person
/// cares about.
/// </summary>
public sealed class JobRow : Observable
{
    public JobRow(LaneProgress lane) => Update(lane);

    public string Key { get; private set; } = "";
    public string Lane { get; private set; } = "";
    public string PresetSummary { get; private set; } = "";
    public string Subject { get; private set; } = "";
    public string Counts { get; private set; } = "";
    public string Elapsed { get; private set; } = "";
    public string Remaining { get; private set; } = "";
    public string Percent { get; private set; } = "";
    public double Fraction { get; private set; }

    public void Update(LaneProgress lane, string? presetSummary = null)
    {
        Key = lane.Lane;
        Lane = lane.Lane;
        Subject = lane.Subject;
        Fraction = lane.Fraction;

        if (presetSummary is not null)
        {
            PresetSummary = presetSummary;
        }

        Counts = lane.VariantsTotal > 0
            ? $"{lane.VariantsDone} / {lane.VariantsTotal}"
            : "starting";

        Percent = lane.VariantsTotal > 0
            ? (lane.Fraction * 100).ToString("0", CultureInfo.InvariantCulture) + "%"
            : "";

        Elapsed = Format(lane.Elapsed) + " elapsed";
        Remaining = lane.Remaining is { } left ? Format(left) + " left" : "";

        Raise(nameof(Key));
        Raise(nameof(Lane));
        Raise(nameof(PresetSummary));
        Raise(nameof(Subject));
        Raise(nameof(Counts));
        Raise(nameof(Elapsed));
        Raise(nameof(Remaining));
        Raise(nameof(Percent));
        Raise(nameof(Fraction));
    }

    private static string Format(TimeSpan span) => span.TotalHours >= 1
        ? $"{(int)span.TotalHours}h {span.Minutes:00}m"
        : $"{span.Minutes}:{span.Seconds:00}";
}

public sealed class LaneRow
{
    public required string Lane { get; init; }
    public required string Detail { get; init; }
    public required string Preset { get; init; }
    public required string Workspace { get; init; }
    public bool Paused { get; init; }
}

public sealed class FinishedRow : Observable
{
    private bool _isBusy;

    public required string JobId { get; init; }
    public required string Lane { get; init; }
    public required string Detail { get; init; }
    public required string Outputs { get; init; }
    public required string When { get; init; }

    public required string Preset { get; init; }
    public required string Workspace { get; init; }
    public required string NameHint { get; init; }

    /// <summary>Output paths relative to the lane's output folder.</summary>
    public required IReadOnlyList<string> OutputPaths { get; init; }

    /// <summary>Zipping is only offered while the output is still recorded.</summary>
    public bool CanArchive => OutputPaths.Count > 0 && !IsBusy;

    public bool IsBusy
    {
        get => _isBusy;
        set
        {
            if (Set(ref _isBusy, value))
            {
                Raise(nameof(CanArchive));
            }
        }
    }
}

public sealed class FailureRow
{
    public required string Lane { get; init; }
    public required string Detail { get; init; }
    public required string Preset { get; init; }
    public required string Workspace { get; init; }
}

/// <summary>
/// Drives the Activity view. Polls on a timer rather than watching the filesystem, because the
/// watcher rewrites its status every sweep anyway and a timer is simpler to reason about than
/// coalescing change notifications.
/// </summary>
public sealed class MainViewModel : Observable, IDisposable
{
    private readonly PipelineMonitor _monitor;
    private readonly WatcherService _watcher;
    private readonly PipelinePaths _paths;
    private readonly ArchiveService _archives;
    private readonly DispatcherTimer _timer;

    private string _statusText = "Checking";
    private string _statusMeta = "";
    private bool _watcherRunning;
    private bool _pausedAll;
    private ActivityState _trayState = ActivityState.Idle;

    public MainViewModel(
        PipelineMonitor monitor,
        WatcherService watcher,
        PipelinePaths paths,
        ArchiveService archives)
    {
        _monitor = monitor;
        _watcher = watcher;
        _paths = paths;
        _archives = archives;

        _timer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1) };
        _timer.Tick += (_, _) => Refresh();
    }

    public ObservableCollection<JobRow> Running { get; } = [];
    public ObservableCollection<LaneRow> Queued { get; } = [];
    public ObservableCollection<FinishedRow> Finished { get; } = [];
    public ObservableCollection<FailureRow> Failures { get; } = [];

    public string StatusText { get => _statusText; private set => Set(ref _statusText, value); }
    public string StatusMeta { get => _statusMeta; private set => Set(ref _statusMeta, value); }
    public bool WatcherRunning { get => _watcherRunning; private set => Set(ref _watcherRunning, value); }
    public bool PausedAll { get => _pausedAll; private set => Set(ref _pausedAll, value); }

    public ActivityState TrayState
    {
        get => _trayState;
        private set
        {
            if (Set(ref _trayState, value))
            {
                TrayStateChanged?.Invoke(this, value);
            }
        }
    }

    public event EventHandler<ActivityState>? TrayStateChanged;

    public string PauseAllLabel => PausedAll ? "Resume all" : "Pause all";

    public void Start()
    {
        _monitor.PrimeFromToday();
        Refresh();
        _timer.Start();
    }

    public void Dispose() => _timer.Stop();

    public void TogglePauseAll()
    {
        _watcher.SetPauseAll(!_watcher.IsPausedAll);
        Refresh();
    }

    public void PauseLane(LaneRow lane)
    {
        _watcher.SetPauseLane(lane.Preset, lane.Workspace, !lane.Paused);
        Refresh();
    }

    public int RequeueFailures()
    {
        var moved = Failures
            .Select(failure => (failure.Preset, failure.Workspace))
            .Distinct()
            .Sum(lane => _watcher.RequeueFailed(lane.Preset, lane.Workspace));

        _monitor.DismissFailures();
        Refresh();
        return moved;
    }

    public void DismissFailures()
    {
        _monitor.DismissFailures();
        Refresh();
    }

    public void OpenLaneFolder(string preset, string workspace) =>
        WatcherService.OpenInExplorer(_paths.OutputDirectory(preset, workspace));

    public void OpenLogsFolder() => WatcherService.OpenInExplorer(_paths.LogsDirectory);

    /// <summary>
    /// Hands a finished zip to the upload queue. Set by the shell, because Activity should not
    /// need to know how uploading works.
    /// </summary>
    public Action<string>? QueueUpload { get; set; }

    /// <summary>
    /// Collects a finished job's output into a zip staged in that workspace's sync folder, and
    /// optionally queues it for upload straight away.
    /// </summary>
    public ArchiveResult ArchiveFinished(FinishedRow row, bool thenUpload)
    {
        row.IsBusy = true;

        try
        {
            var result = _archives.Create(row.Preset, row.Workspace, row.OutputPaths, row.NameHint);

            if (thenUpload)
            {
                QueueUpload?.Invoke(result.Path);
            }

            return result;
        }
        finally
        {
            row.IsBusy = false;
        }
    }

    private void Refresh()
    {
        var snapshot = _monitor.Tick();

        WatcherRunning = snapshot.WatcherRunning;
        PausedAll = snapshot.PausedAll;
        TrayState = snapshot.TrayState;
        Raise(nameof(PauseAllLabel));

        StatusText = snapshot switch
        {
            { WatcherRunning: false } => "Not running",
            { PausedAll: true } => "Paused",
            _ => "Running",
        };

        StatusMeta = BuildMeta(snapshot);

        SyncRunning(snapshot);
        SyncQueued(snapshot);
        SyncFinished(snapshot);
        SyncFailures(snapshot);
    }

    private string BuildMeta(PipelineSnapshot snapshot)
    {
        if (!snapshot.WatcherRunning)
        {
            return "The watcher is not running. Start it to begin processing.";
        }

        var parts = new List<string>();

        if (snapshot.Status is { } status)
        {
            var uptime = DateTimeOffset.UtcNow - status.StartedUtc;
            parts.Add(uptime.TotalHours >= 1
                ? $"{(int)uptime.TotalHours}h {uptime.Minutes:00}m"
                : $"{uptime.Minutes}m");

            if (status.Encoder.Length > 0)
            {
                parts.Add(status.Encoder);
            }
        }

        parts.Add($"{snapshot.Running.Count} active");
        parts.Add($"{snapshot.TotalQueued} queued");
        parts.Add($"{snapshot.OutputsToday} done today");

        return string.Join("  ·  ", parts);
    }

    private string SummaryFor(PipelineSnapshot snapshot, string preset) =>
        snapshot.Status?.Presets.FirstOrDefault(p => p.Name == preset)?.Summary ?? "";

    /// <summary>
    /// Updates rows in place where the job is already on screen, so a progress bar animates
    /// rather than the row being torn down and rebuilt every second.
    /// </summary>
    private void SyncRunning(PipelineSnapshot snapshot)
    {
        var seen = new HashSet<string>(StringComparer.Ordinal);

        foreach (var lane in snapshot.Running)
        {
            seen.Add(lane.Lane);
            var existing = Running.FirstOrDefault(row => row.Key == lane.Lane);

            if (existing is null)
            {
                Running.Add(new JobRow(lane));
                Running[^1].Update(lane, SummaryFor(snapshot, lane.Preset));
            }
            else
            {
                existing.Update(lane, SummaryFor(snapshot, lane.Preset));
            }
        }

        for (var i = Running.Count - 1; i >= 0; i--)
        {
            if (!seen.Contains(Running[i].Key))
            {
                Running.RemoveAt(i);
            }
        }
    }

    private void SyncQueued(PipelineSnapshot snapshot)
    {
        Queued.Clear();

        foreach (var lane in snapshot.Queued)
        {
            var paused = _watcher.IsPaused(lane.Preset, lane.Workspace);
            Queued.Add(new LaneRow
            {
                Lane = $"{lane.Preset} / {lane.Workspace}",
                Detail = paused
                    ? $"{lane.Queued} waiting, paused"
                    : $"{lane.Queued} waiting",
                Preset = lane.Preset,
                Workspace = lane.Workspace,
                Paused = paused,
            });
        }
    }

    /// <summary>
    /// Rebuilds only what changed. A row must survive a tick, or clicking Zip would lose its
    /// in-progress state a second later when the list refreshed underneath it.
    /// </summary>
    private void SyncFinished(PipelineSnapshot snapshot)
    {
        var seen = new HashSet<string>(StringComparer.Ordinal);

        foreach (var job in snapshot.Finished)
        {
            seen.Add(job.JobId);

            if (Finished.Any(row => row.JobId == job.JobId))
            {
                continue;
            }

            var row = new FinishedRow
            {
                JobId = job.JobId,
                Lane = job.Lane,
                Detail = job.Subject,
                Outputs = job.Outputs == 1 ? "1 output" : $"{job.Outputs} outputs",
                When = (job.EndedUtc ?? job.StartedUtc).ToLocalTime().ToString("HH:mm", CultureInfo.InvariantCulture),
                Preset = job.Preset,
                Workspace = job.Workspace,
                NameHint = job.Files.Count == 1 ? job.Files[0] : "",
                OutputPaths = [.. job.OutputPaths],
            };

            // Newest first, matching the order the monitor keeps them in.
            Finished.Insert(0, row);
        }

        for (var i = Finished.Count - 1; i >= 0; i--)
        {
            if (!seen.Contains(Finished[i].JobId))
            {
                Finished.RemoveAt(i);
            }
        }
    }

    private void SyncFailures(PipelineSnapshot snapshot)
    {
        Failures.Clear();

        foreach (var job in snapshot.Failed)
        {
            var when = (job.EndedUtc ?? job.StartedUtc).ToLocalTime().ToString("HH:mm:ss", CultureInfo.InvariantCulture);
            Failures.Add(new FailureRow
            {
                Lane = job.Lane,
                Detail = $"{job.Subject}  ·  {job.Error}  ·  {when}",
                Preset = job.Preset,
                Workspace = job.Workspace,
            });
        }
    }

}
