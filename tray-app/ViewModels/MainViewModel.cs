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

/// <summary>One row in the Running list.</summary>
public sealed class JobRow : Observable
{
    public JobRow(JobProgress job) => Update(job);

    public string JobId { get; private set; } = "";
    public string Lane { get; private set; } = "";
    public string PresetSummary { get; private set; } = "";
    public string Subject { get; private set; } = "";
    public string Counts { get; private set; } = "";
    public string Elapsed { get; private set; } = "";
    public string Remaining { get; private set; } = "";
    public string Percent { get; private set; } = "";
    public double Fraction { get; private set; }

    public void Update(JobProgress job, string? presetSummary = null)
    {
        JobId = job.JobId;
        Lane = job.Lane;
        Subject = job.Subject;
        Fraction = job.Fraction;

        if (presetSummary is not null)
        {
            PresetSummary = presetSummary;
        }

        Counts = job.VariantsTotal > 0
            ? $"{job.VariantsDone} / {job.VariantsTotal}"
            : "starting";

        Percent = job.VariantsTotal > 0
            ? (job.Fraction * 100).ToString("0", CultureInfo.InvariantCulture) + "%"
            : "";

        Elapsed = Format(job.Elapsed) + " elapsed";
        Remaining = job.Remaining is { } left ? Format(left) + " left" : "";

        Raise(nameof(JobId));
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

public sealed class FinishedRow
{
    public required string Lane { get; init; }
    public required string Detail { get; init; }
    public required string Outputs { get; init; }
    public required string When { get; init; }
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
    private readonly DispatcherTimer _timer;

    private string _statusText = "Checking";
    private string _statusMeta = "";
    private bool _watcherRunning;
    private bool _pausedAll;
    private ActivityState _trayState = ActivityState.Idle;

    public MainViewModel(PipelineMonitor monitor, WatcherService watcher, PipelinePaths paths)
    {
        _monitor = monitor;
        _watcher = watcher;
        _paths = paths;

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

        foreach (var job in snapshot.Running)
        {
            seen.Add(job.JobId);
            var existing = Running.FirstOrDefault(row => row.JobId == job.JobId);

            if (existing is null)
            {
                Running.Add(new JobRow(job) { });
                Running[^1].Update(job, SummaryFor(snapshot, job.Preset));
            }
            else
            {
                existing.Update(job, SummaryFor(snapshot, job.Preset));
            }
        }

        for (var i = Running.Count - 1; i >= 0; i--)
        {
            if (!seen.Contains(Running[i].JobId))
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

    private void SyncFinished(PipelineSnapshot snapshot)
    {
        Finished.Clear();

        foreach (var job in snapshot.Finished)
        {
            Finished.Add(new FinishedRow
            {
                Lane = job.Lane,
                Detail = job.Subject,
                Outputs = job.Outputs == 1 ? "1 output" : $"{job.Outputs} outputs",
                When = (job.EndedUtc ?? job.StartedUtc).ToLocalTime().ToString("HH:mm", CultureInfo.InvariantCulture),
            });
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
