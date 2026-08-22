using MediaPipelineTray.Models;

namespace MediaPipelineTray.Services;

/// <summary>Everything the UI needs for one render, produced by a single tick.</summary>
public sealed record PipelineSnapshot
{
    public bool WatcherRunning { get; init; }
    public bool PausedAll { get; init; }
    public WatcherStatus? Status { get; init; }

    public IReadOnlyList<JobProgress> Running { get; init; } = [];
    public IReadOnlyList<JobProgress> Finished { get; init; } = [];
    public IReadOnlyList<JobProgress> Failed { get; init; } = [];
    public IReadOnlyList<LaneInfo> Queued { get; init; } = [];
    public IReadOnlyList<LaneInfo> Idle { get; init; } = [];

    public int TotalQueued => Queued.Sum(lane => lane.Queued);
    public int OutputsToday { get; init; }

    /// <summary>
    /// What the tray icon should show. Failure outranks everything, because the whole point of
    /// tinting the icon is that a glance at the tray tells you something needs you.
    /// </summary>
    public ActivityState TrayState
    {
        get
        {
            if (!WatcherRunning) return ActivityState.Idle;
            if (Failed.Count > 0) return ActivityState.Failed;
            if (Running.Count > 0) return ActivityState.Running;
            if (PausedAll) return ActivityState.Paused;
            return ActivityState.Idle;
        }
    }
}

/// <summary>
/// Folds the watcher's two output surfaces into one snapshot.
///
/// Queue depth and preset configuration come from status\watcher.json, which the watcher
/// rewrites each sweep. Progress comes from the event stream, because the status file only
/// says how much is waiting, not how far along the current work is.
///
/// Jobs are assembled by grouping events on jobId. That is the whole reason the watcher emits
/// one: without it, concurrent jobs interleave in the stream with no way to tell them apart.
/// </summary>
public sealed class PipelineMonitor
{
    /// <summary>How many finished jobs to keep on screen. Older ones fall off.</summary>
    private const int FinishedHistory = 12;

    private readonly WatcherService _watcher;
    private readonly EventStreamReader _events;

    private readonly Dictionary<string, JobProgress> _active = new(StringComparer.Ordinal);
    private readonly LinkedList<JobProgress> _finished = new();
    private readonly Dictionary<string, JobProgress> _failed = new(StringComparer.Ordinal);

    private WatcherStatus? _lastStatus;
    private int _outputsToday;
    private DateOnly _outputsDay = DateOnly.FromDateTime(DateTime.Now);

    public PipelineMonitor(WatcherService watcher, EventStreamReader events)
    {
        _watcher = watcher;
        _events = events;
    }

    /// <summary>
    /// Replays today's events so a freshly opened window already knows what happened this
    /// morning, instead of looking idle until the next job starts.
    /// </summary>
    public void PrimeFromToday()
    {
        foreach (var pipelineEvent in _events.ReadNew())
        {
            Apply(pipelineEvent);
        }

        // Anything still "running" after a replay belongs to a watcher run that has since
        // ended, so it never got its completion event.
        if (!_watcher.IsRunning)
        {
            _active.Clear();
        }
    }

    public PipelineSnapshot Tick()
    {
        var running = _watcher.IsRunning;

        var status = _watcher.ReadStatus();
        if (status is not null)
        {
            _lastStatus = status;
        }

        foreach (var pipelineEvent in _events.ReadNew())
        {
            Apply(pipelineEvent);
        }

        if (!running)
        {
            // A stopped watcher cannot be mid-job. Keeping stale progress bars on screen would
            // be worse than showing nothing.
            _active.Clear();
        }

        var lanes = _lastStatus?.Lanes ?? [];

        return new PipelineSnapshot
        {
            WatcherRunning = running,
            PausedAll = _watcher.IsPausedAll,
            Status = _lastStatus,
            Running = [.. _active.Values.OrderBy(job => job.StartedUtc)],
            Finished = [.. _finished],
            Failed = [.. _failed.Values.OrderByDescending(job => job.EndedUtc)],
            Queued = [.. lanes.Where(lane => lane.Queued > 0).OrderByDescending(lane => lane.Queued)],
            Idle = [.. lanes.Where(lane => lane.Queued == 0)],
            OutputsToday = _outputsToday,
        };
    }

    /// <summary>Clears the failure banner. The files themselves are untouched.</summary>
    public void DismissFailures() => _failed.Clear();

    private void Apply(PipelineEvent pipelineEvent)
    {
        RollOverDayIfNeeded(pipelineEvent.Timestamp);

        switch (pipelineEvent.Name)
        {
            case "watcher.start":
                // A restart invalidates anything we thought was in flight.
                _active.Clear();
                break;

            case "watcher.stop":
                _active.Clear();
                break;

            case "job.start":
                StartJob(pipelineEvent);
                break;

            case "job.variant":
                UpdateJob(pipelineEvent);
                break;

            case "job.done":
                CompleteJob(pipelineEvent);
                break;

            case "job.failed":
                FailJob(pipelineEvent);
                break;
        }
    }

    private void RollOverDayIfNeeded(DateTimeOffset timestamp)
    {
        var day = DateOnly.FromDateTime(timestamp.ToLocalTime().DateTime);
        if (day == _outputsDay)
        {
            return;
        }

        _outputsDay = day;
        _outputsToday = 0;
    }

    private void StartJob(PipelineEvent pipelineEvent)
    {
        if (pipelineEvent.JobId is not { Length: > 0 } jobId)
        {
            return;
        }

        _active[jobId] = new JobProgress
        {
            JobId = jobId,
            Preset = pipelineEvent.Preset ?? "?",
            Workspace = pipelineEvent.Workspace ?? "?",
            StartedUtc = pipelineEvent.Timestamp,
            Files = pipelineEvent.Files ?? [],
            Bytes = pipelineEvent.Bytes ?? 0,
        };
    }

    private void UpdateJob(PipelineEvent pipelineEvent)
    {
        if (pipelineEvent.JobId is not { Length: > 0 } jobId)
        {
            return;
        }

        // A variant can arrive without its job.start when the app opened mid-job, so
        // reconstruct just enough to show progress rather than dropping it.
        if (!_active.TryGetValue(jobId, out var job))
        {
            job = new JobProgress
            {
                JobId = jobId,
                Preset = pipelineEvent.Preset ?? "?",
                Workspace = pipelineEvent.Workspace ?? "?",
                StartedUtc = pipelineEvent.Timestamp,
                Files = pipelineEvent.File is null ? [] : [pipelineEvent.File],
            };
            _active[jobId] = job;
        }

        job.VariantsDone = pipelineEvent.Index ?? job.VariantsDone;
        job.VariantsTotal = pipelineEvent.Total ?? job.VariantsTotal;

        if (job.Files.Count == 0 && pipelineEvent.File is not null)
        {
            job.Files = [pipelineEvent.File];
        }

        _outputsToday++;
    }

    private void CompleteJob(PipelineEvent pipelineEvent)
    {
        if (pipelineEvent.JobId is not { Length: > 0 } jobId ||
            !_active.Remove(jobId, out var job))
        {
            return;
        }

        job.State = ActivityState.Finished;
        job.EndedUtc = pipelineEvent.Timestamp;
        job.Outputs = pipelineEvent.Outputs ?? job.VariantsDone;

        _finished.AddFirst(job);
        while (_finished.Count > FinishedHistory)
        {
            _finished.RemoveLast();
        }
    }

    private void FailJob(PipelineEvent pipelineEvent)
    {
        if (pipelineEvent.JobId is not { Length: > 0 } jobId)
        {
            return;
        }

        if (!_active.Remove(jobId, out var job))
        {
            job = new JobProgress
            {
                JobId = jobId,
                Preset = pipelineEvent.Preset ?? "?",
                Workspace = pipelineEvent.Workspace ?? "?",
                StartedUtc = pipelineEvent.Timestamp,
            };
        }

        job.State = ActivityState.Failed;
        job.EndedUtc = pipelineEvent.Timestamp;
        job.Error = pipelineEvent.Error ?? "unknown error";

        _failed[jobId] = job;
    }
}
