using MediaPipeline.Core.Configuration;
using MediaPipeline.Core.Contracts;
using MediaPipeline.Core.IO;
using MediaPipeline.Core.Media;

namespace MediaPipeline.Core.Runtime;

public sealed class PipelineWorker(
    PipelineConfiguration configuration,
    PipelinePaths paths,
    FfmpegEngine engine,
    PresetProcessor processor,
    WorkerControl control,
    StatusWriter statusWriter,
    EventWriter events,
    PipelineLogger logger,
    ArchiveManager archive)
{
    private readonly FileStabilityTracker _fileStability = new();
    private readonly Dictionary<string, BatchObservation> _batchObservations =
        new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, LaneStatus> _lanes = new(StringComparer.OrdinalIgnoreCase);
    private readonly DateTimeOffset _startedUtc = DateTimeOffset.UtcNow;

    public async Task RunAsync(CancellationToken cancellationToken = default)
    {
        InitializeDirectories();
        control.ClearStopRequest();
        await logger.InfoAsync(
            $"Worker started. Root: {paths.Root}. Encoder: {engine.Encoder.Name}.",
            cancellationToken);
        await events.AppendAsync(new PipelineEvent
        {
            Name = "watcher.start",
            Pid = Environment.ProcessId,
            PipelineRoot = paths.Root,
            Encoder = engine.Encoder.Name,
            Presets = configuration.Presets.Select(preset => preset.Name).ToArray(),
            Workspaces = configuration.Workspaces,
        }, cancellationToken);
        // Publish liveness before the first sweep. A ready long encode can make that sweep take
        // hours, and desktop clients must not mistake the owned root for a stopped worker.
        await WriteStatusAsync(cancellationToken);

        try
        {
            while (!cancellationToken.IsCancellationRequested && !control.StopRequested)
            {
                await RunOnceAsync(assumeStable: false, cancellationToken);
                if (control.StopRequested)
                {
                    break;
                }

                await Task.Delay(
                    TimeSpan.FromSeconds(configuration.Timing.PollSeconds),
                    cancellationToken);
            }
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
        }
        finally
        {
            await events.AppendAsync(new PipelineEvent
            {
                Name = "watcher.stop",
                Pid = Environment.ProcessId,
                Reason = control.StopRequested ? "control" : "cancellation",
            }, CancellationToken.None);
            control.ClearStopRequest();
            await WriteStatusAsync(CancellationToken.None);
            await logger.InfoAsync("Worker stopped cleanly.", CancellationToken.None);
        }
    }

    public async Task RunOnceAsync(
        bool assumeStable,
        CancellationToken cancellationToken = default)
    {
        InitializeDirectories();
        await archive.RunIfDueAsync(DateTimeOffset.Now, cancellationToken);

        foreach (var workspace in configuration.Workspaces)
        {
            foreach (var preset in configuration.Presets)
            {
                if (control.StopRequested)
                {
                    await WriteStatusAsync(cancellationToken);
                    return;
                }

                try
                {
                    await PollLaneAsync(preset, workspace, assumeStable, cancellationToken);
                }
                catch (Exception exception) when (exception is not OperationCanceledException)
                {
                    await logger.ErrorAsync(
                        $"Worker loop error [{workspace}/{preset.Name}]: {exception.Message}",
                        cancellationToken);
                }
            }
        }

        await WriteStatusAsync(cancellationToken);
    }

    private async Task PollLaneAsync(
        PresetOptions preset,
        string workspace,
        bool assumeStable,
        CancellationToken cancellationToken)
    {
        var lane = paths.Lane(preset.Name, workspace);
        var files = CandidateFiles(preset, lane);
        _fileStability.ForgetMissingFilesInDirectory(
            lane.Input,
            files.Select(file => file.FullName));
        var paused = control.IsPaused(preset.Name, workspace);
        _lanes[LaneKey(preset.Name, workspace)] = new LaneStatus
        {
            Preset = preset.Name,
            Workspace = workspace,
            Queued = files.Count,
            Paused = paused,
        };

        if (paused || files.Count == 0)
        {
            if (files.Count == 0)
            {
                _batchObservations.Remove(LaneKey(preset.Name, workspace));
            }

            return;
        }

        if (preset.Batch == BatchMode.PerGroup)
        {
            if (!assumeStable)
            {
                var allReady = true;
                foreach (var file in files)
                {
                    var state = _fileStability.Observe(
                        file.FullName,
                        TimeSpan.FromSeconds(configuration.Timing.StableSeconds),
                        TimeSpan.FromSeconds(configuration.Timing.TimeoutSeconds),
                        DateTimeOffset.Now);
                    if (state == StabilityState.TimedOut)
                    {
                        await RejectUnstableFileAsync(preset, lane, file, cancellationToken);
                        return;
                    }
                    if (state == StabilityState.Rejected)
                    {
                        return;
                    }
                    allReady &= state == StabilityState.Ready;
                }
                if (!allReady)
                {
                    return;
                }
            }
            if (assumeStable || BatchIsStable(preset, workspace, files, DateTimeOffset.Now))
            {
                _batchObservations.Remove(LaneKey(preset.Name, workspace));
                await processor.ProcessGroupAsync(preset, lane, files, cancellationToken);
                foreach (var file in files) { _fileStability.Forget(file.FullName); }
            }

            return;
        }

        var now = DateTimeOffset.Now;
        IReadOnlyList<FileInfo> ready;
        if (assumeStable)
        {
            ready = files;
        }
        else
        {
            var accepted = new List<FileInfo>();
            foreach (var file in files)
            {
                var stability = _fileStability.Observe(
                    file.FullName,
                    TimeSpan.FromSeconds(configuration.Timing.StableSeconds),
                    TimeSpan.FromSeconds(configuration.Timing.TimeoutSeconds),
                    now);
                if (stability == StabilityState.Ready)
                {
                    accepted.Add(file);
                }
                else if (stability == StabilityState.TimedOut)
                {
                    await RejectUnstableFileAsync(preset, lane, file, cancellationToken);
                }
            }
            ready = accepted;
        }

        if (ready.Count == 0)
        {
            return;
        }

        var canParallelize =
            preset.Parallel == ParallelMode.OverFiles &&
            preset.CopiesAlternate <= 0 &&
            ready.Count > 1 &&
            configuration.Images.ProcessingConcurrency > 1;
        if (!canParallelize)
        {
            foreach (var file in ready)
            {
                await processor.ProcessGroupAsync(preset, lane, [file], cancellationToken);
                _fileStability.Forget(file.FullName);
            }

            return;
        }

        using var throttle = new SemaphoreSlim(configuration.Images.ProcessingConcurrency);
        var tasks = ready.Select(async file =>
        {
            await throttle.WaitAsync(cancellationToken);
            try
            {
                await processor.ProcessGroupAsync(preset, lane, [file], cancellationToken);
            }
            finally
            {
                _fileStability.Forget(file.FullName);
                throttle.Release();
            }
        });
        await Task.WhenAll(tasks);
    }

    private static IReadOnlyList<FileInfo> CandidateFiles(PresetOptions preset, LanePaths lane)
    {
        if (!Directory.Exists(lane.Input))
        {
            return [];
        }

        // Admit every file of a known media kind regardless of what this preset
        // accepts: routing runs on probed content, and a file no preset accepts
        // must fail visibly into failed/ instead of looping in input/ forever.
        // Temporary and unknown extensions stay excluded, as before.
        return new DirectoryInfo(lane.Input)
            .EnumerateFiles()
            .Where(file => MediaClassifier.Classify(file.FullName) is MediaKind.Video or MediaKind.Image)
            .OrderBy(file => file.LastWriteTimeUtc)
            .ThenBy(file => file.FullName, StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    private bool BatchIsStable(
        PresetOptions preset,
        string workspace,
        IReadOnlyList<FileInfo> files,
        DateTimeOffset now)
    {
        if (files.Any(file => now - file.LastWriteTimeUtc <
            TimeSpan.FromSeconds(configuration.Timing.StableSeconds)))
        {
            return false;
        }

        var signature = string.Join(';', files.Select(file => $"{file.FullName}|{file.Length}"));
        var key = LaneKey(preset.Name, workspace);
        if (_batchObservations.TryGetValue(key, out var previous) && previous.Signature == signature)
        {
            return true;
        }

        _batchObservations[key] = new BatchObservation(signature);
        return false;
    }

    private async Task WriteStatusAsync(CancellationToken cancellationToken)
    {
        var status = new WorkerStatus
        {
            StartedUtc = _startedUtc,
            UpdatedUtc = DateTimeOffset.UtcNow,
            PipelineRoot = paths.Root,
            Encoder = engine.Encoder.Name,
            PollSeconds = configuration.Timing.PollSeconds,
            PausedAll = control.PausedAll,
            Workspaces = configuration.Workspaces,
            Presets = configuration.Presets.Select(preset => new PresetStatus
            {
                Name = preset.Name,
                VideoCopies = preset.VideoCopies,
                ImageCopies = preset.ImageCopies,
                Grouping = preset.Grouping,
                SetCount = preset.SetCount,
                Batch = preset.Batch,
                Segment = preset.Segment,
                Manifest = preset.Manifest,
                SizeCapMB = preset.SizeCapMB,
            }).ToArray(),
            Lanes = _lanes.Values.ToArray(),
        };
        await statusWriter.WriteAsync(status, cancellationToken);
    }

    private async Task RejectUnstableFileAsync(
        PresetOptions preset,
        LanePaths lane,
        FileInfo file,
        CancellationToken cancellationToken)
    {
        var jobId = Guid.NewGuid().ToString("n")[..8];
        const string error = "Input did not become stable before TimeoutSeconds elapsed.";
        long sourceBytes;
        try { sourceBytes = file.Exists ? file.Length : 0; }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            sourceBytes = 0;
        }
        await events.AppendAsync(new PipelineEvent
        {
            Name = "job.start",
            JobId = jobId,
            Preset = preset.Name,
            Workspace = lane.Workspace,
            Files = [file.Name],
            Bytes = sourceBytes,
        }, cancellationToken);
        await events.AppendAsync(new PipelineEvent
        {
            Name = "job.failed",
            JobId = jobId,
            Preset = preset.Name,
            Workspace = lane.Workspace,
            Error = error,
        }, cancellationToken);
        await logger.ErrorAsync(
            $"Preset '{preset.Name}' [{lane.Workspace}] rejected '{file.Name}': {error}",
            cancellationToken);
    }

    private void InitializeDirectories()
    {
        foreach (var directory in paths.RequiredDirectories(configuration))
        {
            Directory.CreateDirectory(directory);
        }
    }

    private static string LaneKey(string preset, string workspace) => $"{preset}/{workspace}";

    private sealed record BatchObservation(string Signature);
}
