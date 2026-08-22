using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Windows.Threading;
using MediaPipelineTray.Services;

namespace MediaPipelineTray.ViewModels;

/// <summary>One chunk, as a square in the grid.</summary>
public sealed class ChunkTile : Observable
{
    private ChunkState _state;

    public required int Index { get; init; }

    public ChunkState State
    {
        get => _state;
        set
        {
            if (Set(ref _state, value))
            {
                Raise(nameof(IsSent));
                Raise(nameof(IsSending));
                Raise(nameof(IsFailed));
            }
        }
    }

    public bool IsSent => State == ChunkState.Sent;
    public bool IsSending => State is ChunkState.Sending or ChunkState.Splitting;
    public bool IsFailed => State == ChunkState.Failed;
}

/// <summary>A file staged in a workspace's sync folder, ready to upload.</summary>
public sealed class SyncFileRow : Observable
{
    private bool _isQueued;

    public required string FullPath { get; init; }
    public required string Name { get; init; }
    public required long Length { get; init; }
    public required string Workspace { get; init; }

    public string Size => UploadRow.Format(Length);

    /// <summary>True once queued, so the same file cannot be added twice.</summary>
    public bool IsQueued { get => _isQueued; set => Set(ref _isQueued, value); }
}

/// <summary>A workspace's sync folder, with the files waiting in it.</summary>
public sealed class SyncWorkspaceRow : Observable
{
    private bool _isExpanded;

    public required string Name { get; init; }
    public required string Path { get; init; }
    public ObservableCollection<SyncFileRow> Files { get; } = [];

    public bool IsExpanded { get => _isExpanded; set => Set(ref _isExpanded, value); }

    public string Summary => Files.Count switch
    {
        0 => "empty",
        1 => $"1 file, {UploadRow.Format(Files.Sum(f => f.Length))}",
        _ => $"{Files.Count} files, {UploadRow.Format(Files.Sum(f => f.Length))}",
    };

    public bool HasFiles => Files.Count > 0;

    public void Refreshed()
    {
        Raise(nameof(Summary));
        Raise(nameof(HasFiles));
    }
}

public sealed class UploadRow : Observable
{
    private string _phase = "";
    private string _detail = "";
    private string _counts = "";
    private double _fraction;
    private bool _canCancel;

    public required UploadJob Job { get; init; }

    public string FileName => Job.FileName;
    public string Workspace => Job.Workspace;
    public ObservableCollection<ChunkTile> Chunks { get; } = [];

    public string Phase { get => _phase; set => Set(ref _phase, value); }
    public string Detail { get => _detail; set => Set(ref _detail, value); }
    public string Counts { get => _counts; set => Set(ref _counts, value); }
    public double Fraction { get => _fraction; set => Set(ref _fraction, value); }
    public bool CanCancel { get => _canCancel; set => Set(ref _canCancel, value); }

    public void Sync()
    {
        while (Chunks.Count < Job.Chunks.Count)
        {
            Chunks.Add(new ChunkTile { Index = Chunks.Count + 1 });
        }

        for (var i = 0; i < Job.Chunks.Count; i++)
        {
            Chunks[i].State = Job.Chunks[i].State;
        }

        Phase = Job.Phase switch
        {
            UploadPhase.Queued => "Waiting",
            UploadPhase.Splitting => "Splitting",
            UploadPhase.Sending => "Sending",
            UploadPhase.Assembling => "Assembling on the remote",
            UploadPhase.Verifying => "Verifying",
            UploadPhase.Done => Job.SourceDeleted ? "Done, local copy deleted" : "Done",
            UploadPhase.Failed => "Failed",
            UploadPhase.Cancelled => "Cancelled",
            UploadPhase.Paused => "Paused",
            _ => "",
        };

        CanCancel = Job.Phase is UploadPhase.Queued or UploadPhase.Splitting
            or UploadPhase.Sending or UploadPhase.Assembling or UploadPhase.Verifying;

        Fraction = Job.Fraction;

        Counts = Job.Chunks.Count == 0 ? "" : $"{Job.ChunksSent} / {Job.Chunks.Count} chunks";

        var failedChunk = Job.Chunks.FirstOrDefault(c => c.State == ChunkState.Failed);
        var retrying = Job.Chunks.FirstOrDefault(c => c.State == ChunkState.Sending && c.Attempts > 1);

        Detail = Job.Error is not null
            ? Job.Error
            : failedChunk is not null
                ? $"chunk {failedChunk.Index} failed: {failedChunk.Error}"
                : retrying is not null
                    ? $"chunk {retrying.Index} retrying, attempt {retrying.Attempts} of 5"
                    : $"{Format(Job.BytesSent)} of {Format(Job.TotalBytes)}";
    }

    public static string Format(double bytes) => bytes switch
    {
        >= 1024d * 1024 * 1024 => (bytes / (1024d * 1024 * 1024)).ToString("0.0", CultureInfo.InvariantCulture) + " GB",
        >= 1024d * 1024 => (bytes / (1024d * 1024)).ToString("0", CultureInfo.InvariantCulture) + " MB",
        >= 1024 => (bytes / 1024).ToString("0", CultureInfo.InvariantCulture) + " KB",
        _ => bytes.ToString("0", CultureInfo.InvariantCulture) + " B",
    };
}

/// <summary>
/// The upload queue.
///
/// Files can be queued one at a time or a whole workspace at once, but they upload
/// sequentially. Chunks within a file already go in parallel, and running two large files at
/// once on a link that drops under load would make both slower and less likely to finish.
/// </summary>
public sealed class UploadsViewModel : Observable
{
    private readonly UploadService _service;
    private readonly PipelinePaths _paths;
    private readonly DispatcherTimer _timer;
    private readonly Queue<SyncFileRow> _pending = new();

    private CancellationTokenSource? _cancellation;
    private bool _isBusy;
    private string _status = "Nothing queued.";
    private bool _deleteAfterUpload;

    public UploadsViewModel(UploadService service, PipelinePaths paths)
    {
        _service = service;
        _paths = paths;

        _timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(400) };
        _timer.Tick += (_, _) => SyncRows();
        _timer.Start();
    }

    public ObservableCollection<SyncWorkspaceRow> Workspaces { get; } = [];
    public ObservableCollection<UploadRow> Uploads { get; } = [];

    public bool IsBusy { get => _isBusy; private set => Set(ref _isBusy, value); }
    public string Status { get => _status; private set => Set(ref _status, value); }

    /// <summary>Mirrors the config setting, so the tab says what will happen without a detour.</summary>
    public bool DeleteAfterUpload
    {
        get => _deleteAfterUpload;
        private set
        {
            if (Set(ref _deleteAfterUpload, value))
            {
                Raise(nameof(DeleteNote));
            }
        }
    }

    public string DeleteNote => DeleteAfterUpload
        ? "Local files are deleted once the remote copy is verified."
        : "Local files are kept after upload.";

    public int QueuedCount => _pending.Count;

    /// <summary>Rescans every workspace's sync folder.</summary>
    public void Refresh()
    {
        var workspaces = ReadWorkspaceNames();
        var queued = _pending.Select(f => f.FullPath).ToHashSet(StringComparer.OrdinalIgnoreCase);

        DeleteAfterUpload = UploadTarget
            .FromConfig(IniFile.Load(_paths.ConfigFile).ReadGlobals())
            .DeleteAfterUpload;

        Workspaces.Clear();

        foreach (var workspace in workspaces)
        {
            var directory = _paths.SyncDirectory(workspace);
            var row = new SyncWorkspaceRow { Name = workspace, Path = directory };

            if (Directory.Exists(directory))
            {
                var files = new DirectoryInfo(directory)
                    .GetFiles("*", SearchOption.TopDirectoryOnly)
                    .Where(f => f.Length > 0)
                    .Where(f => !f.Name.EndsWith(".chunked.tmp", StringComparison.OrdinalIgnoreCase))
                    .Where(f => !f.Name.EndsWith(".rclone-partial", StringComparison.OrdinalIgnoreCase))
                    .OrderByDescending(f => f.Length);

                foreach (var file in files)
                {
                    row.Files.Add(new SyncFileRow
                    {
                        FullPath = file.FullName,
                        Name = file.Name,
                        Length = file.Length,
                        Workspace = workspace,
                        IsQueued = queued.Contains(file.FullName),
                    });
                }
            }

            row.Refreshed();
            Workspaces.Add(row);
        }

        Raise(nameof(QueuedCount));
    }

    /// <summary>
    /// Workspace names come from the pipeline root rather than a hardcoded list, so a workspace
    /// added to the watcher shows up here without another edit.
    /// </summary>
    private IReadOnlyList<string> ReadWorkspaceNames()
    {
        if (!Directory.Exists(_paths.PipelineRoot))
        {
            return [];
        }

        var reserved = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "logs", "status", "control", "archive", ".sync-parts", "sync",
        };

        return
        [
            .. new DirectoryInfo(_paths.PipelineRoot)
                .GetDirectories()
                .Select(d => d.Name)
                .Where(name => !reserved.Contains(name))
                .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
        ];
    }

    public void QueueFile(SyncFileRow file)
    {
        if (file.IsQueued)
        {
            return;
        }

        file.IsQueued = true;
        _pending.Enqueue(file);
        Raise(nameof(QueuedCount));

        _ = PumpAsync();
    }

    public void QueueWorkspace(SyncWorkspaceRow workspace)
    {
        foreach (var file in workspace.Files.Where(f => !f.IsQueued).ToList())
        {
            file.IsQueued = true;
            _pending.Enqueue(file);
        }

        Raise(nameof(QueuedCount));
        _ = PumpAsync();
    }

    /// <summary>Runs the queue one file at a time until it is empty.</summary>
    private async Task PumpAsync()
    {
        if (IsBusy)
        {
            return;
        }

        IsBusy = true;

        try
        {
            while (_pending.Count > 0)
            {
                var next = _pending.Dequeue();
                Raise(nameof(QueuedCount));

                if (!File.Exists(next.FullPath))
                {
                    Status = $"{next.Name} is gone, skipping.";
                    continue;
                }

                await RunOneAsync(next).ConfigureAwait(true);
            }
        }
        finally
        {
            IsBusy = false;
            Refresh();
        }
    }

    private async Task RunOneAsync(SyncFileRow file)
    {
        var target = UploadTarget.FromConfig(IniFile.Load(_paths.ConfigFile).ReadGlobals());

        var job = new UploadJob
        {
            SourcePath = file.FullPath,
            Target = target,
            WorkspaceOverride = file.Workspace,
        };

        var row = new UploadRow { Job = job };
        Uploads.Insert(0, row);

        Status = $"Uploading {job.FileName} to {job.Workspace}";
        _cancellation = new CancellationTokenSource();

        try
        {
            await _service.RunAsync(job, _cancellation.Token).ConfigureAwait(true);
        }
        finally
        {
            _cancellation.Dispose();
            _cancellation = null;
            row.Sync();

            Status = job.Phase switch
            {
                UploadPhase.Done when job.SourceDeleted => $"{job.FileName} uploaded and removed locally",
                UploadPhase.Done => $"{job.FileName} uploaded",
                UploadPhase.Cancelled => $"{job.FileName} cancelled. Local parts kept, so it resumes where it stopped.",
                UploadPhase.Failed => $"{job.FileName} failed: {job.Error}",
                _ => "",
            };
        }
    }

    private void SyncRows()
    {
        foreach (var row in Uploads)
        {
            row.Sync();
        }
    }

    /// <summary>
    /// Cancels the file in flight. Local parts are kept deliberately, so starting the same file
    /// again skips everything already split and sent. Anything still queued is dropped.
    /// </summary>
    public void Cancel()
    {
        _pending.Clear();
        Raise(nameof(QueuedCount));
        _cancellation?.Cancel();
    }

    public void Dispose() => _timer.Stop();
}
