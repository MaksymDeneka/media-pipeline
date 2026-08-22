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

public sealed class UploadRow : Observable
{
    private string _phase = "";
    private string _detail = "";
    private string _counts = "";
    private double _fraction;
    private bool _canCancel;

    public required UploadJob Job { get; init; }

    public string FileName => Job.FileName;
    public ObservableCollection<ChunkTile> Chunks { get; } = [];

    public string Phase { get => _phase; set => Set(ref _phase, value); }
    public string Detail { get => _detail; set => Set(ref _detail, value); }
    public string Counts { get => _counts; set => Set(ref _counts, value); }
    public double Fraction { get => _fraction; set => Set(ref _fraction, value); }
    public bool CanCancel { get => _canCancel; set => Set(ref _canCancel, value); }

    public void Sync()
    {
        // Chunks only appear once the split has planned them.
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
            UploadPhase.Queued => "Queued",
            UploadPhase.Splitting => "Splitting",
            UploadPhase.Sending => "Sending",
            UploadPhase.Assembling => "Assembling on the remote",
            UploadPhase.Verifying => "Verifying",
            UploadPhase.Done => "Done",
            UploadPhase.Failed => "Failed",
            UploadPhase.Cancelled => "Cancelled",
            UploadPhase.Paused => "Paused",
            _ => "",
        };

        CanCancel = Job.Phase is UploadPhase.Queued or UploadPhase.Splitting
            or UploadPhase.Sending or UploadPhase.Assembling;

        Fraction = Job.Fraction;

        Counts = Job.Chunks.Count == 0
            ? ""
            : $"{Job.ChunksSent} / {Job.Chunks.Count} chunks";

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
/// Uploads run one file at a time. Chunks within a file already go in parallel, and running two
/// large files at once on a link that drops under load would make both slower and less likely
/// to finish.
/// </summary>
public sealed class UploadsViewModel : Observable
{
    private readonly UploadService _service;
    private readonly PipelinePaths _paths;
    private readonly DispatcherTimer _timer;

    private CancellationTokenSource? _cancellation;
    private UploadRow? _current;
    private bool _isBusy;
    private string _status = "";

    public UploadsViewModel(UploadService service, PipelinePaths paths)
    {
        _service = service;
        _paths = paths;

        _service.Progress += OnProgress;

        // The service reports from worker threads; repaint on the UI thread on a timer rather
        // than marshalling every individual chunk update.
        _timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(400) };
        _timer.Tick += (_, _) => SyncRows();
        _timer.Start();
    }

    public ObservableCollection<UploadRow> Uploads { get; } = [];

    public bool IsBusy { get => _isBusy; private set => Set(ref _isBusy, value); }

    public string Status { get => _status; private set => Set(ref _status, value); }

    /// <summary>The folder uploads are picked from, which is what the sync scripts also use.</summary>
    public string SyncFolder => Path.Combine(_paths.PipelineRoot, "sync");

    private void OnProgress(object? sender, UploadJob job)
    {
        // Deliberately empty: the timer drives repainting. This keeps worker threads out of
        // the dispatcher entirely.
    }

    private void SyncRows()
    {
        foreach (var row in Uploads)
        {
            row.Sync();
        }

        if (_current is not null && !_current.CanCancel)
        {
            IsBusy = false;
            _current = null;
        }
    }

    /// <summary>Lists candidates from the sync folder, largest first, skipping work in progress.</summary>
    public IReadOnlyList<FileInfo> FindCandidates()
    {
        if (!Directory.Exists(SyncFolder))
        {
            return [];
        }

        return
        [
            .. new DirectoryInfo(SyncFolder)
                .GetFiles()
                .Where(f => f.Length > 0)
                .Where(f => !f.Name.EndsWith(".chunked.tmp", StringComparison.OrdinalIgnoreCase))
                .Where(f => !f.Name.EndsWith(".rclone-partial", StringComparison.OrdinalIgnoreCase))
                .OrderByDescending(f => f.Length)
        ];
    }

    public async Task StartAsync(string path)
    {
        if (IsBusy)
        {
            return;
        }

        var ini = IniFile.Load(_paths.ConfigFile);
        var target = UploadTarget.FromConfig(ini.ReadGlobals());

        var job = new UploadJob { SourcePath = path, Target = target };
        var row = new UploadRow { Job = job };

        Uploads.Insert(0, row);
        _current = row;
        IsBusy = true;
        Status = $"Uploading {job.FileName}";

        _cancellation = new CancellationTokenSource();

        try
        {
            await _service.RunAsync(job, _cancellation.Token).ConfigureAwait(true);
        }
        finally
        {
            _cancellation.Dispose();
            _cancellation = null;
            IsBusy = false;
            row.Sync();

            Status = job.Phase switch
            {
                UploadPhase.Done => $"{job.FileName} uploaded",
                UploadPhase.Cancelled => $"{job.FileName} cancelled. Local parts kept, so it resumes where it stopped.",
                UploadPhase.Failed => $"{job.FileName} failed: {job.Error}",
                _ => "",
            };
        }
    }

    /// <summary>
    /// Cancels the current upload. Local parts are kept deliberately, so starting the same file
    /// again skips everything already split and sent.
    /// </summary>
    public void Cancel() => _cancellation?.Cancel();

    public void Dispose()
    {
        _timer.Stop();
        _service.Progress -= OnProgress;
    }
}
