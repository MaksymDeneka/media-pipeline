using System.Diagnostics;
using System.IO;
using System.Text.Json;
using System.Threading;
using MediaPipelineTray.Models;

namespace MediaPipelineTray.Services;

/// <summary>
/// Everything the app does *to* the watcher, plus reading its status.
///
/// Liveness comes from the single-instance mutex rather than the status file or a process
/// scan: the watcher holds it for its whole lifetime, so opening it is both cheap and
/// impossible to fool with a stale file.
/// </summary>
public sealed class WatcherService
{
    private static readonly JsonSerializerOptions Options = new()
    {
        PropertyNameCaseInsensitive = true,
    };

    private readonly PipelinePaths _paths;

    public WatcherService(PipelinePaths paths) => _paths = paths;

    public bool IsRunning
    {
        get
        {
            try
            {
                using var mutex = Mutex.OpenExisting(_paths.MutexName);
                return true;
            }
            catch (WaitHandleCannotBeOpenedException)
            {
                return false;
            }
            catch (UnauthorizedAccessException)
            {
                // It exists but this user cannot open it, which still means running.
                return true;
            }
        }
    }

    public WatcherStatus? ReadStatus()
    {
        if (!File.Exists(_paths.StatusFile))
        {
            return null;
        }

        try
        {
            using var stream = new FileStream(
                _paths.StatusFile, FileMode.Open, FileAccess.Read,
                FileShare.ReadWrite | FileShare.Delete);

            return JsonSerializer.Deserialize<WatcherStatus>(stream, Options);
        }
        catch (Exception ex) when (ex is IOException or JsonException)
        {
            // The watcher writes to a temp file and moves it, so a torn read is unlikely,
            // but a miss just means we keep the previous snapshot for one more tick.
            return null;
        }
    }

    // --- control flags -----------------------------------------------------

    private string ControlFile(string name) => Path.Combine(_paths.ControlDirectory, name);

    private void SetFlag(string name, bool on)
    {
        Directory.CreateDirectory(_paths.ControlDirectory);
        var file = ControlFile(name);

        if (on)
        {
            if (!File.Exists(file))
            {
                File.WriteAllText(file, "");
            }
        }
        else if (File.Exists(file))
        {
            File.Delete(file);
        }
    }

    public bool IsPausedAll => File.Exists(ControlFile("pause"));

    public bool IsPaused(string preset, string? workspace = null) =>
        IsPausedAll
        || File.Exists(ControlFile($"pause.{preset}"))
        || (workspace is not null && File.Exists(ControlFile($"pause.{preset}.{workspace}")));

    public void SetPauseAll(bool paused) => SetFlag("pause", paused);

    public void SetPausePreset(string preset, bool paused) => SetFlag($"pause.{preset}", paused);

    public void SetPauseLane(string preset, string workspace, bool paused) =>
        SetFlag($"pause.{preset}.{workspace}", paused);

    /// <summary>
    /// Asks the watcher to finish the file it is on and exit. Preferred over killing it, which
    /// orphans FFmpeg and strands the input file with no failed-move.
    /// </summary>
    public void RequestStop() => SetFlag("stop", true);

    // --- lifecycle ---------------------------------------------------------

    private const string ScheduledTaskName = "Media Pipeline Watcher";

    /// <summary>
    /// Starts the watcher through its scheduled task, which is what the installer registers.
    /// Going through the task rather than spawning a child keeps the watcher detached, so it
    /// survives this app closing.
    /// </summary>
    public void Start() => RunHidden("schtasks.exe", $"/Run /TN \"{ScheduledTaskName}\"");

    /// <summary>
    /// Stops gracefully, waits, and only then reports whether it actually went. Callers decide
    /// what to do if it did not; this never escalates to a kill on its own.
    /// </summary>
    public async Task<bool> StopAsync(TimeSpan timeout, CancellationToken cancellationToken = default)
    {
        if (!IsRunning)
        {
            return true;
        }

        RequestStop();

        var deadline = DateTimeOffset.UtcNow + timeout;
        while (DateTimeOffset.UtcNow < deadline)
        {
            if (!IsRunning)
            {
                return true;
            }

            await Task.Delay(250, cancellationToken).ConfigureAwait(false);
        }

        return !IsRunning;
    }

    public async Task<bool> RestartAsync(TimeSpan timeout, CancellationToken cancellationToken = default)
    {
        var stopped = await StopAsync(timeout, cancellationToken).ConfigureAwait(false);
        if (!stopped)
        {
            return false;
        }

        Start();

        var deadline = DateTimeOffset.UtcNow + TimeSpan.FromSeconds(30);
        while (DateTimeOffset.UtcNow < deadline)
        {
            if (IsRunning)
            {
                return true;
            }

            await Task.Delay(250, cancellationToken).ConfigureAwait(false);
        }

        return IsRunning;
    }

    // --- queue helpers -----------------------------------------------------

    /// <summary>
    /// Moves everything in a lane's failed folder back to its input folder. This needs no
    /// cooperation from the watcher: the input folder is the queue.
    /// </summary>
    public int RequeueFailed(string preset, string workspace)
    {
        var failed = _paths.FailedDirectory(preset, workspace);
        var input = _paths.InputDirectory(preset, workspace);

        if (!Directory.Exists(failed))
        {
            return 0;
        }

        Directory.CreateDirectory(input);
        var moved = 0;

        foreach (var file in Directory.GetFiles(failed))
        {
            var target = Path.Combine(input, Path.GetFileName(file));

            // Never clobber something already queued under the same name.
            if (File.Exists(target))
            {
                var name = Path.GetFileNameWithoutExtension(file);
                var extension = Path.GetExtension(file);
                target = Path.Combine(input, $"{name}-retry-{DateTime.Now:HHmmss}{extension}");
            }

            try
            {
                File.Move(file, target);
                moved++;
            }
            catch (IOException)
            {
                // Locked file; leave it for the next attempt.
            }
        }

        return moved;
    }

    public static void OpenInExplorer(string path)
    {
        if (!Directory.Exists(path) && !File.Exists(path))
        {
            return;
        }

        Process.Start(new ProcessStartInfo
        {
            FileName = "explorer.exe",
            Arguments = $"\"{path}\"",
            UseShellExecute = true,
        });
    }

    private static void RunHidden(string fileName, string arguments)
    {
        Process.Start(new ProcessStartInfo
        {
            FileName = fileName,
            Arguments = arguments,
            UseShellExecute = false,
            CreateNoWindow = true,
        });
    }
}
