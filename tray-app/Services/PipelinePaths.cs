using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace MediaPipelineTray.Services;

/// <summary>
/// Everything the app needs to find on disk, derived from the pipeline root.
///
/// The root comes from config.ini next to watch-media.ps1, which is the same rule the watcher
/// itself uses, so the two can never disagree about where the pipeline lives.
/// </summary>
public sealed class PipelinePaths
{
    public PipelinePaths(string appDirectory, string pipelineRoot)
    {
        AppDirectory = appDirectory;
        PipelineRoot = pipelineRoot;
    }

    /// <summary>Where watch-media.ps1 and config.ini live.</summary>
    public string AppDirectory { get; }

    public string PipelineRoot { get; }

    public string ConfigFile => Path.Combine(AppDirectory, "config.ini");
    public string WatcherScript => Path.Combine(AppDirectory, "watch-media.ps1");

    public string LogsDirectory => Path.Combine(PipelineRoot, "logs");
    public string StatusFile => Path.Combine(PipelineRoot, "status", "watcher.json");
    public string ControlDirectory => Path.Combine(PipelineRoot, "control");

    public string EventFileFor(DateTimeOffset day) =>
        Path.Combine(LogsDirectory, $"events-{day.ToLocalTime():yyyyMMdd}.jsonl");

    public string LogFileFor(DateTimeOffset day) =>
        Path.Combine(LogsDirectory, $"media-pipeline-{day.ToLocalTime():yyyyMMdd}.log");

    /// <summary>
    /// Workspace first, then preset, matching watch-media.ps1's Get-PresetWorkspacePaths and
    /// the layout on the remote. Everything for one client sits together.
    /// </summary>
    public string LaneDirectory(string preset, string workspace) =>
        Path.Combine(PipelineRoot, workspace, preset);

    public string InputDirectory(string preset, string workspace) =>
        Path.Combine(LaneDirectory(preset, workspace), "input");

    public string OutputDirectory(string preset, string workspace) =>
        Path.Combine(LaneDirectory(preset, workspace), "output");

    public string FailedDirectory(string preset, string workspace) =>
        Path.Combine(LaneDirectory(preset, workspace), "failed");

    /// <summary>Uploads stage per workspace, mirroring the remote.</summary>
    public string SyncDirectory(string workspace) =>
        Path.Combine(PipelineRoot, "sync", workspace);

    public string SyncRoot => Path.Combine(PipelineRoot, "sync");

    /// <summary>
    /// The watcher's single-instance mutex for this root. Opening it is the cheapest possible
    /// liveness probe: no process scanning, and it cannot be fooled by a stale status file.
    ///
    /// Must match Get-WatcherMutexName in watch-media.ps1.
    /// </summary>
    public string MutexName
    {
        get
        {
            var normalized = PipelineRoot.TrimEnd('\\', '/').ToLowerInvariant();
            if (normalized == @"d:\mediapipeline")
            {
                return @"Global\MediaPipelineWatcher";
            }

            var hash = SHA256.HashData(Encoding.UTF8.GetBytes(normalized));
            var hex = Convert.ToHexString(hash)[..16];
            return $@"Global\MediaPipelineWatcher_{hex}";
        }
    }

    /// <summary>
    /// Locates the pipeline install. Looks beside the executable first, then walks up, so the
    /// app works both from the repo root and from a published subfolder.
    /// </summary>
    public static PipelinePaths Discover()
    {
        var directory = AppContext.BaseDirectory;

        for (var current = new DirectoryInfo(directory); current is not null; current = current.Parent)
        {
            var script = Path.Combine(current.FullName, "watch-media.ps1");
            if (File.Exists(script))
            {
                return new PipelinePaths(current.FullName, ReadPipelineRoot(current.FullName));
            }
        }

        // Nothing found: fall back to the documented default so the UI can still start and
        // explain itself rather than crashing on launch.
        return new PipelinePaths(directory, @"D:\MediaPipeline");
    }

    private static string ReadPipelineRoot(string appDirectory)
    {
        var configFile = Path.Combine(appDirectory, "config.ini");
        if (!File.Exists(configFile))
        {
            return @"D:\MediaPipeline";
        }

        try
        {
            foreach (var raw in File.ReadLines(configFile))
            {
                var line = raw.Trim();
                if (line.Length == 0 || line[0] is '#' or ';' or '[')
                {
                    continue;
                }

                var equals = line.IndexOf('=');
                if (equals < 1)
                {
                    continue;
                }

                if (!line[..equals].Trim().Equals("PipelineRoot", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var value = IniFile.CleanValue(line[(equals + 1)..]);
                if (value.Length > 0)
                {
                    return value;
                }
            }
        }
        catch (IOException)
        {
            // An unreadable config is not worth failing startup over.
        }

        return @"D:\MediaPipeline";
    }
}
