using MediaPipeline.Core.Configuration;

namespace MediaPipeline.Core.IO;

public sealed record LanePaths
{
    public required string Preset { get; init; }
    public required string Workspace { get; init; }
    public required string Root { get; init; }
    public required string Input { get; init; }
    public required string Output { get; init; }
    public required string Original { get; init; }
    public required string Failed { get; init; }
    public required string Work { get; init; }
    public required string Archive { get; init; }
}

public sealed class PipelinePaths(string root)
{
    public string Root { get; } = Path.GetFullPath(ExpandHome(root));

    public string Logs => Path.Combine(Root, "logs");

    public string Status => Path.Combine(Root, "status");

    public string StatusFile => Path.Combine(Status, "watcher.json");

    public string ActiveJobsFile => Path.Combine(Status, "active-jobs.json");

    public string WorkerLockFile => Path.Combine(Status, "worker.lock");

    public string Control => Path.Combine(Root, "control");

    public string SyncParts => Path.Combine(Root, ".sync-parts");

    public string Sync(string workspace)
    {
        ValidateSegment(workspace, nameof(workspace));
        return Path.Combine(Root, workspace, "sync");
    }

    public LanePaths Lane(string preset, string workspace)
    {
        ValidateSegment(preset, nameof(preset));
        ValidateSegment(workspace, nameof(workspace));
        var laneRoot = Path.Combine(Root, workspace, preset);
        return new LanePaths
        {
            Preset = preset,
            Workspace = workspace,
            Root = laneRoot,
            Input = Path.Combine(laneRoot, "input"),
            Output = Path.Combine(laneRoot, "output"),
            Original = Path.Combine(laneRoot, "original"),
            Failed = Path.Combine(laneRoot, "failed"),
            Work = Path.Combine(laneRoot, "work"),
            Archive = Path.Combine(laneRoot, "archive"),
        };
    }

    /// <summary>
    /// Validates names that become both directory names and control-file scopes. Dots are
    /// intentionally excluded because control flags use them as scope separators.
    /// </summary>
    public static void ValidateSegment(string value, string parameterName)
    {
        if (string.IsNullOrWhiteSpace(value) ||
            value != value.Trim() ||
            Path.IsPathRooted(value) ||
            value.IndexOfAny(['.', ':', '\\', '/']) >= 0 ||
            value.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
        {
            throw new ArgumentException($"Invalid pipeline path segment '{value}'.", parameterName);
        }
    }

    public IEnumerable<string> RequiredDirectories(PipelineConfiguration configuration)
    {
        yield return Logs;
        yield return Status;
        yield return Control;

        foreach (var workspace in configuration.Workspaces)
        {
            yield return Sync(workspace);
            foreach (var preset in configuration.Presets)
            {
                var lane = Lane(preset.Name, workspace);
                yield return lane.Input;
                yield return lane.Output;
                yield return lane.Original;
                yield return lane.Failed;
                yield return lane.Work;
                yield return lane.Archive;
            }
        }
    }

    private static string ExpandHome(string path)
    {
        if (path == "~")
        {
            return Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        }

        if (path.StartsWith("~/", StringComparison.Ordinal) ||
            path.StartsWith("~\\", StringComparison.Ordinal))
        {
            var home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            return Path.Combine(home, path[2..]);
        }

        return path;
    }
}
