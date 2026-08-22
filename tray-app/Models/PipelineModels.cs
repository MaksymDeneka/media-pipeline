using System.Text.Json.Serialization;

namespace MediaPipelineTray.Models;

/// <summary>
/// How a lane or job is doing. Only <see cref="Finished"/> and <see cref="Failed"/> are ever
/// given a colour in the UI; everything else is carried by contrast and shape.
/// </summary>
public enum ActivityState
{
    Idle,
    Queued,
    Running,
    Paused,
    Finished,
    Failed,
}

/// <summary>
/// One preset's configuration, as the watcher resolved it. Mirrors the "presets" entries in
/// status\watcher.json, so this is what the watcher is actually doing rather than what
/// config.ini says on disk.
/// </summary>
public sealed record PresetInfo
{
    [JsonPropertyName("name")] public string Name { get; init; } = "";
    [JsonPropertyName("videoCopies")] public int VideoCopies { get; init; }
    [JsonPropertyName("imageCopies")] public int ImageCopies { get; init; }
    [JsonPropertyName("grouping")] public string Grouping { get; init; } = "Flat";
    [JsonPropertyName("setCount")] public int SetCount { get; init; }
    [JsonPropertyName("batch")] public string Batch { get; init; } = "PerFile";
    [JsonPropertyName("segment")] public bool Segment { get; init; }
    [JsonPropertyName("manifest")] public bool Manifest { get; init; }
    [JsonPropertyName("sizeCapMB")] public double SizeCapMB { get; init; }

    /// <summary>A short human summary, e.g. "100 per image, flat".</summary>
    public string Summary
    {
        get
        {
            var parts = new List<string>();

            if (VideoCopies > 0 && ImageCopies > 0 && VideoCopies == ImageCopies)
            {
                parts.Add($"{VideoCopies} per file");
            }
            else
            {
                if (VideoCopies > 0) parts.Add($"{VideoCopies} per video");
                if (ImageCopies > 0) parts.Add($"{ImageCopies} per image");
            }

            if (Grouping == "PerSource") parts.Add("folder per source");
            if (Grouping == "PerSet") parts.Add($"{SetCount} sets");
            if (Segment) parts.Add("segmented");
            if (Manifest) parts.Add("manifest");

            return parts.Count == 0 ? "no output" : string.Join(", ", parts);
        }
    }
}

/// <summary>One preset in one workspace, with its current queue depth.</summary>
public sealed record LaneInfo
{
    [JsonPropertyName("preset")] public string Preset { get; init; } = "";
    [JsonPropertyName("workspace")] public string Workspace { get; init; } = "";
    [JsonPropertyName("queued")] public int Queued { get; init; }
    [JsonPropertyName("paused")] public bool Paused { get; init; }

    public string Key => $"{Preset}/{Workspace}";
}

/// <summary>
/// The contents of status\watcher.json. Absent or stale means the watcher is not running,
/// which the UI checks against the single-instance mutex rather than trusting this file.
/// </summary>
public sealed record WatcherStatus
{
    [JsonPropertyName("schema")] public string Schema { get; init; } = "";
    [JsonPropertyName("pid")] public int Pid { get; init; }
    [JsonPropertyName("startedUtc")] public DateTimeOffset StartedUtc { get; init; }
    [JsonPropertyName("updatedUtc")] public DateTimeOffset UpdatedUtc { get; init; }
    [JsonPropertyName("pipelineRoot")] public string PipelineRoot { get; init; } = "";
    [JsonPropertyName("encoder")] public string Encoder { get; init; } = "";
    [JsonPropertyName("pollSeconds")] public int PollSeconds { get; init; }
    [JsonPropertyName("pausedAll")] public bool PausedAll { get; init; }
    [JsonPropertyName("workspaces")] public IReadOnlyList<string> Workspaces { get; init; } = [];
    [JsonPropertyName("presets")] public IReadOnlyList<PresetInfo> Presets { get; init; } = [];
    [JsonPropertyName("lanes")] public IReadOnlyList<LaneInfo> Lanes { get; init; } = [];
}

/// <summary>
/// One line from logs\events-YYYYMMDD.jsonl. Every event produced by a single job carries the
/// same <see cref="JobId"/>, which is what makes progress attributable when the watcher runs
/// several files at once.
/// </summary>
public sealed record PipelineEvent
{
    [JsonPropertyName("ts")] public DateTimeOffset Timestamp { get; init; }
    [JsonPropertyName("ev")] public string Name { get; init; } = "";
    [JsonPropertyName("jobId")] public string? JobId { get; init; }
    [JsonPropertyName("preset")] public string? Preset { get; init; }
    [JsonPropertyName("workspace")] public string? Workspace { get; init; }
    [JsonPropertyName("file")] public string? File { get; init; }
    [JsonPropertyName("files")] public IReadOnlyList<string>? Files { get; init; }
    [JsonPropertyName("n")] public int? Index { get; init; }
    [JsonPropertyName("total")] public int? Total { get; init; }
    [JsonPropertyName("outputs")] public int? Outputs { get; init; }
    // Typed as double because PowerShell serializes sums with a decimal point, and a
    // strict integer binding would reject the whole line.
    [JsonPropertyName("bytes")] public double? Bytes { get; init; }
    [JsonPropertyName("error")] public string? Error { get; init; }
    [JsonPropertyName("pid")] public int? Pid { get; init; }
}

/// <summary>
/// A unit of work the watcher took on, assembled by folding the event stream. A job covers one
/// file for a per-file preset, or a whole settled folder for a batch preset.
/// </summary>
public sealed class JobProgress
{
    public required string JobId { get; init; }
    public required string Preset { get; init; }
    public required string Workspace { get; init; }
    public required DateTimeOffset StartedUtc { get; init; }

    public IReadOnlyList<string> Files { get; set; } = [];
    public double Bytes { get; set; }
    public int VariantsDone { get; set; }
    public int VariantsTotal { get; set; }
    public int Outputs { get; set; }
    public string? Error { get; set; }
    public DateTimeOffset? EndedUtc { get; set; }
    public ActivityState State { get; set; } = ActivityState.Running;

    public string Lane => $"{Preset} / {Workspace}";

    /// <summary>The file being worked on, or a count when a batch job covers several.</summary>
    public string Subject => Files.Count switch
    {
        0 => "(no files)",
        1 => Files[0],
        _ => $"{Files.Count} files",
    };

    public double Fraction => VariantsTotal > 0
        ? Math.Clamp((double)VariantsDone / VariantsTotal, 0, 1)
        : 0;

    public TimeSpan Elapsed => (EndedUtc ?? DateTimeOffset.UtcNow) - StartedUtc;

    /// <summary>
    /// Time left, extrapolated from the variants finished so far. Null until at least one
    /// variant has completed, because a single sample is not an estimate.
    /// </summary>
    public TimeSpan? Remaining
    {
        get
        {
            if (State != ActivityState.Running || VariantsDone <= 0 || VariantsTotal <= 0)
            {
                return null;
            }

            var perVariant = Elapsed.TotalSeconds / VariantsDone;
            var left = (VariantsTotal - VariantsDone) * perVariant;
            return TimeSpan.FromSeconds(left);
        }
    }
}
