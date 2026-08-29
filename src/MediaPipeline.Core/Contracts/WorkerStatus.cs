using System.Text.Json.Serialization;
using MediaPipeline.Core.Configuration;

namespace MediaPipeline.Core.Contracts;

public sealed record PresetStatus
{
    public required string Name { get; init; }
    public int VideoCopies { get; init; }
    public int ImageCopies { get; init; }
    [JsonConverter(typeof(JsonStringEnumConverter))]
    public OutputGrouping Grouping { get; init; }
    public int SetCount { get; init; }
    [JsonConverter(typeof(JsonStringEnumConverter))]
    public BatchMode Batch { get; init; }
    public bool Segment { get; init; }
    public bool Manifest { get; init; }
    public double SizeCapMB { get; init; }
}

public sealed record LaneStatus
{
    public required string Preset { get; init; }
    public required string Workspace { get; init; }
    public int Queued { get; init; }
    public bool Paused { get; init; }
}

public sealed record WorkerStatus
{
    public string Schema { get; init; } = "mediaPipeline.status.v2";
    public int Pid { get; init; } = Environment.ProcessId;
    public DateTimeOffset StartedUtc { get; init; }
    public DateTimeOffset UpdatedUtc { get; init; }
    public required string PipelineRoot { get; init; }
    public required string Encoder { get; init; }
    public int PollSeconds { get; init; }
    public bool PausedAll { get; init; }
    public required IReadOnlyList<string> Workspaces { get; init; }
    public required IReadOnlyList<PresetStatus> Presets { get; init; }
    public required IReadOnlyList<LaneStatus> Lanes { get; init; }
}
