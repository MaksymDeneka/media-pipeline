using System.Text.Json.Serialization;

namespace MediaPipeline.Core.Contracts;

public sealed record PipelineEvent
{
    [JsonPropertyName("ts")]
    public DateTimeOffset Timestamp { get; init; } = DateTimeOffset.UtcNow;

    [JsonPropertyName("ev")]
    public required string Name { get; init; }

    [JsonPropertyName("jobId")]
    public string? JobId { get; init; }

    [JsonPropertyName("preset")]
    public string? Preset { get; init; }

    [JsonPropertyName("workspace")]
    public string? Workspace { get; init; }

    [JsonPropertyName("file")]
    public string? File { get; init; }

    [JsonPropertyName("files")]
    public IReadOnlyList<string>? Files { get; init; }

    [JsonPropertyName("n")]
    public int? Index { get; init; }

    [JsonPropertyName("total")]
    public int? Total { get; init; }

    [JsonPropertyName("outputs")]
    public int? Outputs { get; init; }

    [JsonPropertyName("output")]
    public string? Output { get; init; }

    [JsonPropertyName("bytes")]
    public long? Bytes { get; init; }

    [JsonPropertyName("error")]
    public string? Error { get; init; }

    [JsonPropertyName("pid")]
    public int? Pid { get; init; }

    [JsonPropertyName("pipelineRoot")]
    public string? PipelineRoot { get; init; }

    [JsonPropertyName("encoder")]
    public string? Encoder { get; init; }

    [JsonPropertyName("presets")]
    public IReadOnlyList<string>? Presets { get; init; }

    [JsonPropertyName("workspaces")]
    public IReadOnlyList<string>? Workspaces { get; init; }

    [JsonPropertyName("reason")]
    public string? Reason { get; init; }
}
