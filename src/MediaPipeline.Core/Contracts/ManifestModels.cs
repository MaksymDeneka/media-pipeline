namespace MediaPipeline.Core.Contracts;

public sealed record MediaManifest
{
    public required string Schema { get; init; }
    public DateTimeOffset GeneratedAt { get; init; }
    public string ImportRoot { get; init; } = ".";
    public required IReadOnlyList<ManifestVariant> Variants { get; init; }
}

public sealed record ManifestVariant
{
    public required string FamilyKey { get; init; }
    public required string VariantKey { get; init; }
    public required string Path { get; init; }
    public required string RenditionSetKey { get; init; }
    public string? GenerationBatchKey { get; init; }
    public required string SourceOriginalName { get; init; }
    public required string SourceFamilyName { get; init; }
    public long SizeBytes { get; init; }
    public DateTimeOffset GeneratedAt { get; init; }
    public double DurationSeconds { get; init; }
    public required string TransformProfile { get; init; }
    public required ManifestMetadata Metadata { get; init; }
}

public sealed record ManifestMetadata
{
    public required string Encoder { get; init; }
    public int TrimMs { get; init; }
    public int MaxWidth { get; init; }
    public int? SourceWidth { get; init; }
    public int? SourceHeight { get; init; }
}
