namespace MediaPipeline.Core.Configuration;

public enum OutputGrouping
{
    Flat,
    PerSource,
    PerSet,
}

public enum BatchMode
{
    PerFile,
    PerGroup,
}

public enum FailureMode
{
    PreservePartial,
    DeleteFiles,
    DeleteContainer,
}

public enum ParallelMode
{
    OverFiles,
    OverVariants,
    Sequential,
}

public sealed record VideoOptions
{
    public int Crf { get; init; } = 24;
    public string X264Preset { get; init; } = "medium";
    public string AudioBitrate { get; init; } = "128k";
    public int MaxWidth { get; init; } = 1080;
    public double SizeCapMB { get; init; } = 8;
    public int SizeCapFallbackMaxWidth { get; init; } = 720;
    public int MinTrimMs { get; init; } = 15;
    public int MaxTrimMs { get; init; } = 95;
    public int SegmentTargetSeconds { get; init; } = 15;
    public int SegmentMinSeconds { get; init; } = 11;
    public bool PreferVideoToolbox { get; init; } = true;
    public int VideoToolboxBitrateKbps { get; init; } = 6000;
    public bool PreferNvenc { get; init; } = true;
    public bool PreferAmf { get; init; } = true;
    public string NvencPreset { get; init; } = "p4";
    public int NvencCq { get; init; } = 26;
    public string AmfQuality { get; init; } = "balanced";
    public int AmfQp { get; init; } = 24;
    public double MaxrateScale { get; init; } = 0.92;
}

public sealed record ImageOptions
{
    public int ProcessingConcurrency { get; init; }
    public int CropMinPermille { get; init; } = 5;
    public int CropMaxPermille { get; init; } = 20;
    public int JpegQuality { get; init; } = 4;
    public int ConvertedJpegQuality { get; init; } = 12;
    public int PngCompressionLevel { get; init; } = 6;
}

public sealed record TimingOptions
{
    public int StableSeconds { get; init; } = 3;
    public int TimeoutSeconds { get; init; } = 600;
    public int PollSeconds { get; init; } = 2;
}

public sealed record ArchiveOptions
{
    public bool Enabled { get; init; } = true;
    public double AgeHours { get; init; } = 15;
    public int CheckIntervalMinutes { get; init; } = 30;
    public int AssetRetentionDays { get; init; } = 5;
}

public sealed record UploadOptions
{
    public string RemoteName { get; init; } = "heatup-remote";
    public string RemoteSftpPartsRoot { get; init; } = "/D:/MediaPipeline/.sync-parts";
    public string RemotePartsRoot { get; init; } = @"D:\MediaPipeline\.sync-parts";
    public string RemoteDirectory { get; init; } = @"D:\MediaPipeline\sync";
    public string RemoteSshHost { get; init; } = "heatup-remote";
    public int RemoteSshPort { get; init; } = 2222;
    public string RemoteSshKeyFile { get; init; } = "";
    public bool DeleteAfterUpload { get; init; }
    public int ChunkSizeMB { get; init; } = 256;
    public int ParallelChunks { get; init; } = 4;
}

public sealed record PresetOptions
{
    public required string Name { get; init; }
    public bool Enabled { get; init; } = true;
    public int VideoCopies { get; init; } = 1;
    public int ImageCopies { get; init; } = 1;
    public int CopiesAlternate { get; init; }
    public OutputGrouping Grouping { get; init; }
    public int SetCount { get; init; } = 1;
    public BatchMode Batch { get; init; }
    public bool Segment { get; init; }
    public int SegmentTargetSeconds { get; init; }
    public int SegmentMinSeconds { get; init; }
    public bool Manifest { get; init; }
    public string ManifestSchema { get; init; } = "heatup.assetStoreMediaManifest.v1";
    public bool Normalize { get; init; } = true;
    public bool EnhancedVariation { get; init; }
    public FailureMode OnFailure { get; init; }
    public ParallelMode Parallel { get; init; }
    public int MaxWidth { get; init; }
    public string AudioBitrate { get; init; } = "128k";
    public double SizeCapMB { get; init; }
    public int SizeCapFallbackMaxWidth { get; init; }
    public double MaxrateScale { get; init; }
    public int NvencCq { get; init; }
    public int AmfQp { get; init; }
    public int Crf { get; init; }
    public int MinTrimMs { get; init; }
    public int MaxTrimMs { get; init; }
    public int CropMinPermille { get; init; }
    public int CropMaxPermille { get; init; }
    public int JpegQuality { get; init; }
    public int ConvertedJpegQuality { get; init; }
    public int PngCompressionLevel { get; init; }
}

public sealed record PipelineConfiguration
{
    public required string PipelineRoot { get; init; }
    public IReadOnlyList<string> Workspaces { get; init; } = ["LC", "MD", "YL", "PL", "general"];
    public required VideoOptions Video { get; init; }
    public required ImageOptions Images { get; init; }
    public required TimingOptions Timing { get; init; }
    public required ArchiveOptions Archive { get; init; }
    public required UploadOptions Upload { get; init; }
    public required IReadOnlyList<PresetOptions> Presets { get; init; }
    public IReadOnlyList<string> Warnings { get; init; } = [];
}
