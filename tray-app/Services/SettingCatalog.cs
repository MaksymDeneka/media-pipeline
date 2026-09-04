namespace MediaPipelineTray.Services;

public enum SettingKind
{
    Integer,
    Decimal,
    Text,
    Boolean,
    Choice,
}

/// <summary>
/// One editable setting. <see cref="Key"/> matches the config.ini key exactly, which is also
/// the preset option name, because the watcher deliberately uses the same word for both: a
/// global is just the default for every preset.
/// </summary>
public sealed record SettingDefinition
{
    public required string Key { get; init; }
    public required string Label { get; init; }
    public required string Help { get; init; }
    public required SettingKind Kind { get; init; }
    public required string Group { get; init; }

    public IReadOnlyList<string> Choices { get; init; } = [];

    /// <summary>False for options that only make sense once, globally.</summary>
    public bool PresetScoped { get; init; } = true;

    /// <summary>False for options that only make sense per preset, never as a global default.</summary>
    public bool GlobalScoped { get; init; } = true;

    /// <summary>
    /// The watcher's own default, used for preset-only options. Those have no global to
    /// inherit from, so without this the UI would show them as blank and imply they are unset
    /// when the watcher is in fact applying a value.
    /// </summary>
    public string Default { get; init; } = "";
}

/// <summary>
/// The editable surface of config.ini, in one place.
///
/// Both the Presets and Settings views are generated from this rather than hand-written, so a
/// new watcher option becomes a UI field by adding one entry here, and the help text cannot
/// drift away from what the file itself documents.
/// </summary>
public static class SettingCatalog
{
    public static IReadOnlyList<SettingDefinition> All { get; } =
    [
        // --- what comes out -------------------------------------------------
        new()
        {
            Key = "VideoCopies",
            Default = "1",
            Label = "Copies per video",
            Help = "How many outputs each source video produces. Zero makes this preset ignore video.",
            Kind = SettingKind.Integer,
            Group = "Output",
            GlobalScoped = false,
        },
        new()
        {
            Key = "ImageCopies",
            Default = "1",
            Label = "Copies per image",
            Help = "How many outputs each source photo produces. Zero makes this preset ignore images.",
            Kind = SettingKind.Integer,
            Group = "Output",
            GlobalScoped = false,
        },
        new()
        {
            Key = "CopiesAlternate",
            Default = "0",
            Label = "Alternate count",
            Help = "When set, consecutive files alternate between the count above and this one, so a run of files does not all produce the same number. Zero disables it.",
            Kind = SettingKind.Integer,
            Group = "Output",
            GlobalScoped = false,
        },
        new()
        {
            Key = "Grouping",
            Default = "Flat",
            Label = "Grouping",
            Help = "Flat puts every output in one folder. PerSource gives each file its own folder. PerSet builds complete sets.",
            Kind = SettingKind.Choice,
            Choices = ["Flat", "PerSource", "PerSet"],
            Group = "Output",
            GlobalScoped = false,
        },
        new()
        {
            Key = "SetCount",
            Default = "1",
            Label = "Number of sets",
            Help = "How many complete sets to build. Only used when Grouping is PerSet.",
            Kind = SettingKind.Integer,
            Group = "Output",
            GlobalScoped = false,
        },
        new()
        {
            Key = "Batch",
            Default = "PerFile",
            Label = "Batching",
            Help = "PerFile processes each file as it settles. PerGroup waits for the whole folder and treats it as one batch.",
            Kind = SettingKind.Choice,
            Choices = ["PerFile", "PerGroup"],
            Group = "Output",
            GlobalScoped = false,
        },
        new()
        {
            Key = "Segment",
            Default = "false",
            Label = "Split long videos",
            Help = "Legacy: the V2 engine no longer splits videos (its bitrate ladder handles long videos down to 160px). This flag is ignored.",
            Kind = SettingKind.Boolean,
            Group = "Output",
            GlobalScoped = false,
        },
        new()
        {
            Key = "EnhancedVariation",
            Default = "false",
            Label = "Stronger variation",
            Help = "Uses the stronger V2 variation: wider eq ranges, a 4-8 permille recrop, and alternate x264 parameters.",
            Kind = SettingKind.Boolean,
            Group = "Output",
            GlobalScoped = false,
        },
        new()
        {
            Key = "Manifest",
            Default = "false",
            Label = "Write a manifest",
            Help = "Writes a manifest.json next to the output describing every generated file.",
            Kind = SettingKind.Boolean,
            Group = "Output",
            GlobalScoped = false,
        },
        new()
        {
            Key = "Enabled",
            Default = "true",
            Label = "Enabled",
            Help = "Switches this preset off without deleting it or its folders.",
            Kind = SettingKind.Boolean,
            Group = "Output",
            GlobalScoped = false,
        },

        // --- video ----------------------------------------------------------
        new()
        {
            Key = "MaxWidth",
            Label = "Maximum width",
            Help = "Caps the V2 width ladder (1080 down to 160): videos never exceed this width and are never enlarged.",
            Kind = SettingKind.Integer,
            Group = "Video",
        },
        new()
        {
            Key = "SizeCapMB",
            Label = "Size cap, MB",
            Help = "Legacy: the V2 engine targets min(10MB, source size) with its own bitrate ladder. Kept for config compatibility; has no effect.",
            Kind = SettingKind.Decimal,
            Group = "Video",
        },
        new()
        {
            Key = "SizeCapFallbackMaxWidth",
            Label = "Size cap fallback width",
            Help = "Legacy: the V2 ladder descends to 160px on its own. Kept for config compatibility; has no effect.",
            Kind = SettingKind.Integer,
            Group = "Video",
        },
        new()
        {
            Key = "Crf",
            Label = "CPU quality (CRF)",
            Help = "Legacy: the V2 engine uses two-pass libx264 slow CBR from its bitrate ladder. Kept for config compatibility; has no effect.",
            Kind = SettingKind.Integer,
            Group = "Video",
        },
        new()
        {
            Key = "NvencCq",
            Label = "NVIDIA quality (CQ)",
            Help = "Legacy: NVENC support was removed (VideoToolbox or libx264 only). Kept for config compatibility; has no effect.",
            Kind = SettingKind.Integer,
            Group = "Video",
        },
        new()
        {
            Key = "AmfQp",
            Label = "AMD quality (QP)",
            Help = "Legacy: AMF support was removed (VideoToolbox or libx264 only). Kept for config compatibility; has no effect.",
            Kind = SettingKind.Integer,
            Group = "Video",
        },
        new()
        {
            Key = "AudioBitrate",
            Label = "Audio bitrate",
            Help = "Legacy: the V2 engine picks 128/96/64/48/32/16/0k from its audio ladder. Kept for config compatibility; has no effect.",
            Kind = SettingKind.Text,
            Group = "Video",
        },
        new()
        {
            Key = "MinTrimMs",
            Label = "Minimum trim, ms",
            Help = "Legacy: the V2 engine trims a deterministic 10-40ms per copy. Kept for config compatibility; has no effect.",
            Kind = SettingKind.Integer,
            Group = "Video",
        },
        new()
        {
            Key = "MaxTrimMs",
            Label = "Maximum trim, ms",
            Help = "Legacy: see Minimum trim. Kept for config compatibility; has no effect.",
            Kind = SettingKind.Integer,
            Group = "Video",
        },
        new()
        {
            Key = "SegmentTargetSeconds",
            Label = "Segment length, seconds",
            Help = "Legacy: the V2 engine never segments. Kept for config compatibility; has no effect.",
            Kind = SettingKind.Integer,
            Group = "Video",
        },
        new()
        {
            Key = "SegmentMinSeconds",
            Label = "Shortest segment, seconds",
            Help = "Legacy: the V2 engine never segments. Kept for config compatibility; has no effect.",
            Kind = SettingKind.Integer,
            Group = "Video",
        },
        new()
        {
            Key = "X264Preset",
            Label = "CPU encoder speed",
            Help = "Legacy: the V2 engine always uses two-pass slow (CPU) or VideoToolbox. Kept for config compatibility; has no effect.",
            Kind = SettingKind.Choice,
            Choices = ["ultrafast", "superfast", "veryfast", "faster", "fast", "medium", "slow", "slower", "veryslow"],
            Group = "Video",
            PresetScoped = false,
        },
        new()
        {
            Key = "PreferVideoToolbox",
            Label = "Use Apple hardware encoding",
            Help = "Probes the VideoToolbox encoder first and falls back to CPU libx264 when unavailable. Off forces libx264.",
            Kind = SettingKind.Boolean,
            Group = "Video",
            PresetScoped = false,
        },
        new()
        {
            Key = "PreferNvenc",
            Label = "Use the NVIDIA encoder",
            Help = "Legacy: NVENC support was removed. Kept for config compatibility; has no effect.",
            Kind = SettingKind.Boolean,
            Group = "Video",
            PresetScoped = false,
        },
        new()
        {
            Key = "PreferAmf",
            Label = "Use the AMD encoder",
            Help = "Legacy: AMF support was removed. Kept for config compatibility; has no effect.",
            Kind = SettingKind.Boolean,
            Group = "Video",
            PresetScoped = false,
        },

        // --- images ---------------------------------------------------------
        new()
        {
            Key = "JpegQuality",
            Label = "JPEG quality",
            Help = "Legacy: the V2 engine always uses q:v 2 (JPEG) / quality 92 (WebP) / level 6 (PNG). Kept for config compatibility; has no effect.",
            Kind = SettingKind.Integer,
            Group = "Images",
        },
        new()
        {
            Key = "ConvertedJpegQuality",
            Label = "JPEG quality, converted",
            Help = "Legacy: the V2 engine uses the same quality for all images. Kept for config compatibility; has no effect.",
            Kind = SettingKind.Integer,
            Group = "Images",
        },
        new()
        {
            Key = "CropMinPermille",
            Label = "Minimum crop, permille",
            Help = "Legacy: the V2 engine recrops a deterministic 2-5 permille per copy. Kept for config compatibility; has no effect.",
            Kind = SettingKind.Integer,
            Group = "Images",
        },
        new()
        {
            Key = "CropMaxPermille",
            Label = "Maximum crop, permille",
            Help = "Legacy: see Minimum crop. Kept for config compatibility; has no effect.",
            Kind = SettingKind.Integer,
            Group = "Images",
        },
        new()
        {
            Key = "ImageProcessingConcurrency",
            Label = "Files at once",
            Help = "How many files are processed in parallel. Requires PowerShell 7. Use auto to pick a safe value. Manual values are capped at six workers.",
            Kind = SettingKind.Text,
            Group = "Images",
            PresetScoped = false,
        },

        // --- timing ---------------------------------------------------------
        new()
        {
            Key = "StableSeconds",
            Label = "Settle time, seconds",
            Help = "A file must stop changing for this long before it is processed, which lets browser downloads finish.",
            Kind = SettingKind.Integer,
            Group = "Timing",
            PresetScoped = false,
        },
        new()
        {
            Key = "PollSeconds",
            Label = "Poll interval, seconds",
            Help = "How often the watcher checks the input folders.",
            Kind = SettingKind.Integer,
            Group = "Timing",
            PresetScoped = false,
        },
        new()
        {
            Key = "TimeoutSeconds",
            Label = "Give up after, seconds",
            Help = "How long to wait for a single file to finish arriving before giving up on it.",
            Kind = SettingKind.Integer,
            Group = "Timing",
            PresetScoped = false,
        },

        // --- uploads --------------------------------------------------------
        new()
        {
            Key = "DeleteAfterUpload",
            Default = "false",
            Label = "Delete after upload",
            Help = "Removes the local file once the remote copy has been read back and confirmed the right size. Off by default, because it is not reversible.",
            Kind = SettingKind.Boolean,
            Group = "Uploads",
            PresetScoped = false,
        },
        new()
        {
            Key = "ChunkSizeMB",
            Default = "256",
            Label = "Chunk size, MB",
            Help = "Large files are split into chunks this size. Smaller chunks retry faster on a flaky link, larger ones have less overhead.",
            Kind = SettingKind.Integer,
            Group = "Uploads",
            PresetScoped = false,
        },
        new()
        {
            Key = "ParallelChunks",
            Default = "4",
            Label = "Chunks at once",
            Help = "How many chunks are sent in parallel.",
            Kind = SettingKind.Integer,
            Group = "Uploads",
            PresetScoped = false,
        },
        new()
        {
            Key = "RemoteDirectory",
            Default = @"D:\MediaPipeline\sync",
            Label = "Remote folder",
            Help = @"Where uploads land on the remote. A workspace folder is created inside it, so a file staged in LC\sync arrives there under LC.",
            Kind = SettingKind.Text,
            Group = "Uploads",
            PresetScoped = false,
        },

        // --- housekeeping ---------------------------------------------------
        new()
        {
            Key = "ArchiveEnabled",
            Label = "Archive old output",
            Help = "Moves old output files into an archive folder so the output folders stay tidy.",
            Kind = SettingKind.Boolean,
            Group = "Housekeeping",
            PresetScoped = false,
        },
        new()
        {
            Key = "ArchiveAgeHours",
            Label = "Archive after, hours",
            Help = "Output older than this gets archived.",
            Kind = SettingKind.Integer,
            Group = "Housekeeping",
            PresetScoped = false,
        },
        new()
        {
            Key = "ArchiveCheckIntervalMinutes",
            Label = "Archive check, minutes",
            Help = "How often the watcher looks for files to archive.",
            Kind = SettingKind.Integer,
            Group = "Housekeeping",
            PresetScoped = false,
        },
        new()
        {
            Key = "AssetRetentionDays",
            Label = "Delete retained after, days",
            Help = "Deletes archived, original and failed entries this many days after they were created. Zero disables it. The images preset is always excluded.",
            Kind = SettingKind.Integer,
            Group = "Housekeeping",
            PresetScoped = false,
        },
    ];

    public static IEnumerable<SettingDefinition> ForPreset() => All.Where(s => s.PresetScoped);

    public static IEnumerable<SettingDefinition> ForGlobal() => All.Where(s => s.GlobalScoped);

    public static IEnumerable<string> Groups(IEnumerable<SettingDefinition> definitions) =>
        definitions.Select(d => d.Group).Distinct();
}
