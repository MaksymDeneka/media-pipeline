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
            Help = "Splits a long video into segments first, then makes the copy count of each segment.",
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
            Help = "Videos are shrunk so their width is at most this many pixels. They are never enlarged.",
            Kind = SettingKind.Integer,
            Group = "Video",
        },
        new()
        {
            Key = "SizeCapMB",
            Label = "Size cap, MB",
            Help = "Caps each output video at this many megabytes. Zero means no cap and no retry pass.",
            Kind = SettingKind.Decimal,
            Group = "Video",
        },
        new()
        {
            Key = "SizeCapFallbackMaxWidth",
            Label = "Size cap fallback width",
            Help = "If a video still will not fit under the cap, shrink its width to this as a last resort.",
            Kind = SettingKind.Integer,
            Group = "Video",
        },
        new()
        {
            Key = "Crf",
            Label = "CPU quality (CRF)",
            Help = "Used when encoding on the CPU. Lower is better quality and bigger files. Roughly 18 to 28.",
            Kind = SettingKind.Integer,
            Group = "Video",
        },
        new()
        {
            Key = "NvencCq",
            Label = "NVIDIA quality (CQ)",
            Help = "Used when the NVIDIA encoder is active. Lower is better quality.",
            Kind = SettingKind.Integer,
            Group = "Video",
        },
        new()
        {
            Key = "AmfQp",
            Label = "AMD quality (QP)",
            Help = "Used when the AMD encoder is active. Lower is better quality.",
            Kind = SettingKind.Integer,
            Group = "Video",
        },
        new()
        {
            Key = "AudioBitrate",
            Label = "Audio bitrate",
            Help = "For example 96k, 128k, or 192k.",
            Kind = SettingKind.Text,
            Group = "Video",
        },
        new()
        {
            Key = "MinTrimMs",
            Label = "Minimum trim, ms",
            Help = "Each video copy trims a small random amount off the end, which is what makes the copies differ.",
            Kind = SettingKind.Integer,
            Group = "Video",
        },
        new()
        {
            Key = "MaxTrimMs",
            Label = "Maximum trim, ms",
            Help = "The upper end of that random trim.",
            Kind = SettingKind.Integer,
            Group = "Video",
        },
        new()
        {
            Key = "SegmentTargetSeconds",
            Label = "Segment length, seconds",
            Help = "Target length of each segment, for presets that split long videos.",
            Kind = SettingKind.Integer,
            Group = "Video",
        },
        new()
        {
            Key = "SegmentMinSeconds",
            Label = "Shortest segment, seconds",
            Help = "A segment is never shorter than this.",
            Kind = SettingKind.Integer,
            Group = "Video",
        },
        new()
        {
            Key = "X264Preset",
            Label = "CPU encoder speed",
            Help = "Slower settings produce smaller files and take longer.",
            Kind = SettingKind.Choice,
            Choices = ["ultrafast", "superfast", "veryfast", "faster", "fast", "medium", "slow", "slower", "veryslow"],
            Group = "Video",
            PresetScoped = false,
        },
        new()
        {
            Key = "PreferNvenc",
            Label = "Use the NVIDIA encoder",
            Help = "The watcher only uses a GPU encoder if a real test encode succeeds, so this is safe to leave on.",
            Kind = SettingKind.Boolean,
            Group = "Video",
            PresetScoped = false,
        },
        new()
        {
            Key = "PreferAmf",
            Label = "Use the AMD encoder",
            Help = "Same as above, for AMD hardware.",
            Kind = SettingKind.Boolean,
            Group = "Video",
            PresetScoped = false,
        },

        // --- images ---------------------------------------------------------
        new()
        {
            Key = "JpegQuality",
            Label = "JPEG quality",
            Help = "Lower is better quality and larger files. 4 is a high-quality middle ground for photos.",
            Kind = SettingKind.Integer,
            Group = "Images",
        },
        new()
        {
            Key = "ConvertedJpegQuality",
            Label = "JPEG quality, converted",
            Help = "For sources that had to be decoded first, such as HEIC. That round trip is already a re-encode, so it gets more headroom.",
            Kind = SettingKind.Integer,
            Group = "Images",
        },
        new()
        {
            Key = "CropMinPermille",
            Label = "Minimum crop, permille",
            Help = "Every image copy gets a tiny random crop scaled back to the original size, which is what makes the copies differ. 5 is 0.5 percent.",
            Kind = SettingKind.Integer,
            Group = "Images",
        },
        new()
        {
            Key = "CropMaxPermille",
            Label = "Maximum crop, permille",
            Help = "The upper end of that random crop. Keep it small.",
            Kind = SettingKind.Integer,
            Group = "Images",
        },
        new()
        {
            Key = "ImageProcessingConcurrency",
            Label = "Files at once",
            Help = "How many files are processed in parallel. Requires PowerShell 7. Use auto to pick from the CPU count.",
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
