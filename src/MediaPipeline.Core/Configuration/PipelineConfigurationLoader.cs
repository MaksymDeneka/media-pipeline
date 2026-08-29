using System.Globalization;
using System.Text.RegularExpressions;
using MediaPipeline.Core.IO;

namespace MediaPipeline.Core.Configuration;

public static class PipelineConfigurationLoader
{
    public static PipelineConfiguration Load(string path)
    {
        var configuration = Resolve(IniDocument.Load(path));
        if (Path.IsPathRooted(configuration.PipelineRoot) ||
            configuration.PipelineRoot == "~" ||
            configuration.PipelineRoot.StartsWith("~/", StringComparison.Ordinal) ||
            configuration.PipelineRoot.StartsWith("~\\", StringComparison.Ordinal))
        {
            return configuration;
        }

        var configDirectory = Path.GetDirectoryName(Path.GetFullPath(path))!;
        return configuration with
        {
            PipelineRoot = Path.GetFullPath(configuration.PipelineRoot, configDirectory),
        };
    }

    public static PipelineConfiguration Resolve(IniDocument document)
    {
        var warnings = new List<string>();
        var globals = new ValueReader(document.Globals, "global configuration", warnings);

        var video = new VideoOptions
        {
            Crf = globals.Int("Crf", 24),
            X264Preset = globals.String("X264Preset", "medium"),
            AudioBitrate = globals.String("AudioBitrate", "128k"),
            MaxWidth = globals.Int("MaxWidth", 1080),
            SizeCapMB = globals.Double("SizeCapMB", 8),
            SizeCapFallbackMaxWidth = globals.Int("SizeCapFallbackMaxWidth", 720),
            MinTrimMs = globals.Int("MinTrimMs", 15),
            MaxTrimMs = globals.Int("MaxTrimMs", 95),
            SegmentTargetSeconds = globals.Int("SegmentTargetSeconds", 15),
            SegmentMinSeconds = globals.Int("SegmentMinSeconds", 11),
            PreferVideoToolbox = globals.Bool("PreferVideoToolbox", true),
            VideoToolboxBitrateKbps = Math.Max(200, globals.Int("VideoToolboxBitrateKbps", 6000)),
            PreferNvenc = globals.Bool("PreferNvenc", true),
            PreferAmf = globals.Bool("PreferAmf", true),
            NvencPreset = globals.String("NvencPreset", "p4"),
            NvencCq = globals.Int("NvencCq", 26),
            AmfQuality = globals.String("AmfQuality", "balanced"),
            AmfQp = globals.Int("AmfQp", 24),
            MaxrateScale = globals.Double("MaxrateScale", 0.92),
        };

        var maxImageWorkers = Math.Clamp(Environment.ProcessorCount, 1, 6);
        var requestedWorkers = globals.String("ImageProcessingConcurrency", "auto");
        var imageWorkers = maxImageWorkers;
        if (int.TryParse(requestedWorkers, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsedWorkers) &&
            parsedWorkers >= 1)
        {
            imageWorkers = Math.Min(parsedWorkers, maxImageWorkers);
            if (parsedWorkers > maxImageWorkers)
            {
                warnings.Add(
                    $"ImageProcessingConcurrency requested {parsedWorkers} workers; capped at {maxImageWorkers}.");
            }
        }

        var images = new ImageOptions
        {
            ProcessingConcurrency = imageWorkers,
            CropMinPermille = globals.Int("CropMinPermille", 5),
            CropMaxPermille = globals.Int("CropMaxPermille", 20),
            JpegQuality = Math.Clamp(globals.Int("JpegQuality", 4), 2, 31),
            ConvertedJpegQuality = Math.Clamp(globals.Int("ConvertedJpegQuality", 12), 2, 31),
            PngCompressionLevel = Math.Clamp(globals.Int("PngCompressionLevel", 6), 0, 9),
        };

        var timing = new TimingOptions
        {
            StableSeconds = Math.Max(0, globals.Int("StableSeconds", 3)),
            TimeoutSeconds = Math.Max(1, globals.Int("TimeoutSeconds", 600)),
            PollSeconds = Math.Max(1, globals.Int("PollSeconds", 2)),
        };

        var archive = new ArchiveOptions
        {
            Enabled = globals.Bool("ArchiveEnabled", true),
            AgeHours = Math.Max(0, globals.Double("ArchiveAgeHours", 15)),
            CheckIntervalMinutes = Math.Max(1, globals.Int("ArchiveCheckIntervalMinutes", 30)),
            AssetRetentionDays = Math.Max(0, globals.Int("AssetRetentionDays", 5)),
        };

        var fallbackSshKey = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            ".ssh",
            "heatup_remote_debug_ed25519");
        var upload = new UploadOptions
        {
            RemoteName = globals.String("RemoteName", "heatup-remote"),
            RemoteSftpPartsRoot = globals.String(
                "RemoteSftpPartsRoot", "/D:/MediaPipeline/.sync-parts"),
            RemotePartsRoot = globals.String(
                "RemotePartsRoot", @"D:\MediaPipeline\.sync-parts"),
            RemoteDirectory = globals.String("RemoteDirectory", @"D:\MediaPipeline\sync"),
            RemoteSshHost = globals.String("RemoteSshHost", "heatup-remote"),
            RemoteSshPort = Math.Clamp(globals.Int("RemoteSshPort", 2222), 1, 65535),
            RemoteSshKeyFile = ExpandHomePath(Environment.ExpandEnvironmentVariables(
                globals.String("RemoteSshKeyFile", fallbackSshKey))),
            DeleteAfterUpload = globals.Bool("DeleteAfterUpload", false),
            ChunkSizeMB = Math.Max(1, globals.Int("ChunkSizeMB", 256)),
            ParallelChunks = Math.Max(1, globals.Int("ParallelChunks", 4)),
        };

        var presetSections = document.Presets.Count > 0
            ? document.Presets
            : BuiltInPresets();

        var presets = presetSections
            .Select(pair => ResolvePreset(pair.Key, pair.Value, video, images, warnings))
            .Where(preset => preset.Enabled)
            .ToArray();

        return new PipelineConfiguration
        {
            PipelineRoot = ResolvePipelineRoot(globals, warnings),
            Video = video,
            Images = images,
            Timing = timing,
            Archive = archive,
            Upload = upload,
            Presets = presets,
            Warnings = warnings,
        };
    }

    private static PresetOptions ResolvePreset(
        string name,
        IReadOnlyDictionary<string, string> values,
        VideoOptions video,
        ImageOptions images,
        List<string> warnings)
    {
        PipelinePaths.ValidateSegment(name, nameof(name));
        var reader = new ValueReader(values, $"preset '{name}'", warnings);
        var grouping = reader.Enum("Grouping", OutputGrouping.Flat);
        var defaultFailure = grouping == OutputGrouping.Flat
            ? FailureMode.PreservePartial
            : FailureMode.DeleteContainer;

        return new PresetOptions
        {
            Name = name,
            Enabled = reader.Bool("Enabled", true),
            VideoCopies = Math.Max(0, reader.Int("VideoCopies", 1)),
            ImageCopies = Math.Max(0, reader.Int("ImageCopies", 1)),
            CopiesAlternate = Math.Max(0, reader.Int("CopiesAlternate", 0)),
            Grouping = grouping,
            SetCount = Math.Max(1, reader.Int("SetCount", 1)),
            Batch = reader.Enum("Batch", BatchMode.PerFile),
            Segment = reader.Bool("Segment", false),
            SegmentTargetSeconds = Math.Max(1, reader.Int("SegmentTargetSeconds", video.SegmentTargetSeconds)),
            SegmentMinSeconds = Math.Max(1, reader.Int("SegmentMinSeconds", video.SegmentMinSeconds)),
            Manifest = reader.Bool("Manifest", false),
            ManifestSchema = reader.String("ManifestSchema", "heatup.assetStoreMediaManifest.v1"),
            Normalize = reader.Bool("Normalize", true),
            OnFailure = reader.Enum("OnFailure", defaultFailure),
            Parallel = reader.Enum("Parallel", ParallelMode.OverFiles),
            MaxWidth = Math.Max(2, reader.Int("MaxWidth", video.MaxWidth)),
            AudioBitrate = reader.String("AudioBitrate", video.AudioBitrate),
            SizeCapMB = Math.Max(0, reader.Double("SizeCapMB", video.SizeCapMB)),
            SizeCapFallbackMaxWidth = Math.Max(2, reader.Int(
                "SizeCapFallbackMaxWidth", video.SizeCapFallbackMaxWidth)),
            MaxrateScale = reader.Double("MaxrateScale", video.MaxrateScale),
            NvencCq = reader.Int("NvencCq", video.NvencCq),
            AmfQp = reader.Int("AmfQp", video.AmfQp),
            Crf = reader.Int("Crf", video.Crf),
            MinTrimMs = Math.Max(0, reader.Int("MinTrimMs", video.MinTrimMs)),
            MaxTrimMs = Math.Max(0, reader.Int("MaxTrimMs", video.MaxTrimMs)),
            CropMinPermille = Math.Max(0, reader.Int("CropMinPermille", images.CropMinPermille)),
            CropMaxPermille = Math.Max(0, reader.Int("CropMaxPermille", images.CropMaxPermille)),
            JpegQuality = Math.Clamp(reader.Int("JpegQuality", images.JpegQuality), 2, 31),
            ConvertedJpegQuality = Math.Clamp(
                reader.Int("ConvertedJpegQuality", images.ConvertedJpegQuality), 2, 31),
            PngCompressionLevel = Math.Clamp(
                reader.Int("PngCompressionLevel", images.PngCompressionLevel), 0, 9),
        };
    }

    private static IReadOnlyDictionary<string, IReadOnlyDictionary<string, string>> BuiltInPresets()
    {
        static IReadOnlyDictionary<string, string> Values(params (string Key, string Value)[] values) =>
            values.ToDictionary(pair => pair.Key, pair => pair.Value, StringComparer.OrdinalIgnoreCase);

        return new Dictionary<string, IReadOnlyDictionary<string, string>>(StringComparer.OrdinalIgnoreCase)
        {
            ["bulk"] = Values(("VideoCopies", "8"), ("ImageCopies", "8"), ("CopiesAlternate", "7")),
            ["video-clean"] = Values(("VideoCopies", "1"), ("ImageCopies", "0")),
            ["image-clean"] = Values(("VideoCopies", "0"), ("ImageCopies", "1")),
            ["image-bulk"] = Values(("VideoCopies", "0"), ("ImageCopies", "20")),
            ["sets"] = Values(
                ("VideoCopies", "10"), ("ImageCopies", "10"),
                ("Grouping", "PerSource"), ("SizeCapMB", "0")),
            ["sets-batch"] = Values(
                ("VideoCopies", "1"), ("ImageCopies", "1"), ("Grouping", "PerSet"),
                ("SetCount", "10"), ("Batch", "PerGroup"), ("SizeCapMB", "0")),
            ["asset-store"] = Values(
                ("VideoCopies", "1"), ("ImageCopies", "1"), ("Grouping", "PerSet"),
                ("SetCount", "15"), ("Batch", "PerGroup"), ("SizeCapMB", "0"),
                ("Manifest", "true"), ("MinTrimMs", "10"), ("MaxTrimMs", "40")),
            ["long"] = Values(
                ("VideoCopies", "3"), ("ImageCopies", "0"), ("Segment", "true"),
                ("NvencCq", "28"), ("AmfQp", "26")),
        };
    }

    private static string DefaultPipelineRoot()
    {
        if (OperatingSystem.IsWindows())
        {
            return @"D:\MediaPipeline";
        }

        var home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        return Path.Combine(home, "MediaPipeline");
    }

    private static string ExpandHomePath(string path)
    {
        if (path == "~")
        {
            return Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        }

        if (path.StartsWith("~/", StringComparison.Ordinal) ||
            path.StartsWith("~\\", StringComparison.Ordinal))
        {
            var relative = path[2..]
                .Replace('/', Path.DirectorySeparatorChar)
                .Replace('\\', Path.DirectorySeparatorChar);
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                relative);
        }

        return path;
    }

    private static string ResolvePipelineRoot(ValueReader globals, List<string> warnings)
    {
        var fallback = DefaultPipelineRoot();
        var configured = globals.String("PipelineRoot", fallback);
        if (!OperatingSystem.IsWindows() &&
            (Regex.IsMatch(configured, "^[A-Za-z]:[\\\\/]") || configured.StartsWith("\\\\")))
        {
            warnings.Add(
                $"PipelineRoot '{configured}' is a Windows path. Using '{fallback}' on this system.");
            return fallback;
        }

        return configured;
    }

    private sealed class ValueReader(
        IReadOnlyDictionary<string, string> values,
        string context,
        List<string> warnings)
    {
        public string String(string key, string fallback) =>
            values.TryGetValue(key, out var value) && !string.IsNullOrWhiteSpace(value)
                ? value.Trim()
                : fallback;

        public int Int(string key, int fallback)
        {
            if (!values.TryGetValue(key, out var value) || string.IsNullOrWhiteSpace(value))
            {
                return fallback;
            }

            if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed))
            {
                return parsed;
            }

            warnings.Add($"{context}: '{key} = {value}' is not a whole number. Using {fallback}.");
            return fallback;
        }

        public double Double(string key, double fallback)
        {
            if (!values.TryGetValue(key, out var value) || string.IsNullOrWhiteSpace(value))
            {
                return fallback;
            }

            if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed))
            {
                return parsed;
            }

            warnings.Add($"{context}: '{key} = {value}' is not a number. Using {fallback}.");
            return fallback;
        }

        public bool Bool(string key, bool fallback)
        {
            if (!values.TryGetValue(key, out var value) || string.IsNullOrWhiteSpace(value))
            {
                return fallback;
            }

            return value.Trim().ToLowerInvariant() switch
            {
                "true" or "1" or "yes" or "on" => true,
                "false" or "0" or "no" or "off" => false,
                _ => WarnAndReturn(key, value, fallback),
            };
        }

        public T Enum<T>(string key, T fallback) where T : struct, Enum
        {
            if (!values.TryGetValue(key, out var value) || string.IsNullOrWhiteSpace(value))
            {
                return fallback;
            }

            if (System.Enum.TryParse<T>(value.Trim(), ignoreCase: true, out var parsed) &&
                System.Enum.IsDefined(parsed))
            {
                return parsed;
            }

            warnings.Add($"{context}: '{key} = {value}' is not supported. Using {fallback}.");
            return fallback;
        }

        private bool WarnAndReturn(string key, string value, bool fallback)
        {
            warnings.Add($"{context}: '{key} = {value}' is not true or false. Using {fallback}.");
            return fallback;
        }
    }
}
