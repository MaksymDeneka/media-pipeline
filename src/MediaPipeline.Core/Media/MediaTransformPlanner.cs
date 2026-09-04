using System.Globalization;

namespace MediaPipeline.Core.Media;

/// <summary>Planning failed because no rung can satisfy the quality floor.</summary>
public sealed class VideoQualityFloorException(string message) : InvalidOperationException(message);

/// <summary>
/// Pure deterministic V2 recipe ported from heatup sidecar/media-transform/profiles.ts.
/// Current recipe only (recipeVersion 2). No grain/noise path – heatup V2 removed it.
/// </summary>
public static class MediaTransformPlanner
{
    public const string ImageRecropV2Profile = "image-recrop-v2";
    public const string VideoMicrotrimV2Profile = "video-microtrim-v2";
    public const string ImageRecropEnhancedV2Profile = "image-recrop-enhanced-v2";
    public const string VideoMicrotrimEnhancedV2Profile = "video-microtrim-enhanced-v2";

    public const long MaxImageSourceBytes = 100L * 1024 * 1024;
    public const long MaxVideoSourceBytes = 2L * 1024 * 1024 * 1024;
    public const long MaxImagePreparedBytes = 50L * 1024 * 1024;
    public const long MaxVideoPreparedV2Bytes = 10L * 1024 * 1024;
    public const int MaxSourceDimension = 16_384;
    public const long MaxSourcePixels = 60_000_000;

    private const long V2VideoPreferredBytes = 5L * 1024 * 1024;
    private const double V2VideoTargetRatio = 0.96;
    private const long V2ReferencePixels = 720L * 1280L;
    private static readonly int[] V2VideoWidthLadder = [1080, 900, 720, 640, 540, 480, 360, 320, 240, 180, 160];

    public sealed record PlannedImageTransform(
        string Profile,
        string OutputExtension,
        IReadOnlyList<string> Args,
        long OutputByteLimit,
        Dictionary<string, object?> Evidence);

    public sealed record PlannedVideoTransform(
        string Profile,
        string OutputExtension,
        IReadOnlyList<string> Args,
        IReadOnlyList<string>? FirstPassArgs,
        string? PasslogPrefix,
        long OutputByteLimit,
        double TargetDurationSeconds,
        Dictionary<string, object?> Evidence);

    public static string OutputExtensionFor(string mediaKind, string sourceExtension)
    {
        if (mediaKind == "video")
        {
            return ".mp4";
        }

        return sourceExtension.Equals(".heic", StringComparison.OrdinalIgnoreCase) ||
               sourceExtension.Equals(".heif", StringComparison.OrdinalIgnoreCase)
            ? ".jpg"
            : sourceExtension.ToLowerInvariant();
    }

    public static int? NormalizeRotationDegrees(object? value)
    {
        double parsed = value switch
        {
            double d => d,
            float f => f,
            int i => i,
            long l => l,
            string s when double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var p) => p,
            _ => double.NaN,
        };
        if (!double.IsFinite(parsed))
        {
            return null;
        }

        var normalized = ((int)Math.Round(parsed) % 360 + 360) % 360;
        return normalized is 0 or 90 or 180 or 270 ? normalized : null;
    }

    public static PlannedImageTransform PlanImage(
        string seed,
        string sourcePath,
        string outputPath,
        string sourceExtension,
        int probeWidth,
        int probeHeight,
        int? rotationDegrees,
        bool enhanced)
    {
        if (probeWidth < 1 || probeHeight < 1)
        {
            throw new InvalidDataException("Image transform source has no readable video frame.");
        }

        var (width, height) = DisplayDimensions(probeWidth, probeHeight, rotationDegrees);
        var outputExtension = OutputExtensionFor("image", sourceExtension);
        var profile = enhanced ? ImageRecropEnhancedV2Profile : ImageRecropV2Profile;

        var brightnessRange = enhanced ? 2_000 : 1_200;
        var contrastRange = enhanced ? 5_000 : 2_500;
        var saturationRange = enhanced ? 6_000 : 3_000;
        var gammaRange = enhanced ? 2_500 : 1_500;
        var brightness = MediaTransformSeed.SeededSignedFraction(seed, brightnessRange, 3);
        var contrast = 1 + MediaTransformSeed.SeededSignedFraction(seed, contrastRange, 4);
        var saturation = 1 + MediaTransformSeed.SeededSignedFraction(seed, saturationRange, 5);
        var gamma = 1 + MediaTransformSeed.SeededSignedFraction(seed, gammaRange, 6);
        var crop = PlanV2Crop(seed, width, height, enhanced);

        var filter = string.Join(",", new[]
        {
            crop.Cropped ? $"crop={crop.CropWidth}:{crop.CropHeight}:{crop.OffsetX}:{crop.OffsetY}" : null,
            crop.Cropped ? $"scale={width}:{height}:flags=lanczos" : null,
            $"eq=brightness={brightness:F6}:contrast={contrast:F6}:gamma={gamma:F6}:saturation={saturation:F6}",
        }.Where(part => part is not null));

        var args = new List<string>(ImageInputArgs(sourcePath))
        {
            "-filter_complex", $"[0:v:0]{filter}[v]", "-map", "[v]",
        };
        AppendImageEncoderArgs(args, outputExtension);
        args.Add(outputPath);

        var evidence = new Dictionary<string, object?>
        {
            ["profile"] = profile,
            ["variationPreset"] = enhanced ? "enhanced" : "default",
            ["sourceWidth"] = probeWidth,
            ["sourceHeight"] = probeHeight,
            ["displayWidth"] = width,
            ["displayHeight"] = height,
            ["rotationDegrees"] = rotationDegrees ?? 0,
            ["brightness"] = brightness,
            ["contrast"] = contrast,
            ["gamma"] = gamma,
            ["saturation"] = saturation,
            ["grain"] = null,
            ["cropped"] = crop.Cropped,
        };
        if (crop.Cropped)
        {
            evidence["cropPermilleX"] = crop.CropPermilleX;
            evidence["cropPermilleY"] = crop.CropPermilleY;
            evidence["cropWidth"] = crop.CropWidth;
            evidence["cropHeight"] = crop.CropHeight;
            evidence["offsetX"] = crop.OffsetX;
            evidence["offsetY"] = crop.OffsetY;
        }

        return new PlannedImageTransform(profile, outputExtension, args, MaxImagePreparedBytes, evidence);
    }

    public static PlannedVideoTransform PlanVideo(
        string seed,
        int ordinal,
        string sourcePath,
        string outputPath,
        int probeWidth,
        int probeHeight,
        int? rotationDegrees,
        double durationSeconds,
        long sourceByteCount,
        bool hardwareVideoToolbox,
        bool enhanced,
        int maxWidthCap = 1080,
        double? targetRatioOverride = null)
    {
        if (durationSeconds <= 0.1)
        {
            throw new InvalidDataException("Video transform source has no usable duration.");
        }

        // long: 1h of video is 3.6e9µs, past int.MaxValue (~35.8min). Matches the
        // reference (JS number) without overflow.
        var durationMicroseconds = (long)Math.Floor(durationSeconds * 1_000_000);
        var canTrim = durationMicroseconds >= 340_000;
        var trimFloor = enhanced ? 20_000 : 10_000;
        const int trimRange = 30_000;
        var trimMicroseconds = canTrim
            ? trimFloor + ((MediaTransformSeed.SeededInteger(seed, 0, trimRange, 0) + ordinal) % (trimRange + 1))
            : 0;
        var targetDuration = Math.Max(0.1, durationSeconds - trimMicroseconds / 1_000_000.0);

        var outputByteLimit = Math.Min(MaxVideoPreparedV2Bytes, sourceByteCount > 0 ? sourceByteCount : MaxVideoPreparedV2Bytes);
        var (displayWidth, displayHeight) = DisplayDimensions(probeWidth, probeHeight, rotationDegrees);
        var dimensions = PlanV2VideoDimensions(displayWidth, displayHeight, targetDuration, outputByteLimit, maxWidthCap);
        var qualityFloorTotalBps = TotalBitrateForVideoFloorBps(dimensions.QualityFloorVideoBitrateBps);
        var qualityFloorBytes = (long)Math.Ceiling(qualityFloorTotalBps * targetDuration / 8.0);
        // Retry override only shrinks the byte target (same rung, lower bitrate);
        // rung selection above always uses the reference 0.96 constant, so the
        // first attempt of every job is heatup-identical.
        var targetRatio = targetRatioOverride ?? V2VideoTargetRatio;
        var targetOutputBytes = Math.Max(1L, (long)Math.Floor(Math.Min(
            outputByteLimit * targetRatio,
            Math.Max(V2VideoPreferredBytes * V2VideoTargetRatio, qualityFloorBytes))));
        var totalBitrateBps = Math.Max(1L, (long)Math.Floor(targetOutputBytes * 8.0 / targetDuration));
        var audioBitrateBps = PlanV2AudioBitrateBps(totalBitrateBps);
        var budgetedVideoBitrateBps = totalBitrateBps - audioBitrateBps;
        var hardwareCeilingKbps = 8_000;
        var hardwareVideoBitrateKbps = Math.Min((int)(budgetedVideoBitrateBps / 1_000), hardwareCeilingKbps);
        var videoBitrateBps = hardwareVideoToolbox ? hardwareVideoBitrateKbps * 1_000L : budgetedVideoBitrateBps;
        if (videoBitrateBps < dimensions.QualityFloorVideoBitrateBps)
        {
            throw new VideoQualityFloorException(
                "Media preparation could not satisfy its video quality floor within the output byte limit.");
        }

        var maxVideoBitrateBps = Math.Max(videoBitrateBps, (long)Math.Floor(videoBitrateBps * 1.15));
        var bufferSizeBps = videoBitrateBps * 2;

        var brightnessRange = enhanced ? 2_000 : 1_200;
        var contrastRange = enhanced ? 5_000 : 2_500;
        var saturationRange = enhanced ? 6_000 : 3_000;
        var gammaRange = enhanced ? 2_500 : 1_500;
        var brightness = MediaTransformSeed.SeededSignedFraction(seed, brightnessRange, 1);
        var contrast = 1 + MediaTransformSeed.SeededSignedFraction(seed, contrastRange, 2);
        var gamma = 1 + MediaTransformSeed.SeededSignedFraction(seed, gammaRange, 3);
        var saturation = 1 + MediaTransformSeed.SeededSignedFraction(seed, saturationRange, 4);
        var crop = enhanced ? PlanV2Crop(seed, displayWidth, displayHeight, true) : default;

        var videoFilter = string.Join(",", new[]
        {
            crop.Cropped ? $"crop={crop.CropWidth}:{crop.CropHeight}:{crop.OffsetX}:{crop.OffsetY}" : null,
            $"scale={dimensions.Width}:{dimensions.Height}:flags=lanczos",
            $"eq=brightness={brightness:F6}:contrast={contrast:F6}:gamma={gamma:F6}:saturation={saturation:F6}",
        }.Where(part => part is not null));

        var x264Params = enhanced
            ? $"keyint={MediaTransformSeed.SeededInteger(seed, 190, 230, 7)}:min-keyint={MediaTransformSeed.SeededInteger(seed, 24, 30, 8)}:scenecut=35:ref=4:bframes=4:aq-mode=3:deblock=-1,-1"
            : $"keyint={MediaTransformSeed.SeededInteger(seed, 210, 250, 7)}:min-keyint={MediaTransformSeed.SeededInteger(seed, 24, 30, 8)}:scenecut=40:ref=3:bframes=3:aq-mode=2";

        var passlogPrefix = $"{outputPath}.passlog";
        var videoArgs = new List<string>
        {
            "-c:v", hardwareVideoToolbox ? "h264_videotoolbox" : "libx264",
        };
        if (!hardwareVideoToolbox)
        {
            videoArgs.AddRange(["-preset", "slow"]);
        }

        videoArgs.AddRange(["-b:v", hardwareVideoToolbox ? $"{hardwareVideoBitrateKbps}k" : videoBitrateBps.ToString(CultureInfo.InvariantCulture)]);
        videoArgs.AddRange(["-maxrate", maxVideoBitrateBps.ToString(CultureInfo.InvariantCulture)]);
        videoArgs.AddRange(["-bufsize", bufferSizeBps.ToString(CultureInfo.InvariantCulture)]);
        videoArgs.AddRange(["-vf", videoFilter, "-pix_fmt", "yuv420p"]);
        if (!hardwareVideoToolbox)
        {
            videoArgs.AddRange(["-x264-params", x264Params]);
        }

        List<string>? firstPassArgs = null;
        if (!hardwareVideoToolbox)
        {
            firstPassArgs =
            [
                "-y", "-hide_banner", "-loglevel", "error",
                "-i", sourcePath,
                "-t", targetDuration.ToString("F6", CultureInfo.InvariantCulture),
                "-map", "0:v:0",
                .. videoArgs,
                "-pass", "1",
                "-passlogfile", passlogPrefix,
                "-an",
                "-map_metadata", "-1",
                "-f", "null",
                "-",
            ];
        }

        var args = new List<string>
        {
            "-y", "-hide_banner", "-loglevel", "error",
            "-i", sourcePath,
            "-t", targetDuration.ToString("F6", CultureInfo.InvariantCulture),
            "-map", "0:v:0",
        };
        if (audioBitrateBps > 0)
        {
            args.AddRange(["-map", "0:a:0?"]);
        }

        args.AddRange(videoArgs);
        if (!hardwareVideoToolbox)
        {
            args.AddRange(["-pass", "2", "-passlogfile", passlogPrefix]);
        }

        if (audioBitrateBps > 0)
        {
            args.AddRange(["-c:a", "aac", "-b:a", audioBitrateBps.ToString(CultureInfo.InvariantCulture)]);
        }
        else
        {
            args.Add("-an");
        }

        args.AddRange(["-movflags", "+faststart", "-map_metadata", "-1", outputPath]);

        var profile = hardwareVideoToolbox
            ? "video-microtrim-videotoolbox-v1"
            : enhanced ? VideoMicrotrimEnhancedV2Profile : VideoMicrotrimV2Profile;

        var evidence = new Dictionary<string, object?>
        {
            ["profile"] = profile,
            ["variationPreset"] = enhanced ? "enhanced" : "default",
            ["encoder"] = hardwareVideoToolbox ? "h264_videotoolbox" : "libx264",
            ["preset"] = hardwareVideoToolbox ? null : "slow",
            ["rateControl"] = hardwareVideoToolbox ? "bitrate" : "two-pass",
            ["audioBitrateBps"] = audioBitrateBps,
            ["audioOmittedForBudget"] = audioBitrateBps == 0,
            ["videoBitrateBps"] = videoBitrateBps,
            ["maxVideoBitrateBps"] = maxVideoBitrateBps,
            ["videoBufferSizeBps"] = bufferSizeBps,
            ["x264Parameters"] = hardwareVideoToolbox ? null : x264Params,
            ["sourceByteCount"] = sourceByteCount,
            ["targetOutputBytes"] = targetOutputBytes,
            ["targetRatio"] = targetRatio,
            ["outputByteLimit"] = outputByteLimit,
            ["sourceWidth"] = probeWidth,
            ["sourceHeight"] = probeHeight,
            ["displayWidth"] = displayWidth,
            ["displayHeight"] = displayHeight,
            ["rotationDegrees"] = rotationDegrees ?? 0,
            ["outputWidth"] = dimensions.Width,
            ["outputHeight"] = dimensions.Height,
            ["qualityFloorVideoBitrateBps"] = dimensions.QualityFloorVideoBitrateBps,
            ["qualityFloorSatisfied"] = true,
            ["trimMicroseconds"] = trimMicroseconds,
            ["brightness"] = brightness,
            ["contrast"] = contrast,
            ["gamma"] = gamma,
            ["saturation"] = saturation,
            ["grain"] = null,
            ["cropped"] = crop.Cropped,
            ["targetDurationSeconds"] = Math.Round(targetDuration, 6),
        };
        if (crop.Cropped)
        {
            evidence["cropWidth"] = crop.CropWidth;
            evidence["cropHeight"] = crop.CropHeight;
            evidence["offsetX"] = crop.OffsetX;
            evidence["offsetY"] = crop.OffsetY;
        }

        return new PlannedVideoTransform(
            profile, ".mp4", args, firstPassArgs,
            hardwareVideoToolbox ? null : passlogPrefix,
            outputByteLimit, targetDuration, evidence);
    }

    private static IReadOnlyList<string> ImageInputArgs(string sourcePath) =>
        ["-y", "-hide_banner", "-loglevel", "error", "-i", sourcePath, "-frames:v", "1", "-map_metadata", "-1"];

    private static void AppendImageEncoderArgs(List<string> args, string outputExtension)
    {
        if (outputExtension is ".jpg" or ".jpeg")
        {
            args.AddRange(["-q:v", "2"]);
        }
        else if (outputExtension == ".webp")
        {
            args.AddRange(["-quality", "92"]);
        }
        else if (outputExtension == ".png")
        {
            args.AddRange(["-compression_level", "6"]);
        }
        else
        {
            throw new InvalidOperationException($"Unsupported image preparation output: {outputExtension}");
        }
    }

    private static (int Width, int Height) DisplayDimensions(int width, int height, int? rotationDegrees)
    {
        var swaps = rotationDegrees is 90 or 270;
        return swaps ? (height, width) : (width, height);
    }

    private readonly record struct CropPlan(
        bool Cropped,
        int CropPermilleX = 0,
        int CropPermilleY = 0,
        int CropWidth = 0,
        int CropHeight = 0,
        int OffsetX = 0,
        int OffsetY = 0);

    private static CropPlan PlanV2Crop(string seed, int width, int height, bool enhanced)
    {
        if (width < 200 || height < 200)
        {
            return new CropPlan(false);
        }

        var minPermille = enhanced ? 4 : 2;
        var maxPermille = enhanced ? 8 : 5;
        var cropPermilleX = MediaTransformSeed.SeededInteger(seed, minPermille, maxPermille, 9);
        var cropPermilleY = MediaTransformSeed.SeededInteger(seed, minPermille, maxPermille, 10);
        var requestedWidth = Math.Max(2, width - (width * cropPermilleX / 1_000));
        var requestedHeight = Math.Max(2, height - (height * cropPermilleY / 1_000));
        var cropWidth = Math.Max(2, requestedWidth / 2 * 2);
        var cropHeight = Math.Max(2, requestedHeight / 2 * 2);
        var removedX = Math.Max(0, width - cropWidth);
        var removedY = Math.Max(0, height - cropHeight);
        var offsetX = MediaTransformSeed.SeededInteger(seed, 0, removedX, 11);
        var offsetY = MediaTransformSeed.SeededInteger(seed, 0, removedY, 12);
        return new CropPlan(true, cropPermilleX, cropPermilleY, cropWidth, cropHeight, offsetX, offsetY);
    }

    private sealed record VideoDimensions(int Width, int Height, long QualityFloorVideoBitrateBps);

    /// <summary>
    /// Ladder rungs at or below the cap, widest first. When the cap sits below
    /// the smallest rung the cap itself is the single candidate.
    /// </summary>
    public static IReadOnlyList<int> VideoWidthLadderDescending(int maxWidthCap)
    {
        var cap = Math.Max(2, maxWidthCap);
        var rungs = V2VideoWidthLadder.Where(w => w <= cap).ToArray();
        return rungs.Length > 0 ? rungs : [cap];
    }

    private static VideoDimensions PlanV2VideoDimensions(
        int sourceWidth, int sourceHeight, double targetDurationSeconds, long outputByteLimit, int maxWidthCap)
    {
        var maxPayloadBps = (long)Math.Floor(outputByteLimit * V2VideoTargetRatio * 8 / targetDurationSeconds);
        var audioBps = PlanV2AudioBitrateBps(maxPayloadBps);
        var maxVideoBps = maxPayloadBps - audioBps;
        // Preset MaxWidth caps the ladder (a preset asking for 160px never emits 1080p).
        var cap = Math.Max(2, maxWidthCap);
        var ladder = V2VideoWidthLadder.Where(w => w <= cap).ToArray();
        if (ladder.Length == 0)
        {
            var (width, height) = ScaledDimensions(sourceWidth, sourceHeight, cap);
            return new VideoDimensions(width, height, QualityFloorVideoBitrateBps(width, height));
        }

        foreach (var maxWidth in ladder)
        {
            var (width, height) = ScaledDimensions(sourceWidth, sourceHeight, maxWidth);
            var floor = QualityFloorVideoBitrateBps(width, height);
            if (floor <= maxVideoBps)
            {
                return new VideoDimensions(width, height, floor);
            }
        }

        throw new VideoQualityFloorException(
            "The original-size limit is too small to produce a usable H.264 video for this duration.");
    }

    private static (int Width, int Height) ScaledDimensions(int sourceWidth, int sourceHeight, int maxWidth)
    {
        var width = Math.Max(2, Math.Min(sourceWidth, maxWidth) / 2 * 2);
        var height = Math.Max(2, (int)Math.Floor(width / (double)sourceWidth * sourceHeight) / 2 * 2);
        return (width, height);
    }

    private static long QualityFloorVideoBitrateBps(int width, int height)
    {
        // AwayFromZero like JS Math.round (C# default is banker's rounding).
        var scaled = (long)Math.Round(2_000_000.0 * width * height / V2ReferencePixels, MidpointRounding.AwayFromZero);
        return Math.Clamp(scaled, 100_000, 5_000_000);
    }

    private static long PlanV2AudioBitrateBps(long totalBitrateBps) => totalBitrateBps switch
    {
        >= 640_000 => 128_000,
        >= 384_000 => 96_000,
        >= 224_000 => 64_000,
        >= 128_000 => 48_000,
        >= 64_000 => 32_000,
        >= 17_000 => 16_000,
        _ => 0,
    };

    private static long TotalBitrateForVideoFloorBps(long videoFloorBps)
    {
        var total = videoFloorBps;
        for (var attempt = 0; attempt < 8; attempt++)
        {
            var next = videoFloorBps + PlanV2AudioBitrateBps(total);
            if (next == total)
            {
                return next;
            }

            total = next;
        }

        return total;
    }
}
