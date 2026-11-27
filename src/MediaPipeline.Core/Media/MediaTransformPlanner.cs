using System.Globalization;
using MediaPipeline.Core.Configuration;

namespace MediaPipeline.Core.Media;

/// <summary>
/// Deterministic recrop, microtrim, and color variation. Video encoding preserves
/// resolution within MaxWidth and uses quality settings independently of source size.
/// </summary>
public static class MediaTransformPlanner
{
    public const string ImageRecropV2Profile = "image-recrop-v2";
    public const string VideoMicrotrimQualityProfile = "video-microtrim-quality-v3";
    public const string ImageRecropEnhancedV2Profile = "image-recrop-enhanced-v2";
    public const string VideoMicrotrimEnhancedQualityProfile = "video-microtrim-enhanced-quality-v3";

    // 100-nit SDR reference. Hable compresses HDR highlights into SDR without
    // clipping skin and bright backgrounds. Drop inherited HDR/Dolby Vision data.
    private static string HdrToSdrFilter(int width, int height) =>
        $"zscale=w={width}:h={height}:filter=lanczos:transfer=linear:npl=100," +
        "format=gbrpf32le,zscale=primaries=bt709," +
        "tonemap=tonemap=hable:desat=0," +
        "zscale=transfer=bt709:matrix=bt709:range=limited:dither=error_diffusion," +
        "format=yuv420p,sidedata=mode=delete";

    public const long MaxImageSourceBytes = 100L * 1024 * 1024;
    public const long MaxVideoSourceBytes = 2L * 1024 * 1024 * 1024;
    public const long MaxImagePreparedBytes = 50L * 1024 * 1024;
    public const int MaxSourceDimension = 16_384;
    public const long MaxSourcePixels = 60_000_000;

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
        bool enhanced,
        bool autoRotate = true)
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

        var args = new List<string>(ImageInputArgs(sourcePath, autoRotate))
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
        PresetOptions preset,
        MediaInfo? sourceColor = null)
    {
        if (!double.IsFinite(durationSeconds) || durationSeconds <= 0.1 || probeWidth < 1 || probeHeight < 1)
        {
            throw new InvalidDataException("Video transform source has no usable duration.");
        }

        // long: 1h of video is 3.6e9µs, past int.MaxValue (~35.8min). Matches the
        // reference (JS number) without overflow.
        var enhanced = preset.EnhancedVariation;
        var durationMicroseconds = (long)Math.Floor(durationSeconds * 1_000_000);
        var canTrim = durationMicroseconds >= 340_000;
        var trimFloor = enhanced ? 20_000 : 10_000;
        const int trimRange = 30_000;
        var trimMicroseconds = canTrim
            ? trimFloor + ((MediaTransformSeed.SeededInteger(seed, 0, trimRange, 0) + ordinal) % (trimRange + 1))
            : 0;
        var targetDuration = Math.Max(0.1, durationSeconds - trimMicroseconds / 1_000_000.0);

        var (displayWidth, displayHeight) = DisplayDimensions(probeWidth, probeHeight, rotationDegrees);
        var dimensions = ScaledDimensions(displayWidth, displayHeight, Math.Max(2, preset.MaxWidth));
        var crf = Math.Clamp(preset.Crf, 0, 51);
        var hardwareVideoBitrateKbps = Math.Max(200, preset.VideoToolboxBitrateKbps);

        var brightnessRange = enhanced ? 2_000 : 1_200;
        var contrastRange = enhanced ? 5_000 : 2_500;
        var saturationRange = enhanced ? 6_000 : 3_000;
        var gammaRange = enhanced ? 2_500 : 1_500;
        var brightness = MediaTransformSeed.SeededSignedFraction(seed, brightnessRange, 1);
        var contrast = 1 + MediaTransformSeed.SeededSignedFraction(seed, contrastRange, 2);
        var gamma = 1 + MediaTransformSeed.SeededSignedFraction(seed, gammaRange, 3);
        var saturation = 1 + MediaTransformSeed.SeededSignedFraction(seed, saturationRange, 4);
        var crop = enhanced ? PlanV2Crop(seed, displayWidth, displayHeight, true) : default;
        var toneMapHdr = sourceColor?.IsHdr == true;

        var videoFilter = string.Join(",", new[]
        {
            crop.Cropped ? $"crop={crop.CropWidth}:{crop.CropHeight}:{crop.OffsetX}:{crop.OffsetY}" : null,
            // Keep HDR precision through linearization and gamut conversion. Merely
            // tagging HLG/PQ pixels as BT.709 leaves uploads washed out or clipped.
            toneMapHdr ? HdrToSdrFilter(dimensions.Width, dimensions.Height)
                : $"scale={dimensions.Width}:{dimensions.Height}:flags=lanczos",
            $"eq=brightness={brightness:F6}:contrast={contrast:F6}:gamma={gamma:F6}:saturation={saturation:F6}",
        }.Where(part => part is not null));

        var x264Params = enhanced
            ? $"keyint={MediaTransformSeed.SeededInteger(seed, 190, 230, 7)}:min-keyint={MediaTransformSeed.SeededInteger(seed, 24, 30, 8)}:scenecut=35:ref=4:bframes=4:aq-mode=3:deblock=-1,-1"
            : $"keyint={MediaTransformSeed.SeededInteger(seed, 210, 250, 7)}:min-keyint={MediaTransformSeed.SeededInteger(seed, 24, 30, 8)}:scenecut=40:ref=3:bframes=3:aq-mode=2";

        var videoArgs = new List<string>
        {
            "-c:v", hardwareVideoToolbox ? "h264_videotoolbox" : "libx264",
        };
        if (!hardwareVideoToolbox)
        {
            videoArgs.AddRange(["-preset", preset.X264Preset, "-crf", crf.ToString(CultureInfo.InvariantCulture)]);
        }
        else
        {
            videoArgs.AddRange(["-b:v", $"{hardwareVideoBitrateKbps}k"]);
        }
        videoArgs.AddRange(["-vf", videoFilter, "-pix_fmt", "yuv420p"]);
        if (toneMapHdr)
        {
            videoArgs.AddRange([
                "-colorspace", "bt709", "-color_primaries", "bt709",
                "-color_trc", "bt709", "-color_range", "tv",
            ]);
        }
        if (!hardwareVideoToolbox)
        {
            videoArgs.AddRange(["-x264-params", x264Params]);
        }

        var args = new List<string>
        {
            "-y", "-hide_banner", "-loglevel", "error",
            "-i", sourcePath,
            "-t", targetDuration.ToString("F6", CultureInfo.InvariantCulture),
            "-map", "0:v:0", "-map", "0:a:0?",
        };
        args.AddRange(videoArgs);
        args.AddRange(["-c:a", "aac", "-b:a", preset.AudioBitrate]);

        args.AddRange(["-movflags", "+faststart", "-map_metadata", "-1", outputPath]);

        var profile = hardwareVideoToolbox
            ? "video-microtrim-videotoolbox-v2"
            : enhanced ? VideoMicrotrimEnhancedQualityProfile : VideoMicrotrimQualityProfile;
        if (toneMapHdr) { profile += "-hdr-sdr-v1"; }

        var evidence = new Dictionary<string, object?>
        {
            ["profile"] = profile,
            ["variationPreset"] = enhanced ? "enhanced" : "default",
            ["encoder"] = hardwareVideoToolbox ? "h264_videotoolbox" : "libx264",
            ["preset"] = hardwareVideoToolbox ? null : preset.X264Preset,
            ["rateControl"] = hardwareVideoToolbox ? "bitrate" : "crf",
            ["crf"] = hardwareVideoToolbox ? (int?)null : crf,
            ["audioBitrate"] = preset.AudioBitrate,
            ["videoBitrateBps"] = hardwareVideoToolbox ? hardwareVideoBitrateKbps * 1_000L : (long?)null,
            ["x264Parameters"] = hardwareVideoToolbox ? null : x264Params,
            ["sourceByteCount"] = sourceByteCount,
            ["sourceColorTransfer"] = sourceColor?.ColorTransfer,
            ["sourceColorPrimaries"] = sourceColor?.ColorPrimaries,
            ["hdrToneMapped"] = toneMapHdr,
            ["outputColorSpace"] = toneMapHdr ? "bt709" : sourceColor?.ColorSpace,
            ["sourceWidth"] = probeWidth,
            ["sourceHeight"] = probeHeight,
            ["displayWidth"] = displayWidth,
            ["displayHeight"] = displayHeight,
            ["rotationDegrees"] = rotationDegrees ?? 0,
            ["outputWidth"] = dimensions.Width,
            ["outputHeight"] = dimensions.Height,
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
            profile, ".mp4", args, targetDuration, evidence);
    }

    private static IReadOnlyList<string> ImageInputArgs(string sourcePath, bool autoRotate) =>
        autoRotate
            ? ["-y", "-hide_banner", "-loglevel", "error", "-i", sourcePath, "-frames:v", "1", "-map_metadata", "-1"]
            : ["-y", "-hide_banner", "-loglevel", "error", "-noautorotate", "-i", sourcePath, "-frames:v", "1", "-map_metadata", "-1"];

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

    private static (int Width, int Height) ScaledDimensions(int sourceWidth, int sourceHeight, int maxWidth)
    {
        var width = Math.Max(2, Math.Min(sourceWidth, maxWidth) / 2 * 2);
        var height = Math.Max(2, (int)Math.Floor(width / (double)sourceWidth * sourceHeight) / 2 * 2);
        return (width, height);
    }

}
