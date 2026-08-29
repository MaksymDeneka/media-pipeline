using System.Globalization;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using MediaPipeline.Core.Configuration;
using MediaPipeline.Core.IO;
using MediaPipeline.Core.Tools;

namespace MediaPipeline.Core.Media;

public sealed record PreparedSource(string SourcePath, string ProcessingPath, string? TemporaryPath);

public sealed record CreatedVariant(
    string Path,
    MediaKind Kind,
    int TrimMs,
    double DurationSeconds,
    MediaInfo MediaInfo);

public sealed class FfmpegEngine(
    Toolchain tools,
    VideoEncoder encoder,
    VideoOptions videoOptions)
{
    private readonly MediaProbe _probe = new(tools);

    public VideoEncoder Encoder { get; } = encoder;

    public async Task<PreparedSource> PrepareAsync(
        PresetOptions preset,
        LanePaths lane,
        string sourcePath,
        MediaKind kind,
        CancellationToken cancellationToken = default)
    {
        if (kind == MediaKind.Image && await IsHeicAsync(sourcePath, cancellationToken))
        {
            var temporary = Path.Combine(
                Path.GetTempPath(),
                $"media-pipeline-heic-{Guid.NewGuid():n}.png");
            try
            {
                await RunRequiredAsync(
                    tools.FFmpeg,
                    [
                        "-y", "-hide_banner", "-loglevel", "error", "-threads", "1",
                        "-i", sourcePath, "-frames:v", "1", "-map_metadata", "-1",
                        "-threads", "1", temporary,
                    ],
                    cancellationToken);
                return new PreparedSource(sourcePath, temporary, temporary);
            }
            catch
            {
                TryDelete(temporary);
                throw;
            }
        }

        if (kind == MediaKind.Video && preset.Normalize && preset.Segment &&
            Path.GetExtension(sourcePath).Equals(".mov", StringComparison.OrdinalIgnoreCase))
        {
            Directory.CreateDirectory(lane.Work);
            var temporary = Path.Combine(
                lane.Work,
                Path.GetFileNameWithoutExtension(sourcePath) + $"-{Guid.NewGuid():n}.mp4");
            try
            {
                await RemuxAsync(sourcePath, temporary, cancellationToken);
                return new PreparedSource(sourcePath, temporary, temporary);
            }
            catch
            {
                TryDelete(temporary);
                throw;
            }
        }

        return new PreparedSource(sourcePath, sourcePath, null);
    }

    public static void RemoveTemporarySource(PreparedSource source)
    {
        if (source.TemporaryPath is not null)
        {
            TryDelete(source.TemporaryPath);
        }
    }

    public async Task<CreatedVariant> CreateImageVariantAsync(
        PreparedSource source,
        string outputDirectory,
        PresetOptions preset,
        CancellationToken cancellationToken = default)
    {
        Directory.CreateDirectory(outputDirectory);
        var info = await _probe.ReadAsync(source.ProcessingPath, cancellationToken);
        if (info.Width is null or <= 0 || info.Height is null or <= 0)
        {
            throw new InvalidDataException($"Could not read image dimensions from '{source.SourcePath}'.");
        }

        var extension = Path.GetExtension(source.SourcePath).Equals(".webp", StringComparison.OrdinalIgnoreCase)
            ? ".webp"
            : ".jpg";
        var outputPath = OutputNameGenerator.NewFilePath(outputDirectory, extension);
        var arguments = new List<string>
        {
            "-y", "-hide_banner", "-loglevel", "error", "-threads", "1",
            "-filter_complex_threads", "1", "-i", source.ProcessingPath,
            "-frames:v", "1", "-map_metadata", "-1",
        };

        if (info.Width >= 200 && info.Height >= 200)
        {
            var minCrop = Math.Min(preset.CropMinPermille, preset.CropMaxPermille);
            var maxCrop = Math.Max(preset.CropMinPermille, preset.CropMaxPermille);
            var cropPermille = RandomNumberGenerator.GetInt32(minCrop, maxCrop + 1);
            var cropPixelsX = Math.Max(1, (int)Math.Floor(info.Width.Value * cropPermille / 1000.0));
            var cropPixelsY = Math.Max(1, (int)Math.Floor(info.Height.Value * cropPermille / 1000.0));
            var cropWidth = Math.Max(1, info.Width.Value - cropPixelsX * 2);
            var cropHeight = Math.Max(1, info.Height.Value - cropPixelsY * 2);
            var offsetX = RandomNumberGenerator.GetInt32(cropPixelsX * 2 + 1);
            var offsetY = RandomNumberGenerator.GetInt32(cropPixelsY * 2 + 1);
            var filter =
                $"crop={cropWidth}:{cropHeight}:{offsetX}:{offsetY},scale={info.Width}:{info.Height}";
            arguments.AddRange(["-filter_complex", $"[0:v:0]{filter}[v]", "-map", "[v]"]);
        }

        if (extension == ".jpg")
        {
            var quality = source.TemporaryPath is null
                ? preset.JpegQuality
                : preset.ConvertedJpegQuality;
            arguments.AddRange(["-q:v", quality.ToString(CultureInfo.InvariantCulture)]);
        }
        else
        {
            arguments.AddRange(["-quality", "92"]);
        }

        arguments.AddRange(["-threads", "1", outputPath]);

        try
        {
            await RunRequiredAsync(tools.FFmpeg, arguments, cancellationToken);
            await ClearMetadataAsync(outputPath, cancellationToken);
            return new CreatedVariant(outputPath, MediaKind.Image, 0, 0, info);
        }
        catch
        {
            TryDelete(outputPath);
            throw;
        }
    }

    public async Task<CreatedVariant> CreateVideoVariantAsync(
        string inputPath,
        string outputDirectory,
        PresetOptions preset,
        double sourceDurationSeconds,
        int trimMs,
        CancellationToken cancellationToken = default)
    {
        Directory.CreateDirectory(outputDirectory);
        var outputPath = OutputNameGenerator.NewFilePath(outputDirectory, ".mp4");
        var targetDuration = Math.Max(0.1, sourceDurationSeconds - trimMs / 1000.0);
        var quality = Encoder.Name switch
        {
            "h264_nvenc" => preset.NvencCq,
            "h264_amf" => preset.AmfQp,
            _ => preset.Crf,
        };
        var maxBitrate = preset.SizeCapMB > 0
            ? PrimaryMaxVideoBitrate(
                targetDuration, preset.SizeCapMB, preset.AudioBitrate, preset.MaxrateScale)
            : 0;

        try
        {
            await EncodeVideoAsync(
                inputPath,
                outputPath,
                quality,
                preset.MaxWidth,
                preset.AudioBitrate,
                durationSeconds: targetDuration,
                maxVideoBitrateKbps: maxBitrate,
                cancellationToken: cancellationToken);
            await ClearMetadataAsync(outputPath, cancellationToken);

            if (preset.SizeCapMB > 0)
            {
                await EnforceSizeCapAsync(
                    outputPath,
                    preset.SizeCapMB,
                    preset.MaxWidth,
                    preset.SizeCapFallbackMaxWidth,
                    preset.AudioBitrate,
                    inputPath,
                    sourceDurationSeconds,
                    trimMs,
                    cancellationToken);
            }

            var info = await _probe.ReadAsync(outputPath, cancellationToken);
            return new CreatedVariant(outputPath, MediaKind.Video, trimMs, targetDuration, info);
        }
        catch
        {
            TryDelete(outputPath);
            throw;
        }
    }

    public async Task<IReadOnlyList<(SegmentPlan Plan, string Path)>> ExtractSegmentsAsync(
        PreparedSource source,
        LanePaths lane,
        PresetOptions preset,
        double durationSeconds,
        CancellationToken cancellationToken = default)
    {
        Directory.CreateDirectory(lane.Work);
        var plan = SegmentPlanner.Plan(
            durationSeconds,
            preset.SegmentTargetSeconds,
            preset.SegmentMinSeconds);
        var token = Guid.NewGuid().ToString("n")[..8];
        var results = new List<(SegmentPlan Plan, string Path)>(plan.Count);
        string? currentPath = null;

        try
        {
            foreach (var segment in plan)
            {
                var path = Path.Combine(lane.Work, $"segment_{token}_{segment.Index:D3}.mp4");
                currentPath = path;
                await RunRequiredAsync(
                    tools.FFmpeg,
                    [
                        "-y",
                        "-hide_banner",
                        "-loglevel",
                        "error",
                        "-ss",
                        Number(segment.StartSeconds),
                        "-i",
                        source.ProcessingPath,
                        "-t",
                        Number(segment.DurationSeconds),
                        "-map",
                        "0:v:0",
                        "-map",
                        "0:a:0?",
                        "-dn",
                        "-c",
                        "copy",
                        "-map_metadata",
                        "-1",
                        "-movflags",
                        "+faststart",
                        path,
                    ],
                    cancellationToken);
                results.Add((segment, path));
                currentPath = null;
            }

            return results;
        }
        catch
        {
            foreach (var result in results)
            {
                TryDelete(result.Path);
            }
            TryDelete(currentPath);

            throw;
        }
    }

    public async Task<double> DurationAsync(string path, CancellationToken cancellationToken = default)
    {
        var info = await _probe.ReadAsync(path, cancellationToken);
        return info.DurationSeconds is > 0
            ? info.DurationSeconds.Value
            : throw new InvalidDataException($"Could not read a valid duration from '{path}'.");
    }

    public async Task<bool> RecompressIfOversizedAsync(
        string path,
        PresetOptions preset,
        CancellationToken cancellationToken = default)
    {
        if (preset.SizeCapMB <= 0 || !File.Exists(path))
        {
            return false;
        }

        var maxBytes = (long)(preset.SizeCapMB * 1024 * 1024);
        if (new FileInfo(path).Length <= maxBytes)
        {
            return false;
        }

        var duration = await DurationAsync(path, cancellationToken);
        await EnforceSizeCapAsync(
            path,
            preset.SizeCapMB,
            preset.MaxWidth,
            preset.SizeCapFallbackMaxWidth,
            preset.AudioBitrate,
            path,
            duration,
            trimMs: 0,
            cancellationToken);
        return true;
    }

    private async Task RemuxAsync(
        string inputPath,
        string outputPath,
        CancellationToken cancellationToken)
    {
        await RunRequiredAsync(
            tools.FFmpeg,
            [
                "-y",
                "-hide_banner",
                "-loglevel",
                "error",
                "-i",
                inputPath,
                "-map",
                "0:v:0",
                "-map",
                "0:a:0?",
                "-dn",
                "-c",
                "copy",
                "-map_metadata",
                "-1",
                "-movflags",
                "+faststart",
                outputPath,
            ],
            cancellationToken);
    }

    private async Task<bool> IsHeicAsync(string path, CancellationToken cancellationToken)
    {
        if (MediaClassifier.IsHeic(path))
        {
            return true;
        }

        var result = await ProcessRunner.RunAsync(
            tools.FFprobe,
            [
                "-v",
                "error",
                "-show_entries",
                "format_tags=major_brand,compatible_brands",
                "-of",
                "default=noprint_wrappers=1:nokey=1",
                path,
            ],
            cancellationToken);
        return result.Succeeded && Regex.IsMatch(
            result.StandardOutput,
            @"(^|\s)(heic|heix|hevc|hevx|mif1|msf1)(\s|$)",
            RegexOptions.IgnoreCase);
    }

    private async Task EncodeVideoAsync(
        string inputPath,
        string outputPath,
        int quality,
        int maxWidth,
        string audioBitrate,
        double startSeconds = -1,
        double durationSeconds = -1,
        int maxVideoBitrateKbps = 0,
        CancellationToken cancellationToken = default)
    {
        var arguments = new List<string> { "-y", "-hide_banner", "-loglevel", "error" };
        if (startSeconds >= 0 && durationSeconds > 0)
        {
            arguments.AddRange(["-ss", Number(startSeconds), "-i", inputPath, "-t", Number(durationSeconds)]);
        }
        else
        {
            arguments.AddRange(["-i", inputPath]);
            if (durationSeconds > 0)
            {
                arguments.AddRange(["-t", Number(durationSeconds)]);
            }
        }

        arguments.AddRange(["-map", "0:v:0", "-map", "0:a:0?"]);
        arguments.AddRange(EncoderArguments(quality, maxWidth, maxVideoBitrateKbps));
        arguments.AddRange([
            "-c:a",
            "aac",
            "-b:a",
            audioBitrate,
            "-movflags",
            "+faststart",
            "-map_metadata",
            "-1",
            outputPath,
        ]);
        await RunRequiredAsync(tools.FFmpeg, arguments, cancellationToken);
    }

    private IReadOnlyList<string> EncoderArguments(int quality, int maxWidth, int maxVideoBitrateKbps)
    {
        var scale = $"scale='trunc(min({maxWidth},iw)/2)*2':-2";
        var arguments = Encoder.Name switch
        {
            "h264_nvenc" => new List<string>
            {
                "-c:v", "h264_nvenc", "-preset", videoOptions.NvencPreset,
                "-tune", "hq", "-rc", "vbr", "-cq", Number(quality), "-b:v", "0",
                "-spatial_aq", "1", "-temporal_aq", "1",
                "-vf", scale, "-pix_fmt", "yuv420p",
            },
            "h264_amf" when maxVideoBitrateKbps > 0 => new List<string>
            {
                "-c:v", "h264_amf", "-usage", "transcoding", "-quality", videoOptions.AmfQuality,
                "-rc", "vbr_peak", "-b:v", $"{maxVideoBitrateKbps}k",
                "-vf", scale, "-pix_fmt", "yuv420p",
            },
            "h264_amf" => new List<string>
            {
                "-c:v", "h264_amf", "-usage", "transcoding", "-quality", videoOptions.AmfQuality,
                "-rc", "cqp", "-qp_i", Number(quality), "-qp_p", Number(quality),
                "-qp_b", Number(quality), "-vf", scale, "-pix_fmt", "yuv420p",
            },
            "h264_videotoolbox" => new List<string>
            {
                "-c:v", "h264_videotoolbox", "-profile:v", "high", "-allow_sw", "0",
                "-b:v", $"{(maxVideoBitrateKbps > 0 ? maxVideoBitrateKbps : videoOptions.VideoToolboxBitrateKbps)}k",
                "-vf", scale, "-pix_fmt", "yuv420p",
            },
            _ => new List<string>
            {
                "-c:v", "libx264", "-crf", Number(quality), "-preset", videoOptions.X264Preset,
                "-vf", scale, "-pix_fmt", "yuv420p",
            },
        };

        if (maxVideoBitrateKbps > 0)
        {
            arguments.AddRange([
                "-maxrate",
                $"{maxVideoBitrateKbps}k",
                "-bufsize",
                $"{maxVideoBitrateKbps * 2}k",
            ]);
        }

        return arguments;
    }

    private int PrimaryMaxVideoBitrate(
        double durationSeconds,
        double maxSizeMB,
        string audioBitrate,
        double maxrateScale)
    {
        if (Encoder.Name == "libx264" || maxSizeMB <= 0 || durationSeconds <= 0)
        {
            return 0;
        }

        return Math.Max(200, (int)Math.Floor(
            TargetVideoBitrate(durationSeconds, maxSizeMB, audioBitrate) * maxrateScale));
    }

    private static int TargetVideoBitrate(
        double durationSeconds,
        double maxSizeMB,
        string audioBitrate)
    {
        var digits = Regex.Replace(audioBitrate, "[^0-9.]", "");
        var audioKbps = double.TryParse(digits, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed)
            ? parsed
            : 128;
        var totalKbps = maxSizeMB * 8192 / durationSeconds;
        var videoKbps = Math.Max(200, totalKbps - audioKbps);
        return (int)Math.Floor(videoKbps * 0.90);
    }

    private async Task EnforceSizeCapAsync(
        string outputPath,
        double maxSizeMB,
        int primaryMaxWidth,
        int fallbackMaxWidth,
        string audioBitrate,
        string sourceInputPath,
        double sourceDuration,
        int trimMs,
        CancellationToken cancellationToken)
    {
        var maxBytes = (long)(maxSizeMB * 1024 * 1024);
        if (new FileInfo(outputPath).Length <= maxBytes)
        {
            return;
        }

        var duration = Math.Max(0.1, sourceDuration - trimMs / 1000.0);
        var targetBitrate = TargetVideoBitrate(duration, maxSizeMB, audioBitrate);
        var profiles = SizeCapProfiles(primaryMaxWidth, fallbackMaxWidth, targetBitrate);
        string? chosenPath = null;
        long chosenSize = long.MaxValue;

        try
        {
            foreach (var profile in profiles)
            {
                var temporary = Path.Combine(
                    Path.GetDirectoryName(outputPath)!,
                    $"sizecap_{Convert.ToHexString(RandomNumberGenerator.GetBytes(8)).ToLowerInvariant()}.mp4");
                try
                {
                    await EncodeVideoAsync(
                        sourceInputPath,
                        temporary,
                        profile.Quality,
                        profile.MaxWidth,
                        audioBitrate,
                        durationSeconds: duration,
                        maxVideoBitrateKbps: profile.Bitrate,
                        cancellationToken: cancellationToken);
                    var size = new FileInfo(temporary).Length;
                    if (size < chosenSize)
                    {
                        if (chosenPath is not null)
                        {
                            TryDelete(chosenPath);
                        }

                        chosenPath = temporary;
                        chosenSize = size;
                        temporary = "";
                    }

                    if (size <= maxBytes)
                    {
                        break;
                    }
                }
                finally
                {
                    TryDelete(temporary);
                }
            }

            if (chosenPath is not null)
            {
                File.Move(chosenPath, outputPath, overwrite: true);
                chosenPath = null;
                await ClearMetadataAsync(outputPath, cancellationToken);
            }
        }
        finally
        {
            if (chosenPath is not null)
            {
                TryDelete(chosenPath);
            }
        }
    }

    private IReadOnlyList<(int Quality, int MaxWidth, int Bitrate)> SizeCapProfiles(
        int primaryMaxWidth,
        int fallbackMaxWidth,
        int targetBitrate)
    {
        var clampedFallbackWidth = Math.Min(primaryMaxWidth, fallbackMaxWidth);
        return Encoder.Name switch
        {
            "h264_nvenc" =>
            [
                (30, primaryMaxWidth, 0),
                (32, primaryMaxWidth, 0),
                (34, clampedFallbackWidth, 0),
                (36, clampedFallbackWidth, targetBitrate),
            ],
            "h264_amf" =>
            [
                (28, primaryMaxWidth, 0),
                (30, primaryMaxWidth, 0),
                (32, clampedFallbackWidth, 0),
                (34, clampedFallbackWidth, targetBitrate),
            ],
            "h264_videotoolbox" =>
            [
                (0, primaryMaxWidth, targetBitrate),
                (0, clampedFallbackWidth, Math.Max(200, (int)(targetBitrate * 0.92))),
            ],
            _ =>
            [
                (28, primaryMaxWidth, 0),
                (30, primaryMaxWidth, 0),
                (32, clampedFallbackWidth, 0),
                (32, clampedFallbackWidth, targetBitrate),
            ],
        };
    }

    private async Task ClearMetadataAsync(string path, CancellationToken cancellationToken)
    {
        await RunRequiredAsync(
            tools.ExifTool,
            ["-all=", "-overwrite_original", path],
            cancellationToken);
    }

    private static async Task RunRequiredAsync(
        string executable,
        IEnumerable<string> arguments,
        CancellationToken cancellationToken)
    {
        var argumentList = arguments.ToArray();
        var result = await ProcessRunner.RunAsync(executable, argumentList, cancellationToken);
        if (!result.Succeeded)
        {
            var detail = result.CombinedOutput.Trim();
            throw new InvalidOperationException(
                $"'{Path.GetFileName(executable)}' exited with {result.ExitCode}: {detail}");
        }
    }

    private static string Number(double value) => value.ToString("0.###", CultureInfo.InvariantCulture);

    private static void TryDelete(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return;
        }

        try
        {
            File.Delete(path);
        }
        catch (IOException)
        {
        }
    }
}
