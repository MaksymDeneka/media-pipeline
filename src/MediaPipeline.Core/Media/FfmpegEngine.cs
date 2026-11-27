using System.Globalization;
using MediaPipeline.Core.Configuration;
using MediaPipeline.Core.IO;
using MediaPipeline.Core.Tools;

namespace MediaPipeline.Core.Media;

public sealed record PreparedSource(
    string SourcePath,
    string ProcessingPath,
    string? TemporaryPath,
    string DetectedKind,
    string SourceHash);

public sealed record CreatedVariant(
    string Path,
    MediaKind Kind,
    int TrimMs,
    double DurationSeconds,
    MediaInfo MediaInfo,
    string Profile,
    string Seed,
    int? SourceWidth = null,
    int? SourceHeight = null);

/// <summary>
/// Media engine with deterministic transforms and quality-preserving video encoding.
/// Deterministic seeded recipes, no random, no NVENC/AMF, no segmentation.
/// - Image: rotation-aware recrop 2-5‰ (enhanced 4-8‰) + lanczos + eq(brightness/contrast/gamma/saturation).
/// - Video: microtrim 10-40ms (enhanced 20-50ms) + eq, original resolution capped by
///   MaxWidth, libx264 CRF or configured VideoToolbox bitrate, with decode verification.
/// FFmpeg drops input metadata. ExifTool strips JPEG metadata after encoding.
/// Media kind for limits and routing comes from probed content (see MediaProbe.DetectIdentity),
/// never from the filename.
/// </summary>
public sealed class FfmpegEngine(
    Toolchain tools,
    VideoEncoder encoder)
{
    private readonly MediaProbe _probe = new(tools);

    public VideoEncoder Encoder { get; } = encoder;
    public bool IsVideoToolbox => Encoder.Name == "h264_videotoolbox";

    public async Task<PreparedSource> PrepareAsync(
        PresetOptions preset,
        LanePaths lane,
        string sourcePath,
        MediaKind kind,
        CancellationToken cancellationToken = default)
    {
        // Heatup validateRetainedMediaSource: probe + hash happen in parallel,
        // enforce kind byte limits + decode dimensions, decode check one frame.
        // Kind comes from probed content, never from the filename: bytes win.
        // IO/parse failures are data failures (per-file isolation), never infra:
        // wrap them so a vanishing or truncated input cannot discard siblings.
        // The whole validation phase gets its own budget (heatup: 30s probe, 2min
        // validate): a hung probe/decode/hash must fail the file, not the lane.
        var outerToken = cancellationToken;
        using var validationTimeout = CancellationTokenSource.CreateLinkedTokenSource(outerToken);
        validationTimeout.CancelAfter(TimeSpan.FromMinutes(5));
        cancellationToken = validationTimeout.Token;

        var probeTask = _probe.ReadAsync(sourcePath, cancellationToken);
        var hashTask = MediaTransformSeed.HashFileAsync(sourcePath, cancellationToken);
        // If one sibling fails fast, stop and settle the other before throwing: a
        // background hash holding the file open can otherwise race the failed-move
        // (sharing violation on Windows) or burn I/O pointlessly. The settle itself
        // is bounded: a sibling hung in a synchronous open is abandoned (its fault
        // observed so it cannot crash the process) rather than wedging the lane.
        async Task QuiesceValidationAsync()
        {
            try
            {
                validationTimeout.Cancel();
            }
            catch (ObjectDisposedException)
            {
            }

            var settled = Task.WhenAll(probeTask, hashTask);
            if (await Task.WhenAny(settled, Task.Delay(TimeSpan.FromSeconds(30), outerToken)) != settled)
            {
                _ = settled.ContinueWith(
                    static task => _ = task.Exception,
                    CancellationToken.None,
                    TaskContinuationOptions.OnlyOnFaulted | TaskContinuationOptions.ExecuteSynchronously,
                    TaskScheduler.Default);
                return;
            }

            try
            {
                await settled;
            }
            catch
            {
            }
        }

        MediaInfo info;
        string sourceHash;
        try
        {
            await Task.WhenAll(probeTask, hashTask);
            info = await probeTask;
            sourceHash = await hashTask;
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or InvalidDataException)
        {
            await QuiesceValidationAsync();
            throw new InvalidDataException(
                $"Media preparation source is unreadable: '{sourcePath}'. {exception.Message}", exception);
        }
        catch (OperationCanceledException exception) when (!outerToken.IsCancellationRequested)
        {
            await QuiesceValidationAsync();
            throw new InvalidDataException(
                $"Media preparation validation timed out after 5 minutes: '{sourcePath}'.", exception);
        }

        if (info.Width is null or <= 0 || info.Height is null or <= 0)
        {
            throw new InvalidDataException($"Media preparation source has no readable video frame: '{sourcePath}'.");
        }

        if (info.Width > MediaTransformPlanner.MaxSourceDimension ||
            info.Height > MediaTransformPlanner.MaxSourceDimension ||
            (long)info.Width.Value * info.Height.Value > MediaTransformPlanner.MaxSourcePixels)
        {
            throw new InvalidDataException($"Media preparation source exceeds its safe decode dimensions: '{sourcePath}'.");
        }

        var (detectedKind, _) = MediaProbe.DetectIdentity(info);
        long byteCount;
        try
        {
            byteCount = new FileInfo(sourcePath).Length;
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            throw new InvalidDataException(
                $"Media preparation source is unreadable: '{sourcePath}'. {exception.Message}", exception);
        }

        var limit = detectedKind == "video"
            ? MediaTransformPlanner.MaxVideoSourceBytes
            : MediaTransformPlanner.MaxImageSourceBytes;
        if (byteCount < 1 || byteCount > limit)
        {
            throw new InvalidDataException($"Media preparation source exceeds its kind-specific byte limit: '{sourcePath}'.");
        }

        if (detectedKind == "video")
        {
            var duration = info.DurationSeconds;
            if (duration is null or <= 0.1)
            {
                throw new InvalidDataException($"Video preparation source has no usable duration: '{sourcePath}'.");
            }
        }

        string? normalizedPath = null;
        try
        {
            // FFmpeg exposes tiled HEIC photos as independent video streams.
            // Decode the complete primary image once, including its orientation,
            // rather than sending the first tile through the variant recipe.
            if (info.IsHeif && (OperatingSystem.IsMacOS() || info.HasTileGrid))
            {
                normalizedPath = Path.Combine(Path.GetTempPath(), $"media-pipeline-heic-{Guid.NewGuid():N}.png");
                if (OperatingSystem.IsMacOS())
                {
                    var normalizationArgs = new List<string> { "-s", "format", "png" };
                    // sips preserves orientation metadata in PNG, while FFmpeg
                    // ignores that PNG tag. Bake the HEIF display rotation into
                    // pixels explicitly. FFprobe reports counterclockwise angles.
                    if (info.RotationDegrees is > 0)
                        normalizationArgs.AddRange(["--rotate", (360 - info.RotationDegrees.Value).ToString(CultureInfo.InvariantCulture)]);
                    normalizationArgs.AddRange([sourcePath, "--out", normalizedPath]);
                    await RunRequiredAsync("/usr/bin/sips",
                        normalizationArgs, cancellationToken);
                }
                else
                {
                    var converter = Toolchain.FindOptional("heif-convert")
                        ?? throw new InvalidOperationException("Tiled HEIF images require libheif's heif-convert tool.");
                    await RunRequiredAsync(converter, [sourcePath, normalizedPath], cancellationToken);
                }

                var normalized = await _probe.ReadAsync(normalizedPath, cancellationToken);
                if (normalized.Width is null or <= 0 or > MediaTransformPlanner.MaxSourceDimension ||
                    normalized.Height is null or <= 0 or > MediaTransformPlanner.MaxSourceDimension ||
                    (long)normalized.Width.Value * normalized.Height.Value > MediaTransformPlanner.MaxSourcePixels ||
                    normalized.HasTileGrid)
                {
                    throw new InvalidDataException("HEIF normalization did not produce a safe complete image.");
                }
                if (info.HasTileGrid &&
                    !((normalized.Width == info.Width && normalized.Height == info.Height) ||
                      (normalized.Width == info.Height && normalized.Height == info.Width)))
                {
                    throw new InvalidDataException("HEIF normalization dimensions do not match the complete tile grid.");
                }
            }

            // Validate the video frame only. Automatic stream selection can pick
            // an unsupported auxiliary audio track, such as iPhone APAC, even
            // when the first audio track used by the encoder is valid AAC.
            await RunRequiredAsync(
                tools.FFmpeg,
                ["-v", "error", "-i", normalizedPath ?? sourcePath, "-map", "0:v:0", "-frames:v", "1", "-f", "null", "-"],
                cancellationToken);

            return new PreparedSource(sourcePath, normalizedPath ?? sourcePath, normalizedPath, detectedKind, sourceHash);
        }
        catch (OperationCanceledException exception) when (!outerToken.IsCancellationRequested)
        {
            TryDelete(normalizedPath);
            throw new InvalidDataException(
                $"Media preparation validation timed out after 5 minutes: '{sourcePath}'.", exception);
        }
        catch
        {
            TryDelete(normalizedPath);
            throw;
        }
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
        string seed,
        int ordinal,
        CancellationToken cancellationToken = default,
        string? workDirectory = null)
    {
        Directory.CreateDirectory(outputDirectory);
        var info = await WithFileBudgetAsync(
            token => source.TemporaryPath is null
                ? _probe.ReadImageAsync(source.ProcessingPath, token)
                : _probe.ReadAsync(source.ProcessingPath, token),
            $"Image probe for '{source.SourcePath}'",
            cancellationToken);
        if (info.Width is null or <= 0 || info.Height is null or <= 0)
        {
            throw new InvalidDataException($"Could not read image dimensions from '{source.SourcePath}'.");
        }

        var (_, probedExtension) = MediaProbe.DetectIdentity(info);
        // Normalization uses lossless PNG internally; HEIF outputs remain JPEG.
        var sourceExtension = source.TemporaryPath is not null ? ".jpg" : probedExtension;
        var planned = MediaTransformPlanner.PlanImage(
            seed,
            source.ProcessingPath,
            "__planning__",
            sourceExtension,
            info.Width.Value,
            info.Height.Value,
            info.RotationDegrees,
            preset.EnhancedVariation,
            autoRotate: source.TemporaryPath is null);

        string? exifTool = null;
        string? stagingDirectory = null;
        if (planned.OutputExtension is ".jpg" or ".jpeg")
        {
            exifTool = tools.ExifTool ?? throw new InvalidOperationException("ExifTool is required to create JPEG outputs.");
            stagingDirectory = workDirectory is null
                ? Directory.CreateTempSubdirectory("media-pipeline-jpeg-").FullName
                : Directory.CreateDirectory(Path.Combine(workDirectory, $"jpeg-{Guid.NewGuid():N}")).FullName;
        }

        var encodedPath = stagingDirectory is null
            ? OutputNameGenerator.NewFilePath(outputDirectory, planned.OutputExtension)
            : Path.Combine(stagingDirectory, "image" + planned.OutputExtension);

        try
        {
            AssertOutputExtension(encodedPath, planned.OutputExtension);
            var args = ReplaceExactPath(planned.Args, "__planning__", encodedPath);
            var fenced = InsertFsFence(args, planned.OutputByteLimit, encodedPath);
            // Per-invocation budgets like heatup's per-command timeouts.
            await RunBoundedAsync(
                "Image transform", fenced, TimeSpan.FromMinutes(5),
                source.ProcessingPath, cancellationToken);
            if (exifTool is not null)
            {
                // MJPEG adds an encoder comment even with -map_metadata -1.
                await RunBoundedAsync(
                    "JPEG metadata removal", ["-config", "", "-all=", "-overwrite_original", encodedPath],
                    TimeSpan.FromMinutes(5), encodedPath, cancellationToken, executable: exifTool);
            }

            var outputStats = new FileInfo(encodedPath);
            if (!outputStats.Exists || outputStats.Length <= 0 || outputStats.Length > planned.OutputByteLimit)
            {
                throw new InvalidOperationException("Media transform output is empty, unreadable, or exceeds its size limit.");
            }

            var outputProbe = await WithFileBudgetAsync(
                token => _probe.ReadAsync(encodedPath, token),
                $"Image output probe for '{encodedPath}'",
                cancellationToken);
            if (outputProbe.Width is null or <= 0 || outputProbe.Height is null or <= 0)
            {
                throw new InvalidOperationException("Media transform output has no readable video frame.");
            }
            var expected = info.RotationDegrees is 90 or 270
                ? (Width: info.Height.Value, Height: info.Width.Value)
                : (Width: info.Width.Value, Height: info.Height.Value);
            if (outputProbe.Width != expected.Width || outputProbe.Height != expected.Height)
            {
                throw new InvalidOperationException(
                    $"Image output dimensions {outputProbe.Width}x{outputProbe.Height} do not match the complete source {expected.Width}x{expected.Height}.");
            }

            await RunBoundedAsync(
                "Image decode check", ["-v", "error", "-i", encodedPath, "-map", "0:v:0", "-f", "null", "-"],
                TimeSpan.FromMinutes(5), encodedPath, cancellationToken);

            var outputHash = await WithFileBudgetAsync(
                token => MediaTransformSeed.HashFileAsync(encodedPath, token),
                $"Image output hash for '{encodedPath}'",
                cancellationToken);
            if (string.Equals(source.SourceHash, outputHash, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("Media transform output is byte-identical to its immutable source.");
            }

            var outputPath = encodedPath;
            if (stagingDirectory is not null)
            {
                cancellationToken.ThrowIfCancellationRequested();
                outputPath = OutputNameGenerator.NewFilePath(outputDirectory, planned.OutputExtension);
                File.Move(encodedPath, outputPath);
            }

            return new CreatedVariant(outputPath, MediaKind.Image, 0, 0, outputProbe, planned.Profile, seed, info.Width, info.Height);
        }
        catch
        {
            if (stagingDirectory is null)
            {
                TryDelete(encodedPath);
            }
            throw;
        }
        finally
        {
            if (stagingDirectory is not null)
            {
                try
                {
                    Directory.Delete(stagingDirectory, recursive: true);
                }
                catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
                {
                }
            }
        }
    }

    public async Task<CreatedVariant> CreateVideoVariantAsync(
        string inputPath,
        string outputDirectory,
        PresetOptions preset,
        double sourceDurationSeconds,
        long sourceByteCount,
        string seed,
        int ordinal,
        string? sourceHash = null,
        CancellationToken cancellationToken = default)
    {
        Directory.CreateDirectory(outputDirectory);
        var probe = await WithFileBudgetAsync(
            token => _probe.ReadAsync(inputPath, token),
            $"Video input probe for '{inputPath}'",
            cancellationToken);
        if (probe.Width is null or <= 0 || probe.Height is null or <= 0)
        {
            throw new InvalidDataException($"Could not read video dimensions from '{inputPath}'.");
        }

        var outputPath = OutputNameGenerator.NewFilePath(outputDirectory, ".mp4");
        sourceHash ??= await WithFileBudgetAsync(
            token => MediaTransformSeed.HashFileAsync(inputPath, token),
            $"Video source hash for '{inputPath}'",
            cancellationToken);
        var planned = MediaTransformPlanner.PlanVideo(
            seed, ordinal, inputPath, outputPath,
            probe.Width.Value, probe.Height.Value, probe.RotationDegrees,
            sourceDurationSeconds, sourceByteCount, IsVideoToolbox, preset, probe);
        return await ExecuteVideoPlanAsync(
            planned, inputPath, outputPath, seed, sourceHash, cancellationToken);
    }

    private async Task<CreatedVariant> ExecuteVideoPlanAsync(
        MediaTransformPlanner.PlannedVideoTransform planned,
        string inputPath,
        string outputPath,
        string seed,
        string sourceHash,
        CancellationToken cancellationToken)
    {
        AssertOutputExtension(outputPath, planned.OutputExtension);
        var trimMs = planned.Evidence.TryGetValue("trimMicroseconds", out var trimObj) switch
        {
            { } when trimObj is long trimUsLong => (int)(trimUsLong / 1000),
            { } when trimObj is int trimUsInt => trimUsInt / 1000,
            _ => 0,
        };

        try
        {
            await RunBoundedAsync(
                "Video transform", planned.Args, TimeSpan.FromMinutes(30), inputPath, cancellationToken);
            var outputStats = new FileInfo(outputPath);
            if (!outputStats.Exists || outputStats.Length <= 0)
            {
                throw new InvalidOperationException("Media transform output is empty or unreadable.");
            }
            var outputProbe = await VerifyVideoProbeAndDurationAsync(outputPath, planned, cancellationToken);
            await RunBoundedAsync(
                "Video decode check",
                ["-v", "error", "-i", outputPath, "-map", "0:v:0", "-f", "null", "-"],
                TimeSpan.FromMinutes(5), outputPath, cancellationToken);

            var outputHash = await WithFileBudgetAsync(
                token => MediaTransformSeed.HashFileAsync(outputPath, token),
                $"Video output hash for '{outputPath}'",
                cancellationToken);
            if (string.Equals(sourceHash, outputHash, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("Media transform output is byte-identical to its immutable source.");
            }

            int? sourceWidth = planned.Evidence.TryGetValue("sourceWidth", out var sourceWidthObj) && sourceWidthObj is int sw
                ? sw
                : null;
            int? sourceHeight = planned.Evidence.TryGetValue("sourceHeight", out var sourceHeightObj) && sourceHeightObj is int sh
                ? sh
                : null;
            return new CreatedVariant(
                outputPath, MediaKind.Video, trimMs, planned.TargetDurationSeconds,
                outputProbe, planned.Profile, seed, sourceWidth, sourceHeight);
        }
        catch
        {
            TryDelete(outputPath);
            throw;
        }
    }

    /// <summary>
    /// Runs one media tool invocation with its own budget. A timeout is a per-file
    /// data failure, never a worker stop: it converts only when the outer token
    /// is not cancelled, so genuine shutdown still flows as cancellation.
    /// </summary>
    private async Task RunBoundedAsync(
        string description,
        IEnumerable<string> arguments,
        TimeSpan budget,
        string pathForMessage,
        CancellationToken outerToken,
        string? executable = null)
    {
        using var budgetTimeout = CancellationTokenSource.CreateLinkedTokenSource(outerToken);
        budgetTimeout.CancelAfter(budget);
        try
        {
            await RunRequiredAsync(executable ?? tools.FFmpeg, arguments, budgetTimeout.Token);
        }
        catch (OperationCanceledException exception) when (!outerToken.IsCancellationRequested)
        {
            throw new InvalidOperationException(
                $"{description} timed out after {budget.TotalMinutes:g} minute(s): '{pathForMessage}'.", exception);
        }
    }

    /// <summary>
    /// Bounds a variant-phase file read (probe/hash) that otherwise runs on the
    /// unbounded worker token: a hung network file must fail the file, not wedge
    /// the lane. Timeouts convert to data failures; genuine shutdown still flows
    /// as cancellation.
    /// </summary>
    private static async Task<T> WithFileBudgetAsync<T>(
        Func<CancellationToken, Task<T>> operation,
        string description,
        CancellationToken outerToken)
    {
        using var budgetTimeout = CancellationTokenSource.CreateLinkedTokenSource(outerToken);
        budgetTimeout.CancelAfter(TimeSpan.FromMinutes(5));
        try
        {
            return await operation(budgetTimeout.Token);
        }
        catch (OperationCanceledException exception) when (!outerToken.IsCancellationRequested)
        {
            throw new InvalidDataException($"{description} timed out after 5 minutes.", exception);
        }
    }

    private async Task<MediaInfo> VerifyVideoProbeAndDurationAsync(
        string outputPath,
        MediaTransformPlanner.PlannedVideoTransform planned,
        CancellationToken cancellationToken)
    {
        var outputProbe = await WithFileBudgetAsync(
            token => _probe.ReadAsync(outputPath, token),
            $"Video output probe for '{outputPath}'",
            cancellationToken);
        if (outputProbe.Width is null or <= 0 || outputProbe.Height is null or <= 0)
        {
            throw new InvalidOperationException("Media transform output has no readable video frame.");
        }
        if (outputProbe.Width != (int)planned.Evidence["outputWidth"]! ||
            outputProbe.Height != (int)planned.Evidence["outputHeight"]!)
        {
            throw new InvalidOperationException("Media transform video output does not match its planned resolution.");
        }
        if (planned.Evidence.TryGetValue("hdrToneMapped", out var toneMapped) && toneMapped is true &&
            (outputProbe.ColorSpace != "bt709" || outputProbe.ColorPrimaries != "bt709" ||
             outputProbe.ColorTransfer != "bt709" || outputProbe.ColorRange != "tv" ||
             outputProbe.PixelFormat != "yuv420p"))
        {
            throw new InvalidOperationException("HDR conversion output is not limited-range BT.709 SDR H.264.");
        }

        var targetDuration = planned.TargetDurationSeconds;
        var minimumDuration = targetDuration - Math.Max(0.25, targetDuration * 0.02);
        var outputDuration = outputProbe.DurationSeconds;
        if (!double.IsFinite(targetDuration) ||
            outputDuration is null ||
            outputDuration < minimumDuration ||
            outputDuration > targetDuration + 1)
        {
            throw new InvalidOperationException("Media transform video output is truncated or has an invalid duration.");
        }

        return outputProbe;
    }

    public async Task<IReadOnlyList<(SegmentPlan Plan, string Path)>> ExtractSegmentsAsync(
        PreparedSource source,
        LanePaths lane,
        PresetOptions preset,
        double durationSeconds,
        CancellationToken cancellationToken = default)
    {
        // Keep the full clip; video encoding does not split or automatically downscale it.
        await Task.CompletedTask;
        return [(new SegmentPlan(1, -1, durationSeconds), source.ProcessingPath)];
    }

    public async Task<double> DurationAsync(string path, CancellationToken cancellationToken = default)
    {
        var info = await WithFileBudgetAsync(
            token => _probe.ReadAsync(path, token),
            $"Duration probe for '{path}'",
            cancellationToken);
        return info.DurationSeconds is > 0
            ? info.DurationSeconds.Value
            : throw new InvalidDataException($"Could not read a valid duration from '{path}'.");
    }

    public Task<bool> RecompressIfOversizedAsync(
        string path,
        PresetOptions preset,
        CancellationToken cancellationToken = default)
    {
        // Keep CLI compatibility without applying another lossy compression pass.
        return Task.FromResult(false);
    }

    private static List<string> ReplaceExactPath(IReadOnlyList<string> args, string oldValue, string newValue)
    {
        // Exact-element match only: placeholders are full argv elements, and an
        // input path that happens to contain the placeholder text must never be
        // rewritten.
        var result = new List<string>(args.Count);
        foreach (var arg in args)
        {
            result.Add(arg.Equals(oldValue, StringComparison.Ordinal) ? newValue : arg);
        }

        return result;
    }

    private static List<string> InsertFsFence(IReadOnlyList<string> args, long outputByteLimit, string outputPath)
    {
        // Heatup appends -fs <limit> before the output path. The planner always
        // ends argv with the output path; a missing placeholder is a planning bug
        // and must fail loudly instead of emitting "-fs <limit>" as a stray output.
        var result = new List<string>(args);
        var index = result.LastIndexOf(outputPath);
        if (index < 0)
        {
            throw new InvalidOperationException("Media output placeholder not found in planned ffmpeg argv.");
        }

        result.Insert(index, outputByteLimit.ToString(CultureInfo.InvariantCulture));
        result.Insert(index, "-fs");
        return result;
    }

    private static void AssertOutputExtension(string outputPath, string plannedExtension)
    {
        // OutputNameGenerator emits uppercase (IMG_0001.MP4); compare case-insensitively
        // like the reference assertTransformOutputExtension, modulo filesystem case.
        if (!Path.GetExtension(outputPath).Equals(plannedExtension, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException(
                $"Media transform output extension '{Path.GetExtension(outputPath)}' does not match its recipe ('{plannedExtension}').");
        }
    }

    private async Task RunRequiredAsync(
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
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or ArgumentException)
        {
        }
    }
}
