using System.Diagnostics;
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
/// V2 media engine ported from heatup sidecar/media-transform (profiles.ts + executor.ts).
/// Deterministic seeded recipes, no random, no ExifTool, no NVENC/AMF, no segmentation.
/// - Image: rotation-aware recrop 2-5‰ (enhanced 4-8‰) + lanczos + eq(brightness/contrast/gamma/saturation).
/// - Video: microtrim 10-40ms (enhanced 20-50ms) + eq + width ladder (1080..160, even dims,
///   capped by preset MaxWidth) + quality floor + two-pass libx264 slow CBR or single-pass
///   VideoToolbox, -fs fence + verification.
/// Metadata is stripped with -map_metadata -1 only.
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

        try
        {
            await RunRequiredAsync(
                tools.FFmpeg,
                ["-v", "error", "-i", sourcePath, "-frames:v", "1", "-f", "null", "-"],
                cancellationToken);
        }
        catch (OperationCanceledException exception) when (!outerToken.IsCancellationRequested)
        {
            throw new InvalidDataException(
                $"Media preparation validation timed out after 5 minutes: '{sourcePath}'.", exception);
        }

        // No HEIC temp PNG, no MOV remux – ffmpeg handles normalization directly
        // via the planned output extension + filters, like heatup. The source hash
        // is captured once here so variants never re-read the input for comparison.
        return new PreparedSource(sourcePath, sourcePath, null, detectedKind, sourceHash);
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
        CancellationToken cancellationToken = default)
    {
        Directory.CreateDirectory(outputDirectory);
        var info = await WithFileBudgetAsync(
            token => _probe.ReadAsync(source.ProcessingPath, token),
            $"Image probe for '{source.SourcePath}'",
            cancellationToken);
        if (info.Width is null or <= 0 || info.Height is null or <= 0)
        {
            throw new InvalidDataException($"Could not read image dimensions from '{source.SourcePath}'.");
        }

        var (_, probedExtension) = MediaProbe.DetectIdentity(info);
        var sourceExtension = probedExtension;
        var planned = MediaTransformPlanner.PlanImage(
            seed,
            source.ProcessingPath,
            "__planning__",
            sourceExtension,
            info.Width.Value,
            info.Height.Value,
            info.RotationDegrees,
            preset.EnhancedVariation);

        var outputPath = OutputNameGenerator.NewFilePath(outputDirectory, planned.OutputExtension);
        AssertOutputExtension(outputPath, planned.OutputExtension);
        var args = ReplaceExactPath(planned.Args, "__planning__", outputPath);
        // Heatup -fs fence: output must fit its kind limit.
        var fenced = InsertFsFence(args, planned.OutputByteLimit, outputPath);

        try
        {
            // Per-invocation budgets like heatup's per-command timeouts.
            await RunBoundedAsync(
                "Image transform", fenced, TimeSpan.FromMinutes(5),
                source.ProcessingPath, cancellationToken);
            var outputStats = new FileInfo(outputPath);
            if (!outputStats.Exists || outputStats.Length <= 0 || outputStats.Length > planned.OutputByteLimit)
            {
                throw new InvalidOperationException("Media transform output is empty, unreadable, or exceeds its size limit.");
            }

            var outputProbe = await WithFileBudgetAsync(
                token => _probe.ReadAsync(outputPath, token),
                $"Image output probe for '{outputPath}'",
                cancellationToken);
            if (outputProbe.Width is null or <= 0 || outputProbe.Height is null or <= 0)
            {
                throw new InvalidOperationException("Media transform output has no readable video frame.");
            }

            await RunBoundedAsync(
                "Image decode check", ["-v", "error", "-i", outputPath, "-map", "0:v:0", "-f", "null", "-"],
                TimeSpan.FromMinutes(5), outputPath, cancellationToken);

            var outputHash = await WithFileBudgetAsync(
                token => MediaTransformSeed.HashFileAsync(outputPath, token),
                $"Image output hash for '{outputPath}'",
                cancellationToken);
            if (string.Equals(source.SourceHash, outputHash, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("Media transform output is byte-identical to its immutable source.");
            }

            return new CreatedVariant(outputPath, MediaKind.Image, 0, 0, outputProbe, planned.Profile, seed, info.Width, info.Height);
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
        // Overshoot retries: the recipe's 4% headroom (limit*0.96 target) usually
        // absorbs CBR overshoot, but small files plus +faststart muxing overhead can
        // still land over the -fs fence. Retry with the same seed (same trim/eq):
        // first lower byte targets at the same rung, then lower rungs (each with
        // its own [0.88, 0.80] targets, since a lower rung at 0.88 can still beat a
        // same-rung 0.80 on quality per byte). Attempt 1 is heatup-identical when
        // the preset cap is >= 1080 (preset cap, ratio 0.96). Plans that would
        // encode byte-identical argv to an attempted plan are skipped without
        // running ffmpeg. Bounded by maxAttempts.
        static double[] RatiosForCap(int capIndex) => capIndex == 0 ? [0.96, 0.88, 0.80] : [0.88, 0.80];
        const int maxAttempts = 24;
        var attemptCaps = new List<int> { preset.MaxWidth };
        var attemptedSignatures = new HashSet<string>(StringComparer.Ordinal);
        OutputOvershootException? overshoot = null;
        Exception? lastPlanningError = null;
        var executions = 0;
        var ratioIndex = 0;
        var capIndex = 0;

        // A quality-floor failure means this rung can never fit: skip the ratio
        // phase (a lower target only lowers the bitrate further) and descend rungs.
        bool AdvanceSchedule(int attemptedWidth, bool includeRatios = true)
        {
            if (includeRatios && ratioIndex + 1 < RatiosForCap(capIndex).Length)
            {
                ratioIndex++;
                return true;
            }

            foreach (var rung in MediaTransformPlanner.VideoWidthLadderDescending(int.MaxValue))
            {
                if (rung < attemptedWidth && !attemptCaps.Contains(rung))
                {
                    attemptCaps.Add(rung);
                }
            }

            if (attemptCaps.Count == capIndex + 1)
            {
                var next = attemptedWidth / 2 * 2 - 2;
                if (next >= 32 && next < attemptedWidth)
                {
                    attemptCaps.Add(next);
                }
            }

            capIndex++;
            ratioIndex = 0;
            return capIndex < attemptCaps.Count;
        }

        // Hash once up front: retries must never re-read a multi-GB input.
        sourceHash ??= await WithFileBudgetAsync(
            token => MediaTransformSeed.HashFileAsync(inputPath, token),
            $"Video source hash for '{inputPath}'",
            cancellationToken);

        // Cumulative bound for the whole retry schedule: a persistently failing
        // input (e.g. corrupt output landing at the fence every time) must surface
        // its last error instead of re-encoding for days.
        var retryStopwatch = Stopwatch.StartNew();
        var maxRetryDuration = TimeSpan.FromMinutes(120);

        while (executions < maxAttempts && capIndex < attemptCaps.Count)
        {
            var ratios = RatiosForCap(capIndex);
            var ratio = ratioIndex < ratios.Length ? ratios[ratioIndex] : 0.88;
            MediaTransformPlanner.PlannedVideoTransform planned;
            try
            {
                planned = MediaTransformPlanner.PlanVideo(
                    seed,
                    ordinal,
                    inputPath,
                    "__planning__",
                    probe.Width.Value,
                    probe.Height.Value,
                    probe.RotationDegrees,
                    sourceDurationSeconds,
                    sourceByteCount,
                    IsVideoToolbox,
                    preset.EnhancedVariation,
                    attemptCaps[capIndex],
                    ratio);
            }
            catch (VideoQualityFloorException exception)
            {
                lastPlanningError = exception;
                if (ratio >= 0.96)
                {
                    // Missed at the selection ratio: no rung at/above this cap can
                    // fit (the planner already tried them all), so jump straight
                    // below the ladder instead of re-walking rungs that just failed.
                    var next = Math.Min(attemptCaps[capIndex], 160) / 2 * 2 - 2;
                    if (next >= 32 && !attemptCaps.Contains(next))
                    {
                        attemptCaps.Add(next);
                        capIndex = attemptCaps.Count - 1;
                    }
                    else
                    {
                        capIndex++;
                    }
                }
                else if (!AdvanceSchedule(attemptCaps[capIndex], includeRatios: false))
                {
                    // Missed only because this retry lowered the byte target: a
                    // narrower rung may still fit, so descend rungs (skipping
                    // further ratio cuts, which can only lower the bitrate more).
                    break;
                }

                ratioIndex = 0;
                if (capIndex >= attemptCaps.Count)
                {
                    break;
                }

                continue;
            }

            if (!attemptedSignatures.Add(PlanSignature(planned)))
            {
                if (!AdvanceSchedule(PlannedWidth(planned)))
                {
                    break;
                }

                continue;
            }

            executions++;
            try
            {
                return await ExecuteVideoPlanAsync(
                    planned,
                    inputPath,
                    outputPath,
                    seed,
                    executions,
                    sourceHash,
                    cancellationToken);
            }
            catch (OutputOvershootException exception)
            {
                overshoot = exception;
                if (retryStopwatch.Elapsed > maxRetryDuration ||
                    !AdvanceSchedule(exception.AttemptedWidth))
                {
                    break;
                }
            }
        }

        throw (Exception?)overshoot ?? lastPlanningError
            ?? new InvalidOperationException("Media transform output is empty, unreadable, or exceeds its size limit.");
    }

    private async Task<CreatedVariant> ExecuteVideoPlanAsync(
        MediaTransformPlanner.PlannedVideoTransform planned,
        string inputPath,
        string outputPath,
        string seed,
        int attemptNumber,
        string sourceHash,
        CancellationToken cancellationToken)
    {
        AssertOutputExtension(outputPath, planned.OutputExtension);
        // Placeholders are full argv elements, replaced by exact match so an input
        // path that happens to contain the placeholder text is never rewritten.
        // Suffix the passlog per attempt so a retry never reads the previous
        // attempt's two-pass statistics.
        var passlogPrefix = planned.PasslogPrefix is not null ? $"{outputPath}.passlog.{attemptNumber}" : null;
        var args = ReplaceExactPath(planned.Args, "__planning__", outputPath);
        if (planned.PasslogPrefix is not null && passlogPrefix is not null)
        {
            args = ReplaceExactPath(args, "__planning__.passlog", passlogPrefix);
        }

        var fenced = InsertFsFence(args, planned.OutputByteLimit, outputPath);
        List<string>? firstPass = null;
        if (planned.FirstPassArgs is not null && passlogPrefix is not null)
        {
            firstPass = ReplaceExactPath(planned.FirstPassArgs, "__planning__", outputPath);
            firstPass = ReplaceExactPath(firstPass, "__planning__.passlog", passlogPrefix);
        }

        var trimMs = planned.Evidence.TryGetValue("trimMicroseconds", out var trimObj) switch
        {
            { } when trimObj is long trimUsLong => (int)(trimUsLong / 1000),
            { } when trimObj is int trimUsInt => trimUsInt / 1000,
            _ => 0,
        };
        var attemptedWidth = PlannedWidth(planned);

        try
        {
            // Per-invocation budgets like heatup's per-command timeouts (30min per
            // pass, 5min decode): a hung ffmpeg fails the file, never the lane.
            if (firstPass is not null)
            {
                await RunBoundedAsync(
                    "Video transform first pass", firstPass, TimeSpan.FromMinutes(30),
                    inputPath, cancellationToken);
            }

            // Run the fenced pass directly: when -fs truncates, ffmpeg still exits
            // 0 with a short file, so detect the fence hit from stderr (StandardError
            // only: CombinedOutput could false-positive on a file path containing
            // the phrase). The message text is ffmpeg's ("File size limit exceeded");
            // if a future build rewords it, the size-near-fence rule below still
            // catches truncations.
            var secondPass = await RunBoundedProcessAsync(
                fenced, TimeSpan.FromMinutes(30), inputPath, cancellationToken);
            var fenceHit = secondPass.StandardError.Contains(
                "File size limit exceeded", StringComparison.OrdinalIgnoreCase);
            if (!secondPass.Succeeded && !fenceHit)
            {
                var detail = secondPass.CombinedOutput.Trim();
                throw new InvalidOperationException(
                    detail.Length > 0
                        ? $"'{Path.GetFileName(tools.FFmpeg)}' exited with {secondPass.ExitCode}: {detail}"
                        : $"'{Path.GetFileName(tools.FFmpeg)}' exited with {secondPass.ExitCode}.");
            }

            var outputStats = new FileInfo(outputPath);
            if (!outputStats.Exists || outputStats.Length <= 0)
            {
                throw new InvalidOperationException("Media transform output is empty or unreadable.");
            }

            if (outputStats.Length > planned.OutputByteLimit)
            {
                throw new OutputOvershootException(
                    planned.OutputByteLimit, outputStats.Length, attemptedWidth, fenceHit, null);
            }

            // Near-fence outputs (within 2% under the limit) with a bad duration
            // window are far more likely truncated than genuinely broken: legit
            // CBR encodes target at most 96% of the limit, so a file sitting at
            // the fence with a short duration is a truncation. Decode failures
            // only retry on an explicit fence hit – a corrupt encode must fail
            // fast with its real error, not burn retries relabeled "overshoot".
            var nearFence = outputStats.Length >= (long)(planned.OutputByteLimit * 0.98);
            MediaInfo outputProbe;
            try
            {
                outputProbe = await VerifyVideoProbeAndDurationAsync(outputPath, planned, cancellationToken);
            }
            catch (InvalidOperationException exception) when (fenceHit || nearFence)
            {
                throw new OutputOvershootException(
                    planned.OutputByteLimit, outputStats.Length, attemptedWidth, fenceHit, exception);
            }

            // A corrupt encode must fail fast with its real error: only an explicit
            // fence hit justifies relabeling a decode failure as truncation.
            try
            {
                await RunBoundedAsync(
                    "Video decode check",
                    ["-v", "error", "-i", outputPath, "-map", "0:v:0", "-f", "null", "-"],
                    TimeSpan.FromMinutes(5), outputPath, cancellationToken);
            }
            catch (InvalidOperationException exception) when (fenceHit)
            {
                throw new OutputOvershootException(
                    planned.OutputByteLimit, outputStats.Length, attemptedWidth, fenceHit, exception);
            }

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
        finally
        {
            if (passlogPrefix is not null)
            {
                RemovePasslogs(passlogPrefix);
            }
        }
    }

    /// <summary>
    /// Runs one ffmpeg invocation with its own budget. A timeout is a per-file
    /// data failure, never a worker stop: it converts only when the outer token
    /// is not cancelled, so genuine shutdown still flows as cancellation.
    /// </summary>
    private async Task RunBoundedAsync(
        string description,
        IEnumerable<string> arguments,
        TimeSpan budget,
        string pathForMessage,
        CancellationToken outerToken)
    {
        using var budgetTimeout = CancellationTokenSource.CreateLinkedTokenSource(outerToken);
        budgetTimeout.CancelAfter(budget);
        try
        {
            await RunRequiredAsync(tools.FFmpeg, arguments, budgetTimeout.Token);
        }
        catch (OperationCanceledException exception) when (!outerToken.IsCancellationRequested)
        {
            throw new InvalidOperationException(
                $"{description} timed out after {budget.TotalMinutes:g} minute(s): '{pathForMessage}'.", exception);
        }
    }

    private async Task<ProcessResult> RunBoundedProcessAsync(
        IEnumerable<string> arguments,
        TimeSpan budget,
        string pathForMessage,
        CancellationToken outerToken)
    {
        using var budgetTimeout = CancellationTokenSource.CreateLinkedTokenSource(outerToken);
        budgetTimeout.CancelAfter(budget);
        try
        {
            return await ProcessRunner.RunAsync(tools.FFmpeg, arguments, budgetTimeout.Token);
        }
        catch (OperationCanceledException exception) when (!outerToken.IsCancellationRequested)
        {
            throw new InvalidOperationException(
                $"Video transform pass timed out after {budget.TotalMinutes:g} minute(s): '{pathForMessage}'.", exception);
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

    private static string PlanSignature(MediaTransformPlanner.PlannedVideoTransform planned)
    {
        static string Value(Dictionary<string, object?> evidence, string key) =>
            evidence.TryGetValue(key, out var value) ? value?.ToString() ?? "" : "";
        return string.Join("|",
            Value(planned.Evidence, "outputWidth"),
            Value(planned.Evidence, "outputHeight"),
            Value(planned.Evidence, "videoBitrateBps"),
            Value(planned.Evidence, "audioBitrateBps"),
            Value(planned.Evidence, "targetOutputBytes"),
            Value(planned.Evidence, "profile"));
    }

    private static int PlannedWidth(MediaTransformPlanner.PlannedVideoTransform planned) =>
        planned.Evidence.TryGetValue("outputWidth", out var widthObj) && widthObj is int width ? width : 0;

    public async Task<IReadOnlyList<(SegmentPlan Plan, string Path)>> ExtractSegmentsAsync(
        PreparedSource source,
        LanePaths lane,
        PresetOptions preset,
        double durationSeconds,
        CancellationToken cancellationToken = default)
    {
        // Kept for backward compat but V2 no longer segments – heatup handles long
        // videos with its bitrate ladder down to 160px instead of stream-copy splits.
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
        // V2 precomputes its bitrate ladder to fit min(10MiB, sourceBytes); no post-hoc
        // size-cap recompression passes. Kept as a no-op for CLI compat.
        return Task.FromResult(false);
    }

    private sealed class OutputOvershootException(
        long limitBytes,
        long actualBytes,
        int attemptedWidth,
        bool fenceHit,
        Exception? verificationError)
        : InvalidOperationException(
            actualBytes > limitBytes
                ? $"Media transform output ({actualBytes} bytes) exceeds its size limit ({limitBytes} bytes). " +
                  $"Fence hit: {fenceHit}." +
                  (verificationError is null ? "" : $" Verification: {verificationError.Message}")
                : $"Media transform output ({actualBytes} bytes) sits at its size limit ({limitBytes} bytes) " +
                  $"with failed verification (truncation suspected). Fence hit: {fenceHit}." +
                  (verificationError is null ? "" : $" Verification: {verificationError.Message}"),
            verificationError)
    {
        public int AttemptedWidth { get; } = attemptedWidth;
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

    private static void RemovePasslogs(string? passlogPrefix)
    {
        // Cleanup must never mask the transform error it follows in finally.
        try
        {
            if (string.IsNullOrWhiteSpace(passlogPrefix))
            {
                return;
            }

            var directory = Path.GetDirectoryName(passlogPrefix);
            var prefix = Path.GetFileName(passlogPrefix);
            if (directory is null || !Directory.Exists(directory))
            {
                return;
            }

            foreach (var entry in Directory.EnumerateFiles(directory, prefix + "*"))
            {
                TryDelete(entry);
            }
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or ArgumentException)
        {
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
