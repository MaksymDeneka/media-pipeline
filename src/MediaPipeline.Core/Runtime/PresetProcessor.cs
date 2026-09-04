using System.Text.Json;
using System.Text.RegularExpressions;
using MediaPipeline.Core.Configuration;
using MediaPipeline.Core.Contracts;
using MediaPipeline.Core.IO;
using MediaPipeline.Core.Media;

namespace MediaPipeline.Core.Runtime;

/// <summary>
/// Outcome of one preset group. Partial success is possible: per-file failures
/// move only the culprit to failed/ (each with its own job.failed event) while
/// good siblings keep their outputs. In that case Succeeded is false but Outputs
/// still lists the kept files and a terminal job.done is emitted. On outer
/// (batch-wide) failure Outputs is empty: everything was rolled back.
/// </summary>
public sealed record JobResult(
    string JobId,
    bool Succeeded,
    IReadOnlyList<string> Outputs,
    string? Error);

public sealed class PresetProcessor(
    FfmpegEngine engine,
    EventWriter events,
    PipelineLogger logger)
{
    private static readonly JsonSerializerOptions ManifestOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true,
    };

    private int _presetEntryCount;

    public async Task<JobResult> ProcessGroupAsync(
        PresetOptions preset,
        LanePaths lane,
        IReadOnlyList<FileInfo> files,
        CancellationToken cancellationToken = default)
    {
        var jobId = Guid.NewGuid().ToString("n")[..8];
        var outputs = new List<string>();
        var createdDirectories = new List<string>();
        var processed = new List<FileInfo>();
        var variants = preset.Manifest ? new List<ManifestVariant>() : null;

        await events.AppendAsync(new PipelineEvent
        {
            Name = "job.start",
            JobId = jobId,
            Preset = preset.Name,
            Workspace = lane.Workspace,
            Files = files.Select(file => file.Name).ToArray(),
            Bytes = files.Sum(file => file.Length),
        }, cancellationToken);
        await logger.InfoAsync(
            $"Preset '{preset.Name}' [{lane.Workspace}] processing {files.Count} file(s): " +
            string.Join(", ", files.Select(file => file.Name)),
            cancellationToken);

        Directory.CreateDirectory(lane.Output);
        Directory.CreateDirectory(lane.Original);
        Directory.CreateDirectory(lane.Failed);

        string? batchDirectory = null;
        string? batchKey = null;
        IReadOnlyList<string> setDirectories = [];
        IReadOnlyList<string> setNames = ["."];

        try
        {
            if (preset.Grouping == OutputGrouping.PerSet)
            {
                batchDirectory = OutputNameGenerator.NewDirectory(lane.Output);
                batchKey = Path.GetFileName(batchDirectory);
                createdDirectories.Add(batchDirectory);

                var directories = new List<string>(preset.SetCount);
                for (var index = 0; index < preset.SetCount; index++)
                {
                    directories.Add(OutputNameGenerator.NewDirectory(batchDirectory));
                }

                setDirectories = directories;
                setNames = directories.Select(Path.GetFileName).ToArray()!;
            }

            var usedFamilyKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var failed = new List<(FileInfo File, string Error)>();
            foreach (var file in files)
            {
                var familyKey = preset.Manifest
                    ? UniqueFamilyKey(file.Name, usedFamilyKeys)
                    : "";
                PreparedSource? prepared = null;
                var outputMark = outputs.Count;
                var directoryMark = createdDirectories.Count;
                var variantMark = variants?.Count ?? 0;
                try
                {
                    prepared = await engine.PrepareAsync(
                        preset,
                        lane,
                        file.FullName,
                        MediaClassifier.Classify(file.FullName),
                        cancellationToken);

                    // Route by probed content, never by filename: bytes win. A file
                    // whose content kind the preset does not accept fails visibly
                    // into failed/ (with a job.failed event) instead of crashing in
                    // the wrong transform or looping in input/ forever.
                    var detectedKind = prepared.DetectedKind == "video" ? MediaKind.Video : MediaKind.Image;
                    if (!Accepts(preset, detectedKind))
                    {
                        throw new InvalidOperationException(
                            $"Input '{file.Name}' contains {prepared.DetectedKind} data, which preset '{preset.Name}' does not accept.");
                    }

                    var copyCount = CopyCount(preset, detectedKind);
                    var targets = CreateTargets(
                        preset,
                        lane,
                        copyCount,
                        setDirectories,
                        createdDirectories);
                    if (detectedKind == MediaKind.Image)
                    {
                        await ProcessImagesAsync(
                            preset, lane, file, prepared, targets, copyCount, setNames,
                            batchKey, familyKey, variants, outputs, jobId, cancellationToken);
                    }
                    else
                    {
                        await ProcessVideosAsync(
                            preset, lane, file, prepared, targets, copyCount, setNames,
                            batchKey, familyKey, variants, outputs, jobId, cancellationToken);
                    }

                    processed.Add(file);
                }
                catch (Exception exception) when (exception is InvalidDataException or InvalidOperationException)
                {
                    // Infrastructure failures (disk full, read-only FS, planning
                    // bugs) are batch-wide: rethrow past isolation so the outer
                    // handler aborts instead of blaming files one by one.
                    if (IsInfrastructureError(exception.Message, file.Name, file.FullName))
                    {
                        throw;
                    }

                    // Isolate the failure to this file so one bad sibling cannot
                    // discard the batch's good work: drop only its partial outputs,
                    // manifest rows and per-source directories, move only it to
                    // failed/, and continue with the rest.
                    for (var index = outputs.Count - 1; index >= outputMark; index--)
                    {
                        TryDelete(outputs[index]);
                        outputs.RemoveAt(index);
                    }

                    for (var index = createdDirectories.Count - 1; index >= directoryMark; index--)
                    {
                        TryDeleteDirectory(createdDirectories[index]);
                        createdDirectories.RemoveAt(index);
                    }

                    if (variants is not null)
                    {
                        while (variants.Count > variantMark)
                        {
                            variants.RemoveAt(variants.Count - 1);
                        }
                    }

                    // Reporting must never escalate a contained failure into a batch
                    // failure: move/event/log are best effort from here on.
                    try
                    {
                        if (File.Exists(file.FullName))
                        {
                            MoveInput(file.FullName, lane.Failed);
                        }

                        failed.Add((file, exception.Message));
                        // A distinct id per file failure: reusing the batch jobId would
                        // collide with the terminal job.done below and let consumers
                        // drop one side or the other.
                        var fileJobId = Guid.NewGuid().ToString("n")[..8];
                        await events.AppendAsync(new PipelineEvent
                        {
                            Name = "job.failed",
                            JobId = fileJobId,
                            Preset = preset.Name,
                            Workspace = lane.Workspace,
                            File = file.Name,
                            Error = exception.Message,
                        }, cancellationToken);
                        await logger.ErrorAsync(
                            $"Preset '{preset.Name}' [{lane.Workspace}] file '{file.Name}' failed: {exception.Message}",
                            cancellationToken);
                    }
                    catch (Exception reporting) when (reporting is IOException or UnauthorizedAccessException or InvalidOperationException)
                    {
                    }
                }
                finally
                {
                    if (prepared is not null)
                    {
                        FfmpegEngine.RemoveTemporarySource(prepared);
                    }
                }
            }

            // Per-file job.variant progress events for rolled-back copies cannot be
            // retracted; the terminal job.done (with the true output count) and any
            // per-file job.failed entries tell the final truth.
            if (preset.Manifest && batchDirectory is not null && variants is not null && outputs.Count > 0)
            {
                await WriteManifestAsync(preset, batchDirectory, variants, cancellationToken);
            }
            else if (batchDirectory is not null && outputs.Count == 0)
            {
                // An all-failed batch must not leave an empty container with an
                // empty manifest that looks like success-with-zero.
                TryDeleteDirectory(batchDirectory);
                createdDirectories.Remove(batchDirectory);
            }

            foreach (var file in processed)
            {
                MoveInput(file.FullName, lane.Original);
            }

            await events.AppendAsync(new PipelineEvent
            {
                Name = "job.done",
                JobId = jobId,
                Preset = preset.Name,
                Workspace = lane.Workspace,
                Outputs = outputs.Count,
            }, cancellationToken);
            await logger.InfoAsync(
                failed.Count == 0
                    ? $"Preset '{preset.Name}' [{lane.Workspace}] created {outputs.Count} output(s)."
                    : $"Preset '{preset.Name}' [{lane.Workspace}] created {outputs.Count} output(s); {failed.Count} file(s) failed.",
                cancellationToken);
            return failed.Count == 0
                ? new JobResult(jobId, true, outputs, null)
                : new JobResult(
                    jobId,
                    false,
                    outputs,
                    string.Join("; ", failed.Select(entry => $"{entry.File.Name}: {entry.Error}")));
        }
        catch (OperationCanceledException)
        {
            Rollback(preset, outputs, createdDirectories);
            await events.AppendAsync(new PipelineEvent
            {
                Name = "job.cancelled",
                JobId = jobId,
                Preset = preset.Name,
                Workspace = lane.Workspace,
                Error = "Worker stopped before the job completed.",
            }, CancellationToken.None);
            throw;
        }
        catch (Exception exception)
        {
            Rollback(preset, outputs, createdDirectories);
            foreach (var file in files)
            {
                if (File.Exists(file.FullName))
                {
                    MoveInput(file.FullName, lane.Failed);
                }
            }

            await events.AppendAsync(new PipelineEvent
            {
                Name = "job.failed",
                JobId = jobId,
                Preset = preset.Name,
                Workspace = lane.Workspace,
                Error = exception.Message,
            }, cancellationToken);
            await logger.ErrorAsync(
                $"Preset '{preset.Name}' [{lane.Workspace}] failed: {exception.Message}",
                cancellationToken);
            // Rollback deleted the outputs above; report none, not stale paths.
            return new JobResult(jobId, false, [], exception.Message);
        }
    }

    private static bool IsInfrastructureError(string message, string fileName, string fullPath)
    {
        // The exact planning-bug sentence never contains a path.
        if (message.Contains(
            "Media output placeholder not found in planned ffmpeg argv.",
            StringComparison.Ordinal))
        {
            return true;
        }

        // Disk phrases must match the tool output, never the file's own name:
        // remove this file's paths verbatim FIRST (quote-pairing and line-position
        // tricks cannot survive verbatim removal), then strip any remaining quoted
        // spans and line-leading "<path>:" prefixes while KEEPING the text after
        // the colon, so a genuine "/out/IMG.MP4: No space left on device" still
        // matches.
        var scrubbed = message.Replace(fullPath, "", StringComparison.OrdinalIgnoreCase);
        scrubbed = scrubbed.Replace(fileName, "", StringComparison.OrdinalIgnoreCase);
        scrubbed = Regex.Replace(scrubbed, "'[^']*'", "");
        scrubbed = Regex.Replace(scrubbed, "\"[^\"]*\"", "");
        scrubbed = Regex.Replace(scrubbed, @"(?m)^\s*\S.*?:(?=\s)", "");
        return scrubbed.Contains("No space left", StringComparison.OrdinalIgnoreCase) ||
            scrubbed.Contains("Read-only file system", StringComparison.OrdinalIgnoreCase) ||
            scrubbed.Contains("Disk quota", StringComparison.OrdinalIgnoreCase);
    }

    private async Task ProcessImagesAsync(
        PresetOptions preset,
        LanePaths lane,
        FileInfo file,
        PreparedSource source,
        IReadOnlyList<string> targets,
        int copyCount,
        IReadOnlyList<string> setNames,
        string? batchKey,
        string familyKey,
        List<ManifestVariant>? manifest,
        List<string> outputs,
        string jobId,
        CancellationToken cancellationToken)
    {
        // Heatup: deterministic seed per immutable source + ordinal, no random.
        // The source hash was captured once in PrepareAsync; never re-read the input here.
        for (var index = 0; index < targets.Count; index++)
        {
            var seed = MediaTransformSeed.Derive(source.SourceHash, index);
            var variant = await engine.CreateImageVariantAsync(
                source, targets[index], preset, seed, index, cancellationToken);
            outputs.Add(variant.Path);
            await WriteVariantEventAsync(
                jobId, preset, lane, file, index + 1, targets.Count, variant.Path, cancellationToken);

            if (manifest is not null)
            {
                var setName = setNames[index / copyCount];
                manifest.Add(CreateManifestVariant(
                    preset, file, variant, setName, batchKey, familyKey));
            }
        }

    }

    private async Task ProcessVideosAsync(
        PresetOptions preset,
        LanePaths lane,
        FileInfo file,
        PreparedSource source,
        IReadOnlyList<string> targets,
        int copyCount,
        IReadOnlyList<string> setNames,
        string? batchKey,
        string familyKey,
        List<ManifestVariant>? manifest,
        List<string> outputs,
        string jobId,
        CancellationToken cancellationToken)
    {
        // Heatup V2: no segmentation – long videos use the bitrate ladder down to
        // 160px instead of stream-copy splits. Process the full file with V2.
        var duration = await engine.DurationAsync(source.ProcessingPath, cancellationToken);
        long sourceByteCount;
        try
        {
            sourceByteCount = new FileInfo(source.ProcessingPath).Length;
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            throw new InvalidDataException(
                $"Media preparation source is unreadable: '{source.ProcessingPath}'. {exception.Message}", exception);
        }
        var variantsDone = 0;
        var totalVariants = targets.Count;

        for (var index = 0; index < targets.Count; index++)
        {
            var seed = MediaTransformSeed.Derive(source.SourceHash, index);
            var variant = await engine.CreateVideoVariantAsync(
                source.ProcessingPath,
                targets[index],
                preset,
                duration,
                sourceByteCount,
                seed,
                index,
                source.SourceHash,
                cancellationToken);
            outputs.Add(variant.Path);
            variantsDone++;
            await WriteVariantEventAsync(
                jobId, preset, lane, file, variantsDone, totalVariants,
                variant.Path, cancellationToken);

            if (manifest is not null)
            {
                var setName = setNames[index / copyCount];
                manifest.Add(CreateManifestVariant(
                    preset, file, variant, setName, batchKey, familyKey));
            }
        }
    }

    private static bool Accepts(PresetOptions preset, MediaKind kind) => kind switch
    {
        MediaKind.Video => preset.VideoCopies > 0,
        MediaKind.Image => preset.ImageCopies > 0,
        _ => false,
    };

    private int CopyCount(PresetOptions preset, MediaKind kind)
    {
        var count = kind == MediaKind.Video ? preset.VideoCopies : preset.ImageCopies;
        if (preset.CopiesAlternate > 0 && Interlocked.Increment(ref _presetEntryCount) % 2 == 1)
        {
            return preset.CopiesAlternate;
        }

        return count;
    }

    private static IReadOnlyList<string> CreateTargets(
        PresetOptions preset,
        LanePaths lane,
        int copyCount,
        IReadOnlyList<string> setDirectories,
        List<string> createdDirectories)
    {
        switch (preset.Grouping)
        {
            case OutputGrouping.PerSource:
                {
                    var directory = OutputNameGenerator.NewDirectory(lane.Output);
                    createdDirectories.Add(directory);
                    return Enumerable.Repeat(directory, copyCount).ToArray();
                }
            case OutputGrouping.PerSet:
                return setDirectories
                    .SelectMany(directory => Enumerable.Repeat(directory, copyCount))
                    .ToArray();
            default:
                return Enumerable.Repeat(lane.Output, copyCount).ToArray();
        }
    }

    private async Task WriteVariantEventAsync(
        string jobId,
        PresetOptions preset,
        LanePaths lane,
        FileInfo file,
        int index,
        int total,
        string outputPath,
        CancellationToken cancellationToken)
    {
        var relative = Path.GetRelativePath(lane.Output, outputPath);
        await events.AppendAsync(new PipelineEvent
        {
            Name = "job.variant",
            JobId = jobId,
            Preset = preset.Name,
            Workspace = lane.Workspace,
            File = file.Name,
            Index = index,
            Total = total,
            Output = relative,
        }, cancellationToken);
    }

    private ManifestVariant CreateManifestVariant(
        PresetOptions preset,
        FileInfo file,
        CreatedVariant variant,
        string setName,
        string? batchKey,
        string familyKey)
    {
        var isVideo = variant.Kind == MediaKind.Video;
        var seedSuffix = variant.Seed.Length >= 20 ? variant.Seed[..20] : variant.Seed;
        long sizeBytes;
        try
        {
            sizeBytes = new FileInfo(variant.Path).Length;
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            throw new InvalidDataException(
                $"Media transform output is unreadable: '{variant.Path}'. {exception.Message}", exception);
        }

        return new ManifestVariant
        {
            FamilyKey = familyKey,
            VariantKey = $"{familyKey}__{seedSuffix}",
            Path = $"{setName}/{Path.GetFileName(variant.Path)}",
            RenditionSetKey = setName,
            GenerationBatchKey = batchKey,
            SourceOriginalName = file.Name,
            SourceFamilyName = familyKey,
            SizeBytes = sizeBytes,
            GeneratedAt = DateTimeOffset.UtcNow,
            DurationSeconds = isVideo ? variant.DurationSeconds : 0,
            TransformProfile = variant.Profile,
            Metadata = new ManifestMetadata
            {
                Encoder = engine.Encoder.Name,
                TrimMs = variant.TrimMs,
                MaxWidth = preset.MaxWidth,
                SourceWidth = variant.SourceWidth,
                SourceHeight = variant.SourceHeight,
            },
        };
    }

    private static async Task WriteManifestAsync(
        PresetOptions preset,
        string batchDirectory,
        IReadOnlyList<ManifestVariant> variants,
        CancellationToken cancellationToken)
    {
        var manifest = new MediaManifest
        {
            Schema = preset.ManifestSchema,
            GeneratedAt = DateTimeOffset.UtcNow,
            Variants = variants,
        };
        var path = Path.Combine(batchDirectory, "manifest.json");
        await using var stream = File.Create(path);
        await JsonSerializer.SerializeAsync(stream, manifest, ManifestOptions, cancellationToken);
    }

    private static string UniqueFamilyKey(string fileName, ISet<string> used)
    {
        var baseName = Path.GetFileNameWithoutExtension(fileName);
        var family = Regex.Replace(baseName, "[^A-Za-z0-9._-]", "_").Trim('_');
        if (family.Length == 0)
        {
            family = "media";
        }

        var candidate = family;
        var suffix = 2;
        while (!used.Add(candidate))
        {
            candidate = $"{family}_{suffix++}";
        }

        return candidate;
    }

    private static void MoveInput(string path, string destinationDirectory)
    {
        Directory.CreateDirectory(destinationDirectory);
        var destination = OutputNameGenerator.UniqueDestination(destinationDirectory, Path.GetFileName(path));
        File.Move(path, destination);
    }

    private static void Rollback(
        PresetOptions preset,
        IReadOnlyList<string> outputs,
        IReadOnlyList<string> createdDirectories)
    {
        switch (preset.OnFailure)
        {
            case FailureMode.DeleteContainer:
                if (createdDirectories.Count > 0)
                {
                    foreach (var directory in createdDirectories)
                    {
                        TryDeleteDirectory(directory);
                    }
                }
                else
                {
                    foreach (var output in outputs)
                    {
                        TryDelete(output);
                    }
                }

                break;
            case FailureMode.DeleteFiles:
                foreach (var output in outputs)
                {
                    TryDelete(output);
                }

                break;
        }
    }

    private static void TryDelete(string path)
    {
        // Cleanup must never mask the job error it follows.
        try
        {
            File.Delete(path);
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
        }
    }

    private static void TryDeleteDirectory(string path)
    {
        // Cleanup must never mask the job error it follows.
        try
        {
            Directory.Delete(path, recursive: true);
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
        }
    }
}
