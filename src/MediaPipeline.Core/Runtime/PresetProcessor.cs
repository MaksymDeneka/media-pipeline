using System.Text.Json;
using System.Text.RegularExpressions;
using MediaPipeline.Core.Configuration;
using MediaPipeline.Core.Contracts;
using MediaPipeline.Core.IO;
using MediaPipeline.Core.Media;

namespace MediaPipeline.Core.Runtime;

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
            foreach (var file in files)
            {
                var kind = MediaClassifier.Classify(file.FullName);
                if (!Accepts(preset, kind))
                {
                    continue;
                }

                var familyKey = preset.Manifest
                    ? UniqueFamilyKey(file.Name, usedFamilyKeys)
                    : "";
                var prepared = await engine.PrepareAsync(
                    preset,
                    lane,
                    file.FullName,
                    kind,
                    cancellationToken);

                try
                {
                    var copyCount = CopyCount(preset, kind);
                    var targets = CreateTargets(
                        preset,
                        lane,
                        copyCount,
                        setDirectories,
                        createdDirectories);
                    if (kind == MediaKind.Image)
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
                finally
                {
                    FfmpegEngine.RemoveTemporarySource(prepared);
                }
            }

            if (preset.Manifest && batchDirectory is not null && variants is not null)
            {
                await WriteManifestAsync(preset, batchDirectory, variants, cancellationToken);
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
                $"Preset '{preset.Name}' [{lane.Workspace}] created {outputs.Count} output(s).",
                cancellationToken);
            return new JobResult(jobId, true, outputs, null);
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
            return new JobResult(jobId, false, outputs, exception.Message);
        }
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
        for (var index = 0; index < targets.Count; index++)
        {
            var variant = await engine.CreateImageVariantAsync(
                source, targets[index], preset, cancellationToken);
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
        var duration = await engine.DurationAsync(source.ProcessingPath, cancellationToken);
        var segmentFiles = preset.Segment
            ? await engine.ExtractSegmentsAsync(source, lane, preset, duration, cancellationToken)
            : [(new SegmentPlan(1, -1, duration), source.ProcessingPath)];
        var variantsDone = 0;
        var totalVariants = targets.Count * segmentFiles.Count;

        try
        {
            foreach (var segment in segmentFiles)
            {
                var range = TrimPlanner.GetRange(
                    segment.Plan.DurationSeconds,
                    preset.MinTrimMs,
                    preset.MaxTrimMs);
                var usedTrims = new HashSet<int>();

                for (var index = 0; index < targets.Count; index++)
                {
                    var trim = TrimPlanner.PickMilliseconds(range, usedTrims, targets.Count);
                    var variant = await engine.CreateVideoVariantAsync(
                        segment.Path,
                        targets[index],
                        preset,
                        segment.Plan.DurationSeconds,
                        trim,
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

        }
        finally
        {
            if (preset.Segment)
            {
                foreach (var segment in segmentFiles)
                {
                    TryDelete(segment.Path);
                }
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
        return new ManifestVariant
        {
            FamilyKey = familyKey,
            VariantKey = $"{familyKey}__{setName}",
            Path = $"{setName}/{Path.GetFileName(variant.Path)}",
            RenditionSetKey = setName,
            GenerationBatchKey = batchKey,
            SourceOriginalName = file.Name,
            SourceFamilyName = familyKey,
            SizeBytes = new FileInfo(variant.Path).Length,
            GeneratedAt = DateTimeOffset.UtcNow,
            DurationSeconds = isVideo ? variant.DurationSeconds : 0,
            TransformProfile = isVideo ? "preset_video_micro_trim" : "preset_image_recrop",
            Metadata = new ManifestMetadata
            {
                Encoder = engine.Encoder.Name,
                TrimMs = variant.TrimMs,
                MaxWidth = preset.MaxWidth,
                SourceWidth = isVideo ? null : variant.MediaInfo.Width,
                SourceHeight = isVideo ? null : variant.MediaInfo.Height,
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
        try
        {
            File.Delete(path);
        }
        catch (IOException)
        {
        }
    }

    private static void TryDeleteDirectory(string path)
    {
        try
        {
            Directory.Delete(path, recursive: true);
        }
        catch (IOException)
        {
        }
    }
}
