using System.Text.Json;
using System.Text;
using System.Security.Cryptography;
using System.Runtime.InteropServices;
using MediaPipeline.Core.Configuration;
using MediaPipeline.Core.Contracts;
using MediaPipeline.Core.IO;
using MediaPipeline.Core.Media;
using MediaPipeline.Core.Runtime;
using MediaPipeline.Core.Tools;
using MediaPipeline.Core.Uploads;

namespace MediaPipeline.Core.SelfTest;

internal static class Program
{
    private static int _failures;

    private static async Task<int> Main()
    {
        var root = Path.Combine(Path.GetTempPath(), $"media-pipeline-core-{Guid.NewGuid():n}");
        Directory.CreateDirectory(root);

        try
        {
            ConfigurationResolvesTypedValues(root);
            ConfigurationExpandsHomePaths();
            PathsUseNativeSeparators(root);
            PathsRejectEscapingSegments(root);
            ControlFlagsRespectScope(root);
            JobArchiveStaysInsideItsLane(root);
            FileClassificationMatchesTheWatcher();
            FileStabilityRequiresAnUnchangedWindow(root);
            TrimPlanningProtectsShortVideos();
            SegmentPlanningPreservesDuration();
            OutputNamesMatchTheContract(root);
            WorkerLockIsExclusive(root);
            LegacyAndNativeUseTheSameWindowsGuard(root);
            await StatusWritesAtomically(root);
            await EventsRemainJsonLines(root);
            await UploadChunksRoundTrip(root);
            await RemoteAssemblyScriptRoundTrip(root);
            await UploadRetentionProtectsFreshAndActiveParts(root);
            EncoderCandidatesMatchThePlatform();
            ToolchainFindsBundledExecutables(root);
        }
        finally
        {
            try
            {
                Directory.Delete(root, recursive: true);
            }
            catch (IOException)
            {
            }
        }

        Console.WriteLine();
        Console.WriteLine(_failures == 0
            ? "MEDIA PIPELINE CORE SELF-TEST OK"
            : $"MEDIA PIPELINE CORE SELF-TEST FAILED: {_failures} check(s)");

        return _failures == 0 ? 0 : 1;
    }

    private static void ConfigurationExpandsHomePaths()
    {
        Section("Home-relative configuration");
        var document = IniDocument.Parse(
        [
            "RemoteSshKeyFile = ~/.ssh/media-pipeline-test-key",
            "[preset test]",
            "VideoCopies = 1",
        ]);

        var configuration = PipelineConfigurationLoader.Resolve(document);
        var expected = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            ".ssh",
            "media-pipeline-test-key");
        Check("expands a home-relative SSH key", configuration.Upload.RemoteSshKeyFile == expected,
            configuration.Upload.RemoteSshKeyFile);
    }

    private static void ToolchainFindsBundledExecutables(string root)
    {
        Section("Bundled tool discovery");
        var applicationDirectory = Path.Combine(root, "app");
        var executableName = OperatingSystem.IsWindows() ? "sample-tool.exe" : "sample-tool";
        var executable = Path.Combine(
            applicationDirectory,
            "tools",
            RuntimeInformation.RuntimeIdentifier,
            executableName);
        Directory.CreateDirectory(Path.GetDirectoryName(executable)!);
        File.WriteAllText(executable, "test");

        Check("finds a runtime-specific bundled executable",
            Toolchain.FindRequired("sample-tool", applicationDirectory) == executable,
            executable);
    }

    private static void ConfigurationResolvesTypedValues(string root)
    {
        Section("Typed configuration");
        var document = IniDocument.Parse(
        [
            $"PipelineRoot = {root}",
            "Crf = nope",
            "ImageProcessingConcurrency = 99",
            "[preset photos]",
            "VideoCopies = 0",
            "ImageCopies = 12",
            "Grouping = PerSet",
            "SetCount = 3",
            "Manifest = yes",
        ]);

        var configuration = PipelineConfigurationLoader.Resolve(document);
        var preset = configuration.Presets.Single();

        Check("reads the pipeline root", configuration.PipelineRoot == root, configuration.PipelineRoot);
        Check("falls back after an invalid integer", configuration.Video.Crf == 24, configuration.Video.Crf);
        Check("caps image concurrency", configuration.Images.ProcessingConcurrency <= 6,
            configuration.Images.ProcessingConcurrency);
        Check("resolves enum values", preset.Grouping == OutputGrouping.PerSet, preset.Grouping);
        Check("resolves boolean aliases", preset.Manifest, preset.Manifest);
        Check("keeps separate media counts", preset.VideoCopies == 0 && preset.ImageCopies == 12,
            $"{preset.VideoCopies}/{preset.ImageCopies}");
        Check("reports invalid configuration", configuration.Warnings.Count >= 1,
            configuration.Warnings.Count);
        Check("reads a quoted value before its trailing comment",
            IniDocument.CleanValue("\"/Users/me/Media #1\" ; keep this note") ==
            "/Users/me/Media #1",
            IniDocument.CleanValue("\"/Users/me/Media #1\" ; keep this note"));
    }

    private static void PathsUseNativeSeparators(string root)
    {
        Section("Native paths");
        var paths = new PipelinePaths(root);
        var lane = paths.Lane("photos", "LC");

        Check("uses workspace-first layout",
            lane.Input == Path.Combine(root, "LC", "photos", "input"), lane.Input);
        Check("puts sync beside lanes", paths.Sync("LC") == Path.Combine(root, "LC", "sync"),
            paths.Sync("LC"));

        var homePath = new PipelinePaths(Path.Combine("~", "MediaPipelineTest"));
        var expectedHome = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "MediaPipelineTest");
        Check("expands a home-relative root", homePath.Root == expectedHome, homePath.Root);
    }

    private static void PathsRejectEscapingSegments(string root)
    {
        Section("Path containment");
        var paths = new PipelinePaths(root);
        foreach (var segment in new[] { "../outside", "..\\outside", "/outside", "C:\\outside", "bad.name" })
        {
            var rejected = false;
            try
            {
                _ = paths.Lane(segment, "LC");
            }
            catch (ArgumentException)
            {
                rejected = true;
            }

            Check($"rejects lane segment {segment}", rejected, segment);
        }

        var configPath = Path.Combine(root, "relative-root.ini");
        File.WriteAllLines(configPath,
        [
            "PipelineRoot = relative-pipeline",
            "[preset safe]",
            "VideoCopies = 1",
        ]);
        var loaded = PipelineConfigurationLoader.Load(configPath);
        Check("resolves relative roots beside the config file",
            loaded.PipelineRoot == Path.Combine(root, "relative-pipeline"), loaded.PipelineRoot);

        var invalidPresetRejected = false;
        try
        {
            _ = PipelineConfigurationLoader.Resolve(IniDocument.Parse(
            [
                "[preset ../outside]",
                "VideoCopies = 1",
            ]));
        }
        catch (ArgumentException)
        {
            invalidPresetRejected = true;
        }
        Check("rejects an escaping preset while loading configuration",
            invalidPresetRejected, invalidPresetRejected);
    }

    private static void ControlFlagsRespectScope(string root)
    {
        Section("Control flags");
        var paths = new PipelinePaths(root);
        Directory.CreateDirectory(paths.Control);
        var control = new WorkerControl(paths);

        File.WriteAllText(Path.Combine(paths.Control, "pause.photos.LC"), "");
        Check("pauses one lane", control.IsPaused("photos", "LC"), false);
        Check("does not pause its sibling", !control.IsPaused("photos", "MD"), true);

        File.WriteAllText(Path.Combine(paths.Control, "stop"), "");
        Check("sees a stop request", control.StopRequested, false);
        control.ClearStopRequest();
        Check("clears a stop request", !control.StopRequested, true);

        control.SetPaused(true, "photos", "MD");
        Check("sets a lane pause", control.IsPaused("photos", "MD"), false);
        control.SetPaused(false, "photos", "MD");
        Check("clears a lane pause", !control.IsPaused("photos", "MD"), true);

        var lane = paths.Lane("photos", "LC");
        Directory.CreateDirectory(lane.Failed);
        File.WriteAllText(Path.Combine(lane.Failed, "retry.jpg"), "retry");
        Check("requeues failed files", control.RequeueFailed("photos", "LC") == 1 &&
            File.Exists(Path.Combine(lane.Input, "retry.jpg")), lane.Input);
    }

    private static void JobArchiveStaysInsideItsLane(string root)
    {
        Section("Job archives");
        var paths = new PipelinePaths(root);
        var lane = paths.Lane("photos", "LC");
        var nested = Path.Combine(lane.Output, "set-one");
        Directory.CreateDirectory(nested);
        File.WriteAllText(Path.Combine(nested, "photo.jpg"), "image");

        var result = new JobArchiveService(paths).Create(
            "photos", "LC", [Path.Combine("set-one", "photo.jpg")], "sample");
        Check("stages a zip in workspace sync",
            result.Path.StartsWith(paths.Sync("LC"), StringComparison.OrdinalIgnoreCase),
            result.Path);
        Check("collects nested output", result.FileCount == 1 && result.Bytes > 0, result);

        var rejectedTraversal = false;
        try
        {
            new JobArchiveService(paths).Create(
                "photos", "LC", [Path.Combine("..", "original", "secret.jpg")]);
        }
        catch (InvalidOperationException)
        {
            rejectedTraversal = true;
        }

        Check("rejects output path traversal", rejectedTraversal, rejectedTraversal);
    }

    private static void WorkerLockIsExclusive(string root)
    {
        Section("Worker lock");
        var paths = new PipelinePaths(root);
        using var first = WorkerLock.TryAcquire(paths);
        using var second = WorkerLock.TryAcquire(paths);

        Check("the first worker gets the lock", first is not null, first is null);
        Check("a second worker is rejected", second is null, second is not null);
    }

    private static void LegacyAndNativeUseTheSameWindowsGuard(string root)
    {
        Section("Legacy/native ownership guard");
        using var first = LegacyWatcherGuard.TryAcquire(root);
        using var second = LegacyWatcherGuard.TryAcquire(
            Path.Combine(root, "equivalent-segment", ".."));

        Check("the native worker acquires the legacy mutex", first is not null, first is null);
        Check("equivalent root paths resolve to the shared mutex",
            !OperatingSystem.IsWindows() || second is null, second is not null);
    }

    private static void TrimPlanningProtectsShortVideos()
    {
        Section("Trim planning");

        var tooShort = TrimPlanner.GetRange(0.4, 15, 95);
        var shortVideo = TrimPlanner.GetRange(1.2, 15, 95);
        var regular = TrimPlanner.GetRange(10, 15, 95);

        Check("does not trim below 500 ms", !tooShort.CanTrim, tooShort);
        Check("limits a short video to ten percent", shortVideo is { MinMs: 10, MaxMs: 100 },
            shortVideo);
        Check("uses the configured normal range", regular is { MinMs: 15, MaxMs: 95 }, regular);

        var used = new HashSet<int>();
        var picks = Enumerable.Range(0, 20)
            .Select(_ => TrimPlanner.PickMilliseconds(regular, used, 20))
            .ToArray();
        Check("keeps trim values unique when the range permits", picks.Distinct().Count() == 20,
            picks.Distinct().Count());
    }

    private static void SegmentPlanningPreservesDuration()
    {
        Section("Segment planning");

        var regular = SegmentPlanner.Plan(47, targetSeconds: 15, minimumSeconds: 11);
        var impossibleRemainder = SegmentPlanner.Plan(21, targetSeconds: 15, minimumSeconds: 11);

        Check("splits a regular long video", regular.Count == 4, regular.Count);
        Check("preserves total duration", Math.Abs(regular.Sum(s => s.DurationSeconds) - 47) < 0.001,
            regular.Sum(s => s.DurationSeconds));
        Check("keeps every regular segment above the minimum",
            regular.All(s => s.DurationSeconds >= 11), string.Join(", ", regular));
        Check("folds an impossible short remainder back in",
            impossibleRemainder.Count == 1 &&
            Math.Abs(impossibleRemainder.Sum(s => s.DurationSeconds) - 21) < 0.001,
            string.Join(", ", impossibleRemainder));
    }

    private static void OutputNamesMatchTheContract(string root)
    {
        Section("Output naming");
        var output = Path.Combine(root, "names");
        Directory.CreateDirectory(output);

        var names = Enumerable.Range(0, 20)
            .Select(_ => Path.GetFileName(OutputNameGenerator.NewFilePath(output, ".jpg")))
            .ToArray();

        Check("uses iPhone-style uppercase names",
            names.All(name => name.Length == 12 && name.StartsWith("IMG_") && name.EndsWith(".JPG")),
            string.Join(", ", names));

        var directory = OutputNameGenerator.NewDirectory(output);
        Check("creates a human-named output directory", Directory.Exists(directory), directory);
    }

    private static void FileClassificationMatchesTheWatcher()
    {
        Section("Media classification");

        Check("recognizes mixed-case video", MediaClassifier.Classify("clip.MOV") == MediaKind.Video,
            MediaClassifier.Classify("clip.MOV"));
        Check("recognizes HEIF images", MediaClassifier.Classify("photo.heif") == MediaKind.Image,
            MediaClassifier.Classify("photo.heif"));
        Check("ignores browser partials",
            MediaClassifier.Classify("download.crdownload") == MediaKind.Temporary,
            MediaClassifier.Classify("download.crdownload"));
        Check("rejects unknown files", MediaClassifier.Classify("notes.txt") == MediaKind.Unsupported,
            MediaClassifier.Classify("notes.txt"));
    }

    private static void FileStabilityRequiresAnUnchangedWindow(string root)
    {
        Section("Stable-file detection");
        var path = Path.Combine(root, "settling.jpg");
        File.WriteAllBytes(path, [1, 2, 3]);
        var tracker = new FileStabilityTracker();
        var firstSeen = DateTimeOffset.UtcNow;

        Check("does not trust the first observation",
            !tracker.IsReady(path, TimeSpan.FromSeconds(3), firstSeen), false);
        Check("waits for the stable interval",
            !tracker.IsReady(path, TimeSpan.FromSeconds(3), firstSeen.AddSeconds(2)), false);

        File.AppendAllText(path, "more");
        Check("resets after the length changes",
            !tracker.IsReady(path, TimeSpan.FromSeconds(3), firstSeen.AddSeconds(4)), false);
        Check("accepts an unchanged unlocked file",
            tracker.IsReady(path, TimeSpan.FromSeconds(3), firstSeen.AddSeconds(8)), true);

        var stuckPath = Path.Combine(root, "never-ready.tmp");
        File.WriteAllBytes(stuckPath, []);
        Check("starts the configured readiness timeout",
            tracker.Observe(stuckPath, TimeSpan.FromSeconds(3), TimeSpan.FromSeconds(10), firstSeen) ==
            StabilityState.Waiting, false);
        Check("reports an input that never becomes ready",
            tracker.Observe(
                stuckPath,
                TimeSpan.FromSeconds(3),
                TimeSpan.FromSeconds(10),
                firstSeen.AddSeconds(11)) == StabilityState.TimedOut, false);
        Check("retains the rejected state until the same file can recover",
            tracker.Observe(
                stuckPath,
                TimeSpan.FromSeconds(3),
                TimeSpan.FromSeconds(10),
                firstSeen.AddSeconds(12)) == StabilityState.Rejected, false);
        File.WriteAllBytes(stuckPath, [1, 2, 3]);
        Check("keeps a timed-out path terminal when its locked writer changes metadata",
            tracker.Observe(
                stuckPath,
                TimeSpan.FromSeconds(3),
                TimeSpan.FromSeconds(10),
                firstSeen.AddSeconds(13)) == StabilityState.Rejected, false);
        Check("restarts readiness after a rejected file becomes stable and exclusive",
            tracker.Observe(
                stuckPath,
                TimeSpan.FromSeconds(3),
                TimeSpan.FromSeconds(10),
                firstSeen.AddSeconds(17)) == StabilityState.Waiting, false);
        Check("accepts the recovered file after a fresh stability window",
            tracker.Observe(
                stuckPath,
                TimeSpan.FromSeconds(3),
                TimeSpan.FromSeconds(10),
                firstSeen.AddSeconds(21)) == StabilityState.Ready, false);

        File.Delete(stuckPath);
        tracker.ForgetMissingFiles([]);
        File.WriteAllBytes(stuckPath, [4, 5, 6]);
        Check("does not reject a later file that reuses a missing path",
            tracker.Observe(
                stuckPath,
                TimeSpan.FromSeconds(3),
                TimeSpan.FromSeconds(10),
                firstSeen.AddSeconds(14)) == StabilityState.Waiting, false);
    }

    private static async Task StatusWritesAtomically(string root)
    {
        Section("Status contract");
        var paths = new PipelinePaths(root);
        var timestamp = DateTimeOffset.UtcNow;
        var status = new WorkerStatus
        {
            StartedUtc = timestamp,
            UpdatedUtc = timestamp,
            PipelineRoot = root,
            Encoder = "libx264",
            PollSeconds = 2,
            Workspaces = ["LC"],
            Presets =
            [
                new PresetStatus
                {
                    Name = "photos",
                    ImageCopies = 12,
                    Grouping = OutputGrouping.Flat,
                    Batch = BatchMode.PerFile,
                },
            ],
            Lanes =
            [
                new LaneStatus { Preset = "photos", Workspace = "LC", Queued = 2 },
            ],
        };

        await new StatusWriter(paths).WriteAsync(status);
        using var json = JsonDocument.Parse(await File.ReadAllTextAsync(paths.StatusFile));

        Check("writes schema v2", json.RootElement.GetProperty("schema").GetString() ==
            "mediaPipeline.status.v2", json.RootElement.GetProperty("schema").GetString() ?? "null");
        Check("writes camel-case properties", json.RootElement.TryGetProperty("pipelineRoot", out _), false);
        Check("leaves no temporary file", Directory.GetFiles(paths.Status, "*.tmp").Length == 0,
            Directory.GetFiles(paths.Status, "*.tmp").Length);
    }

    private static void EncoderCandidatesMatchThePlatform()
    {
        Section("Encoder selection");
        var candidates = VideoEncoderSelector.Candidates(new VideoOptions());

        Check("always has a CPU fallback", candidates[^1] == "libx264", string.Join(", ", candidates));
        Check("does not offer a foreign hardware API",
            !OperatingSystem.IsWindows() || !candidates.Contains("h264_videotoolbox"),
            string.Join(", ", candidates));
    }

    private static async Task EventsRemainJsonLines(string root)
    {
        Section("Event contract");
        var paths = new PipelinePaths(root);
        var timestamp = DateTimeOffset.UtcNow;
        var writer = new EventWriter(paths);

        await writer.AppendAsync(new PipelineEvent
        {
            Timestamp = timestamp,
            Name = "job.start",
            JobId = "abc123",
            Preset = "photos",
            Workspace = "LC",
            Files = ["one.jpg"],
        });
        await writer.AppendAsync(new PipelineEvent
        {
            Timestamp = timestamp,
            Name = "job.variant",
            JobId = "abc123",
            Index = 1,
            Total = 12,
            Output = "IMG_0001.JPG",
        });

        var eventPath = Path.Combine(paths.Logs, $"events-{timestamp.ToLocalTime():yyyyMMdd}.jsonl");
        var lines = await File.ReadAllLinesAsync(eventPath);
        using var first = JsonDocument.Parse(lines[0]);

        Check("writes one JSON object per event", lines.Length == 2, lines.Length);
        Check("keeps the existing short field names",
            first.RootElement.GetProperty("ev").GetString() == "job.start" &&
            first.RootElement.GetProperty("jobId").GetString() == "abc123",
            lines[0]);
        using (var active = JsonDocument.Parse(await File.ReadAllTextAsync(paths.ActiveJobsFile)))
        {
            Check("publishes a worker-owned active job snapshot",
                active.RootElement.GetArrayLength() == 1 &&
                active.RootElement[0].GetProperty("jobId").GetString() == "abc123",
                active.RootElement.GetRawText());
        }
        await writer.AppendAsync(new PipelineEvent
        {
            Timestamp = timestamp,
            Name = "job.done",
            JobId = "abc123",
        });
        using var completed = JsonDocument.Parse(await File.ReadAllTextAsync(paths.ActiveJobsFile));
        Check("removes terminal jobs from the active snapshot",
            completed.RootElement.GetArrayLength() == 0,
            completed.RootElement.GetRawText());
    }

    private static async Task UploadChunksRoundTrip(string root)
    {
        Section("Resumable upload chunks");
        var sourcePath = Path.Combine(root, "upload-source.bin");
        var sourceBytes = new byte[2_500_123];
        new Random(42).NextBytes(sourceBytes);
        await File.WriteAllBytesAsync(sourcePath, sourceBytes);

        var chunks = FileChunker.Plan(sourcePath, chunkSizeMB: 1);
        var partsDirectory = Path.Combine(root, "upload-parts");
        Directory.CreateDirectory(partsDirectory);
        foreach (var chunk in chunks)
        {
            var partPath = Path.Combine(partsDirectory, chunk.FileName);
            await FileChunker.WritePartAsync(sourcePath, chunk, partPath, chunkSizeMB: 1);
            chunk.Sha256 = await FileChunker.HashAsync(partPath);
        }

        var reused = await FileChunker.WritePartAsync(
            sourcePath,
            chunks[0],
            Path.Combine(partsDirectory, chunks[0].FileName),
            chunkSizeMB: 1);
        sourceBytes[0] ^= 0xff;
        await File.WriteAllBytesAsync(sourcePath, sourceBytes);
        var rewroteChangedSource = await FileChunker.WritePartAsync(
            sourcePath,
            chunks[0],
            Path.Combine(partsDirectory, chunks[0].FileName),
            chunkSizeMB: 1);
        await using var assembled = new MemoryStream();
        foreach (var chunk in chunks)
        {
            var part = await File.ReadAllBytesAsync(Path.Combine(partsDirectory, chunk.FileName));
            await assembled.WriteAsync(part);
        }

        Check("plans the expected fixed-size parts", chunks.Count == 3, chunks.Count);
        Check("names parts in stable lexical order",
            chunks[0].FileName == "part00001",
            chunks[0].FileName);
        Check("reuses a complete existing part", !reused, reused);
        Check("rewrites a same-length part when source content changes",
            rewroteChangedSource, rewroteChangedSource);
        Check("assigns a SHA-256 to every part",
            chunks.All(chunk => chunk.Sha256.Length == 64),
            string.Join(", ", chunks.Select(chunk => chunk.Sha256.Length)));
        Check("reassembles the original bytes", sourceBytes.SequenceEqual(assembled.ToArray()),
            assembled.Length);
        var partsHash = await FileChunker.HashPartsAsync(
            chunks.Select(chunk => Path.Combine(partsDirectory, chunk.FileName)));
        Check("hashes the exact concatenated transfer",
            partsHash == await FileChunker.HashAsync(sourcePath), partsHash);
        var uploadJob = new UploadJob
        {
            SourcePath = sourcePath,
            WorkspaceOverride = "LC",
            SourceSha256 = partsHash,
            TotalBytes = sourceBytes.LongLength,
            Target = new UploadTarget
            {
                RemoteName = "test",
                RemoteSftpPartsRoot = "/parts",
                RemotePartsRoot = @"D:\parts",
                RemoteDirectory = @"D:\sync",
                SshHost = "test",
                SshPort = 22,
                SshKeyFile = "",
                ChunkSizeMB = 1,
                ParallelChunks = 1,
                DeleteAfterUpload = true,
            },
        };
        uploadJob.Chunks.AddRange(chunks);
        var remoteScript = UploadService.BuildRemoteScript(
            @"D:\parts\LC\transfer.parts", new string('A', 64));
        Check("uses streaming remote assembly",
            remoteScript.Contains("CopyTo($output)", StringComparison.Ordinal),
            "assembly script did not contain a stream copy");
        Check("authenticates manifest data instead of executing a staged script",
            remoteScript.Contains("MEDIA_PIPELINE_VERIFIED", StringComparison.Ordinal) &&
            remoteScript.Contains("ReadAllBytes($manifestPath)", StringComparison.Ordinal) &&
            remoteScript.Contains("ConvertFrom-Json", StringComparison.Ordinal) &&
            !remoteScript.Contains("Get-Content", StringComparison.Ordinal) &&
            !remoteScript.Contains("PSScriptRoot", StringComparison.Ordinal), remoteScript);
        var encodedCommandLength =
            "powershell -NoProfile -NonInteractive -EncodedCommand ".Length +
            Convert.ToBase64String(Encoding.Unicode.GetBytes(remoteScript)).Length;
        Check("keeps the remote command below the Windows command-line limit",
            encodedCommandLength < 32_767, encodedCommandLength);

        uploadJob.RemoteVerified = true;
        await UploadService.ClaimSourceForDeletionAsync(uploadJob);
        var replacementBytes = Encoding.UTF8.GetBytes("new producer output");
        await File.WriteAllBytesAsync(sourcePath, replacementBytes);
        UploadService.DeleteClaimedSource(uploadJob);
        if (OperatingSystem.IsWindows())
        {
            Check("deletes only the verified source identity",
                uploadJob.SourceDeleted &&
                (await File.ReadAllBytesAsync(sourcePath)).SequenceEqual(replacementBytes),
                uploadJob.RetainedSourcePath ?? sourcePath);
        }
        else
        {
            var retainedPath = uploadJob.RetainedSourcePath;
            Check("returns the verified claim identity for platform-native Trash",
                !uploadJob.SourceDeleted &&
                retainedPath is not null &&
                !string.Equals(retainedPath, sourcePath, StringComparison.Ordinal) &&
                File.Exists(retainedPath) &&
                (await File.ReadAllBytesAsync(sourcePath)).SequenceEqual(replacementBytes),
                retainedPath ?? sourcePath);
            if (retainedPath is not null)
            {
                File.Delete(retainedPath);
            }
        }

        var changedPath = Path.Combine(root, "changed-before-claim.bin");
        await File.WriteAllBytesAsync(changedPath, sourceBytes.Select(value => (byte)(value ^ 1)).ToArray());
        var changedJob = new UploadJob
        {
            SourcePath = changedPath,
            WorkspaceOverride = "LC",
            SourceSha256 = partsHash,
            TotalBytes = sourceBytes.LongLength,
            RemoteVerified = true,
            Target = uploadJob.Target,
        };
        try
        {
            await UploadService.ClaimSourceForDeletionAsync(changedJob);
            Check("refuses a changed source at the deletion claim", false, "claim unexpectedly succeeded");
        }
        catch (IOException)
        {
            Check("refuses a changed source at the deletion claim",
                File.Exists(changedJob.RetainedSourcePath ?? changedPath),
                changedJob.RetainedSourcePath ?? changedPath);
        }

        var openWriterPath = Path.Combine(root, "open-writer-before-claim.bin");
        await File.WriteAllBytesAsync(openWriterPath, sourceBytes);
        var openWriterJob = new UploadJob
        {
            SourcePath = openWriterPath,
            WorkspaceOverride = "LC",
            SourceSha256 = partsHash,
            TotalBytes = sourceBytes.LongLength,
            RemoteVerified = true,
            Target = uploadJob.Target,
        };
        await using (var writer = new FileStream(
            openWriterPath,
            FileMode.Open,
            FileAccess.ReadWrite,
            FileShare.ReadWrite | FileShare.Delete))
        {
            try
            {
                await UploadService.ClaimSourceForDeletionAsync(openWriterJob);
                Check("refuses deletion while an existing writer owns the inode",
                    false, "claim unexpectedly acquired an exclusive lease");
            }
            catch (IOException)
            {
                Check("refuses deletion while an existing writer owns the inode",
                    File.Exists(openWriterJob.RetainedSourcePath ?? openWriterPath),
                    openWriterJob.RetainedSourcePath ?? openWriterPath);
            }
        }
    }

    private static async Task UploadRetentionProtectsFreshAndActiveParts(string root)
    {
        Section("Upload retention");
        var paths = new PipelinePaths(Path.Combine(root, "retention"));
        var workspaceParts = Path.Combine(paths.SyncParts, "LC");
        var fresh = Path.Combine(workspaceParts, "fresh.parts");
        var active = Path.Combine(workspaceParts, "active.parts");
        var inactive = Path.Combine(workspaceParts, "inactive.parts");
        Directory.CreateDirectory(fresh);
        Directory.CreateDirectory(active);
        Directory.CreateDirectory(inactive);
        Directory.SetLastWriteTime(workspaceParts, DateTime.Now.AddDays(-30));
        Directory.SetLastWriteTime(active, DateTime.Now.AddDays(-30));
        var lockPath = active + ".upload.lock";
        await File.WriteAllTextAsync(lockPath, "active");
        var inactiveLockPath = inactive + ".upload.lock";
        await File.WriteAllTextAsync(inactiveLockPath, "inactive");
        File.SetLastWriteTime(lockPath, DateTime.Now.AddDays(-30));
        File.SetLastWriteTime(inactiveLockPath, DateTime.Now.AddDays(-30));
        Directory.SetLastWriteTime(active, DateTime.Now.AddDays(-30));
        Directory.SetLastWriteTime(inactive, DateTime.Now.AddDays(-30));

        var document = IniDocument.Parse(
        [
            $"PipelineRoot = {paths.Root}",
            "ArchiveEnabled = false",
            "AssetRetentionDays = 5",
            "[preset safe]",
            "VideoCopies = 1",
        ]);
        var configuration = PipelineConfigurationLoader.Resolve(document);
        await using var held = new FileStream(
            lockPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        var logger = new PipelineLogger(paths);
        await new ArchiveManager(configuration, paths, logger).RunIfDueAsync(DateTimeOffset.Now);

        Check("keeps fresh parts below an old workspace container", Directory.Exists(fresh), fresh);
        Check("keeps an active old transfer", Directory.Exists(active), active);
        Check("atomically claims and deletes an inactive old transfer",
            !Directory.Exists(inactive), inactive);
        Check("keeps the stable sibling lock inode for future owners",
            File.Exists(inactiveLockPath), inactiveLockPath);
    }

    private static async Task RemoteAssemblyScriptRoundTrip(string root)
    {
        if (!OperatingSystem.IsWindows())
        {
            return;
        }

        Section("Remote Windows assembly");
        var testRoot = Path.Combine(root, "remote-assembly");
        var partsDirectory = Path.Combine(testRoot, "parts");
        var destinationDirectory = Path.Combine(testRoot, "destination");
        Directory.CreateDirectory(partsDirectory);
        Directory.CreateDirectory(destinationDirectory);
        var sourceBytes = Encoding.UTF8.GetBytes("verified remote assembly");
        var firstBytes = sourceBytes[..9];
        var secondBytes = sourceBytes[9..];
        var firstPath = Path.Combine(partsDirectory, "part00001");
        var secondPath = Path.Combine(partsDirectory, "part00002");
        await File.WriteAllBytesAsync(firstPath, firstBytes);
        await File.WriteAllBytesAsync(secondPath, secondBytes);
        var fileName = "assembled.bin";
        var destinationPath = Path.Combine(destinationDirectory, fileName);
        var sourceHash = Convert.ToHexString(SHA256.HashData(sourceBytes));
        var destinationLock = Convert.ToHexString(SHA256.HashData(
            Encoding.UTF8.GetBytes(destinationPath.ToLowerInvariant())));
        var manifest = new
        {
            fileName,
            expectedLength = sourceBytes.LongLength,
            sourceSha256 = sourceHash,
            remoteDirectory = destinationDirectory,
            destinationLock,
            parts = new[]
            {
                new
                {
                    name = Path.GetFileName(firstPath),
                    length = firstBytes.LongLength,
                    sha256 = await FileChunker.HashAsync(firstPath),
                },
                new
                {
                    name = Path.GetFileName(secondPath),
                    length = secondBytes.LongLength,
                    sha256 = await FileChunker.HashAsync(secondPath),
                },
            },
        };
        var manifestPath = Path.Combine(partsDirectory, "manifest.json");
        await File.WriteAllTextAsync(
            manifestPath,
            JsonSerializer.Serialize(manifest),
            new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        var script = UploadService.BuildRemoteScript(
            partsDirectory,
            await FileChunker.HashAsync(manifestPath));
        var encoded = Convert.ToBase64String(Encoding.Unicode.GetBytes(script));
        var powershell = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Windows),
            "System32",
            "WindowsPowerShell",
            "v1.0",
            "powershell.exe");
        var result = await ProcessRunner.RunAsync(
            powershell,
            ["-NoProfile", "-NonInteractive", "-EncodedCommand", encoded]);

        Check("executes under Windows PowerShell and verifies the final hash",
            result.Succeeded &&
            File.Exists(destinationPath) &&
            await FileChunker.HashAsync(destinationPath) == sourceHash,
            result.CombinedOutput);
    }

    private static void Section(string title) => Console.WriteLine($"\n== {title}");

    private static void Check(string label, bool passed, object detail)
    {
        Console.WriteLine($"   {(passed ? "PASS" : "FAIL")}  {label}");
        if (passed)
        {
            return;
        }

        _failures++;
        Console.WriteLine($"         {detail}");
    }
}
