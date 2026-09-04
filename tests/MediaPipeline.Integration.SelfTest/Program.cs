using System.Text.Json;
using MediaPipeline.Core.Configuration;
using MediaPipeline.Core.IO;
using MediaPipeline.Core.Media;
using MediaPipeline.Core.Runtime;
using MediaPipeline.Core.Tools;

namespace MediaPipeline.Integration.SelfTest;

internal static class Program
{
    private static int _failures;

    private static async Task<int> Main()
    {
        var root = Path.Combine(Path.GetTempPath(), $"media-pipeline-integration-{Guid.NewGuid():n}");
        Directory.CreateDirectory(root);

        try
        {
            var tools = Toolchain.Discover();
            var sources = Path.Combine(root, "corpus");
            Directory.CreateDirectory(sources);
            var image = Path.Combine(sources, "source.jpg");
            var video = Path.Combine(sources, "source.mp4");
            await GenerateCorpusAsync(tools, image, video);

            var configuration = CreateConfiguration(root);
            var paths = new PipelinePaths(configuration.PipelineRoot);
            StageInputs(configuration, paths, image, video);

            var encoder = new VideoEncoder("libx264", "integration test CPU encoder");
            var engine = new FfmpegEngine(tools, encoder);
            var events = new EventWriter(paths);
            var logger = new PipelineLogger(paths);
            var processor = new PresetProcessor(engine, events, logger);
            var worker = new PipelineWorker(
                configuration,
                paths,
                engine,
                processor,
                new WorkerControl(paths),
                new StatusWriter(paths),
                events,
                logger,
                new ArchiveManager(configuration, paths, logger));

            await worker.RunOnceAsync(assumeStable: true);
            await VerifyAsync(paths);
            await VerifySizeCapKeepsPresetWidthAsync(root, configuration, engine);
            await VerifyTemporaryCleanupAsync(root, configuration, paths, engine);
            await VerifyReadinessTimeoutAsync(root, tools, encoder);
            await VerifyInitialStatusAsync(root, tools, encoder);
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
            ? "MEDIA PIPELINE INTEGRATION SELF-TEST OK"
            : $"MEDIA PIPELINE INTEGRATION SELF-TEST FAILED: {_failures} check(s)");
        return _failures == 0 ? 0 : 1;
    }

    private static PipelineConfiguration CreateConfiguration(string root)
    {
        var document = IniDocument.Parse(
        [
            $"PipelineRoot = {root}",
            "PreferNvenc = false",
            "PreferAmf = false",
            "ArchiveEnabled = false",
            "AssetRetentionDays = 0",
            "SizeCapMB = 0",
            "X264Preset = ultrafast",
            "[preset image-clean]",
            "VideoCopies = 0",
            "ImageCopies = 1",
            "[preset video-clean]",
            "VideoCopies = 1",
            "ImageCopies = 0",
            "[preset sets-batch]",
            "VideoCopies = 1",
            "ImageCopies = 1",
            "Grouping = PerSet",
            "SetCount = 2",
            "Batch = PerGroup",
            "Manifest = true",
            "[preset video-long]",
            "VideoCopies = 1",
            "ImageCopies = 0",
            "Segment = true",
            "SegmentTargetSeconds = 3",
            "SegmentMinSeconds = 2",
        ]);

        return PipelineConfigurationLoader.Resolve(document) with { Workspaces = ["test"] };
    }

    private static async Task GenerateCorpusAsync(Toolchain tools, string image, string video)
    {
        await RunRequiredAsync(tools.FFmpeg,
        [
            "-y",
            "-hide_banner",
            "-loglevel",
            "error",
            "-f",
            "lavfi",
            "-i",
            "testsrc=size=320x240:rate=1",
            "-frames:v",
            "1",
            image,
        ]);
        await RunRequiredAsync(tools.FFmpeg,
        [
            "-y",
            "-hide_banner",
            "-loglevel",
            "error",
            "-f",
            "lavfi",
            "-i",
            "testsrc=size=320x240:rate=24:duration=6",
            "-f",
            "lavfi",
            "-i",
            "sine=frequency=440:duration=6",
            "-c:v",
            "libx264",
            "-preset",
            "ultrafast",
            "-pix_fmt",
            "yuv420p",
            "-c:a",
            "aac",
            "-shortest",
            video,
        ]);
    }

    private static void StageInputs(
        PipelineConfiguration configuration,
        PipelinePaths paths,
        string image,
        string video)
    {
        foreach (var directory in paths.RequiredDirectories(configuration))
        {
            Directory.CreateDirectory(directory);
        }

        File.Copy(image, Path.Combine(paths.Lane("image-clean", "test").Input, "photo.jpg"));
        File.Copy(video, Path.Combine(paths.Lane("video-clean", "test").Input, "clip.mp4"));

        var setsInput = paths.Lane("sets-batch", "test").Input;
        File.Copy(image, Path.Combine(setsInput, "set-photo.jpg"));
        File.Copy(video, Path.Combine(setsInput, "set-video.mp4"));

        File.Copy(video, Path.Combine(paths.Lane("video-long", "test").Input, "long.mp4"));
    }

    private static async Task VerifyAsync(PipelinePaths paths)
    {
        Section("Flat image preset");
        var imageLane = paths.Lane("image-clean", "test");
        Check("drains image input", Directory.GetFiles(imageLane.Input).Length == 0,
            Directory.GetFiles(imageLane.Input).Length);
        Check("creates one image", Directory.GetFiles(imageLane.Output, "*.JPG").Length == 1,
            Directory.GetFiles(imageLane.Output).Length);
        Check("moves the image source", Directory.GetFiles(imageLane.Original).Length == 1,
            Directory.GetFiles(imageLane.Original).Length);

        Section("Flat video preset");
        var videoLane = paths.Lane("video-clean", "test");
        Check("creates one video", Directory.GetFiles(videoLane.Output, "*.MP4").Length == 1,
            Directory.GetFiles(videoLane.Output).Length);
        Check("moves the video source", Directory.GetFiles(videoLane.Original).Length == 1,
            Directory.GetFiles(videoLane.Original).Length);

        Section("Per-set manifest preset");
        var setsLane = paths.Lane("sets-batch", "test");
        var batches = Directory.GetDirectories(setsLane.Output);
        Check("creates one batch container", batches.Length == 1, batches.Length);
        if (batches.Length == 1)
        {
            var sets = Directory.GetDirectories(batches[0]);
            var manifestPath = Path.Combine(batches[0], "manifest.json");
            Check("creates two sets", sets.Length == 2, sets.Length);
            Check("puts both media files in every set",
                sets.All(set => Directory.GetFiles(set).Length == 2),
                string.Join(", ", sets.Select(set => Directory.GetFiles(set).Length)));
            Check("writes a manifest", File.Exists(manifestPath), manifestPath);
            if (File.Exists(manifestPath))
            {
                using var manifest = JsonDocument.Parse(await File.ReadAllTextAsync(manifestPath));
                Check("records every generated variant",
                    manifest.RootElement.GetProperty("variants").GetArrayLength() == 4,
                    manifest.RootElement.GetProperty("variants").GetArrayLength());
            }
        }

        Check("moves both batch sources", Directory.GetFiles(setsLane.Original).Length == 2,
            Directory.GetFiles(setsLane.Original).Length);

        Section("Long video preset (V2: no segmentation, bitrate ladder instead)");
        var longLane = paths.Lane("video-long", "test");
        Check("creates one ladder variant", Directory.GetFiles(longLane.Output, "*.MP4").Length == 1,
            Directory.GetFiles(longLane.Output).Length);
        Check("leaves no working files behind", Directory.GetFiles(longLane.Work).Length == 0,
            Directory.GetFiles(longLane.Work).Length);

        Section("Runtime contracts");
        Check("writes worker status", File.Exists(paths.StatusFile), paths.StatusFile);
        var eventFiles = Directory.GetFiles(paths.Logs, "events-*.jsonl");
        Check("writes one event stream", eventFiles.Length == 1, eventFiles.Length);
        if (eventFiles.Length == 1)
        {
            var events = await File.ReadAllLinesAsync(eventFiles[0]);
            Check("finishes all four jobs", events.Count(line => line.Contains("job.done")) == 4,
                events.Count(line => line.Contains("job.done")));
            Check("records no failed job", events.All(line => !line.Contains("job.failed")),
                string.Join(Environment.NewLine, events.Where(line => line.Contains("job.failed"))));
        }
    }

    private static async Task VerifyInitialStatusAsync(
        string root,
        Toolchain tools,
        VideoEncoder encoder)
    {
        Section("Worker startup liveness");
        var configuration = CreateConfiguration(Path.Combine(root, "startup"));
        var paths = new PipelinePaths(configuration.PipelineRoot);
        var events = new EventWriter(paths);
        var logger = new PipelineLogger(paths);
        var engine = new FfmpegEngine(tools, encoder);
        var control = new WorkerControl(paths);
        var worker = new PipelineWorker(
            configuration,
            paths,
            engine,
            new PresetProcessor(engine, events, logger),
            control,
            new StatusWriter(paths),
            events,
            logger,
            new ArchiveManager(configuration, paths, logger));

        var run = worker.RunAsync();
        var deadline = DateTime.UtcNow.AddSeconds(3);
        while (!File.Exists(paths.StatusFile) && DateTime.UtcNow < deadline)
        {
            await Task.Delay(25);
        }
        Check("publishes status before the first worker loop finishes",
            File.Exists(paths.StatusFile) && !run.IsCompleted, paths.StatusFile);
        control.RequestStop();
        await run;
    }

    private static async Task VerifyReadinessTimeoutAsync(
        string root,
        Toolchain tools,
        VideoEncoder encoder)
    {
        Section("Input readiness timeout");
        var configuration = CreateConfiguration(Path.Combine(root, "timeout")) with
        {
            Timing = new TimingOptions
            {
                StableSeconds = 5,
                TimeoutSeconds = 1,
                PollSeconds = 1,
            },
        };
        var paths = new PipelinePaths(configuration.PipelineRoot);
        var events = new EventWriter(paths);
        var logger = new PipelineLogger(paths);
        var engine = new FfmpegEngine(tools, encoder);
        var lane = paths.Lane("image-clean", "test");
        Directory.CreateDirectory(lane.Input);
        await File.WriteAllBytesAsync(Path.Combine(lane.Input, "stuck.jpg"), []);
        var worker = new PipelineWorker(
            configuration,
            paths,
            engine,
            new PresetProcessor(engine, events, logger),
            new WorkerControl(paths),
            new StatusWriter(paths),
            events,
            logger,
            new ArchiveManager(configuration, paths, logger));

        await worker.RunOnceAsync(assumeStable: false);
        await Task.Delay(1100);
        await worker.RunOnceAsync(assumeStable: false);

        var eventFile = Directory.GetFiles(paths.Logs, "events-*.jsonl").Single();
        var eventText = await File.ReadAllTextAsync(eventFile);
        Check("retains a timed-out input without moving a later path owner",
            File.Exists(Path.Combine(lane.Input, "stuck.jpg")) &&
            !File.Exists(Path.Combine(lane.Failed, "stuck.jpg")), lane.Input);
        Check("emits a visible failure for the timed-out input",
            eventText.Contains("job.failed", StringComparison.Ordinal), eventText);

        var lockedPath = Path.Combine(lane.Input, "locked.jpg");
        await File.WriteAllBytesAsync(lockedPath, [1, 2, 3]);
        await using (var locked = new FileStream(
            lockedPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
        {
            await worker.RunOnceAsync(assumeStable: false);
            await Task.Delay(1100);
            await worker.RunOnceAsync(assumeStable: false);
            locked.Seek(0, SeekOrigin.End);
            await locked.WriteAsync(new byte[] { 4, 5, 6 });
            await locked.FlushAsync();
            await Task.Delay(1100);
            await worker.RunOnceAsync(assumeStable: false);
            eventText = await File.ReadAllTextAsync(eventFile);
            Check("reports a locked input once its readiness timeout elapses",
                eventText.Contains("locked.jpg", StringComparison.Ordinal) &&
                File.Exists(lockedPath), eventText);
            Check("emits only one terminal failure while the rejected writer keeps changing",
                eventText.Split("\"ev\":\"job.failed\"", StringSplitOptions.None).Length - 1 == 2,
                eventText);
        }
        await worker.RunOnceAsync(assumeStable: false);
        Check("does not move a rejected path after its lock is released",
            File.Exists(lockedPath) &&
            !File.Exists(Path.Combine(lane.Failed, "locked.jpg")), lockedPath);

        var reusedPath = Path.Combine(lane.Input, "reused.jpg");
        await File.WriteAllBytesAsync(reusedPath, [1, 2, 3]);
        await using (var locked = new FileStream(
            reusedPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
        {
            await worker.RunOnceAsync(assumeStable: false);
            await Task.Delay(1100);
            await worker.RunOnceAsync(assumeStable: false);
        }
        File.Delete(reusedPath);
        await worker.RunOnceAsync(assumeStable: false);
        await File.WriteAllBytesAsync(reusedPath, [4, 5, 6]);
        await worker.RunOnceAsync(assumeStable: false);
        Check("waits normally for a later file that reuses a rejected path",
            File.Exists(reusedPath) && !File.Exists(Path.Combine(lane.Failed, "reused.jpg")),
            reusedPath);
    }

    private static async Task VerifyTemporaryCleanupAsync(
        string root,
        PipelineConfiguration configuration,
        PipelinePaths paths,
        FfmpegEngine engine)
    {
        Section("Failed-operation cleanup");
        var imagePreset = configuration.Presets.Single(preset => preset.Name == "image-clean");
        var videoPreset = configuration.Presets.Single(preset => preset.Name == "video-long");
        var imageLane = paths.Lane(imagePreset.Name, "test");
        var videoLane = paths.Lane(videoPreset.Name, "test");

        var invalidHeic = Path.Combine(root, "invalid.heic");
        await File.WriteAllTextAsync(invalidHeic, "not a heic image");
        var priorHeicTemps = Directory.GetFiles(Path.GetTempPath(), "media-pipeline-heic-*.png")
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        try
        {
            _ = await engine.PrepareAsync(
                imagePreset, imageLane, invalidHeic, MediaKind.Image);
            Check("rejects an invalid HEIC source", false, "PrepareAsync did not throw");
        }
        catch (Exception exception) when (exception is InvalidDataException or InvalidOperationException)
        {
        }
        var leakedHeicTemps = Directory.GetFiles(Path.GetTempPath(), "media-pipeline-heic-*.png")
            .Where(path => !priorHeicTemps.Contains(path))
            .ToArray();
        Check("leaves no HEIC conversion temp behind", leakedHeicTemps.Length == 0,
            string.Join(", ", leakedHeicTemps));

        var invalidMov = Path.Combine(root, "invalid.mov");
        await File.WriteAllTextAsync(invalidMov, "not a mov video");
        var beforeWork = Directory.Exists(videoLane.Work)
            ? Directory.GetFiles(videoLane.Work).ToHashSet(StringComparer.OrdinalIgnoreCase)
            : [];
        try
        {
            _ = await engine.PrepareAsync(
                videoPreset, videoLane, invalidMov, MediaKind.Video);
            Check("rejects an invalid MOV source", false, "PrepareAsync did not throw");
        }
        catch (Exception exception) when (exception is InvalidDataException or InvalidOperationException)
        {
        }
        // V2 has no segmentation: the stub passes the source through without
        // touching the work directory.
        var segments = await engine.ExtractSegmentsAsync(
            new PreparedSource(invalidMov, invalidMov, null, "video", "deadbeef"),
            videoLane,
            videoPreset,
            durationSeconds: 6);
        Check("segment stub passes a single plan through", segments.Count == 1, segments.Count);
        var leakedWork = Directory.GetFiles(videoLane.Work)
            .Where(path => !beforeWork.Contains(path))
            .ToArray();
        Check("leaves no remux or segment temps behind", leakedWork.Length == 0,
            string.Join(", ", leakedWork));
    }

    private static async Task VerifySizeCapKeepsPresetWidthAsync(
        string root,
        PipelineConfiguration configuration,
        FfmpegEngine engine)
    {
        Section("Preset width cap (V2 ladder)");
        var preset = configuration.Presets.Single(item => item.Name == "video-clean") with
        {
            MaxWidth = 160,
        };
        var output = Path.Combine(root, "size-cap-output");
        var source = Path.Combine(root, "corpus", "source.mp4");
        var sourceBytes = new FileInfo(source).Length;
        var sourceHash = MediaTransformSeed.HashFile(source);
        var variant = await engine.CreateVideoVariantAsync(
            source,
            output,
            preset,
            sourceDurationSeconds: 6,
            sourceByteCount: sourceBytes,
            seed: MediaTransformSeed.Derive(sourceHash, 0),
            ordinal: 0,
            sourceHash: sourceHash);

        Check("ladder-capped output never exceeds the preset width",
            variant.MediaInfo.Width is > 0 and <= 160, variant.MediaInfo.Width ?? 0);
    }

    private static async Task RunRequiredAsync(string executable, IReadOnlyList<string> arguments)
    {
        var result = await ProcessRunner.RunAsync(executable, arguments);
        if (!result.Succeeded)
        {
            throw new InvalidOperationException(result.CombinedOutput);
        }
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
