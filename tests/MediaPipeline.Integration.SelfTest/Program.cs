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
            var processor = new PresetProcessor(engine, events, logger, new WorkerControl(paths));
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
            await VerifyAsync(paths, tools);
            await VerifyMediaRequestsAsync(root, tools, encoder, image, video);
            await VerifySizeCapKeepsPresetWidthAsync(root, configuration, engine);
            await VerifyLowBitrateVideoQualityAsync(root, tools, configuration);
            await VerifyHdrVideoColorAsync(root, tools, configuration);
            await VerifyAuxiliaryAudioAsync(root, tools, configuration, paths, engine);
            await VerifyTemporaryCleanupAsync(root, configuration, paths, engine);
            await VerifyExifJpegAsync(root, tools, configuration, paths, engine);
            await VerifyHeicImageAsync(root, tools, configuration, paths, engine);
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

    private static async Task VerifyMediaRequestsAsync(
        string root, Toolchain tools, VideoEncoder encoder, string image, string video)
    {
        Section("Saved media request processing");
        var configuration = PipelineConfigurationLoader.Resolve(IniDocument.Parse(
        [
            $"PipelineRoot = {Path.Combine(root, "saved-media-runs")}",
            "ArchiveEnabled = false", "AssetRetentionDays = 0", "X264Preset = ultrafast",
            "[preset bulk]", "ImageCopies = 100", "VideoCopies = 100", "CopiesAlternate = 50",
            "[preset sets]", "ImageCopies = 100", "VideoCopies = 100",
            "Grouping = PerSet", "SetCount = 10", "Manifest = true",
        ]));
        var paths = new PipelinePaths(configuration.PipelineRoot);
        var events = new EventWriter(paths);
        var logger = new PipelineLogger(paths);
        var control = new WorkerControl(paths);
        var engine = new FfmpegEngine(tools, encoder);
        PipelineWorker CreateWorker() => new(configuration, paths, engine,
            new PresetProcessor(engine, events, logger, control), control,
            new StatusWriter(paths), events, logger, new ArchiveManager(configuration, paths, logger));
        var queue = new MediaRequestQueue();
        queue.Enqueue(configuration, paths, "bulk", "LC", image, 3);
        queue.Enqueue(configuration, paths, "bulk", "LC", video, 2);
        queue.Enqueue(configuration, paths, "sets", "LC", image, 2);
        control.SetPaused(true, "bulk", "LC");
        await CreateWorker().RunOnceAsync(assumeStable: true);
        var lane = paths.Lane("bulk", "LC");
        Check("paused requests remain queued", MediaRequestQueue.Pending(lane).Count == 2, lane.Root);
        Check("paused requests produce no outputs", Directory.GetFiles(lane.Output).Length == 0, lane.Output);
        control.SetPaused(false, "bulk", "LC");
        // A fresh worker reads the durable queue, as it would after an app restart.
        await CreateWorker().RunOnceAsync(assumeStable: true);
        var images = Directory.GetFiles(lane.Output, "*.JPG");
        var videos = Directory.GetFiles(lane.Output, "*.MP4");
        Check("uses exact image count instead of preset/alternate count", images.Length == 3, images.Length);
        Check("uses exact video count instead of preset/alternate count", videos.Length == 2, videos.Length);
        Check("image copies have distinct bytes", images.Select(path =>
            Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(path))))
            .Distinct().Count() == 3, images.Length);
        Check("video copies have distinct bytes", videos.Select(path =>
            Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(path))))
            .Distinct().Count() == 2, videos.Length);
        Check("saved originals stay available", File.Exists(image) && File.Exists(video), image);
        Check("completed requests leave the queue", MediaRequestQueue.Pending(lane).Count == 0, lane.Root);
        var setLane = paths.Lane("sets", "LC");
        Check("set presets still produce the exact requested count",
            Directory.GetFiles(setLane.Output, "*.JPG", SearchOption.AllDirectories).Length == 2, setLane.Output);
        Check("set preset keeps its manifest",
            Directory.GetFiles(setLane.Output, "manifest.json", SearchOption.AllDirectories).Length == 1, setLane.Output);
        queue.Enqueue(configuration, paths, "bulk", "LC", image, 1);
        await CreateWorker().RunOnceAsync(assumeStable: true);
        Check("saved media can run again with a different amount",
            Directory.GetFiles(lane.Output, "*.JPG").Length == 4, lane.Output);
        var broken = Path.Combine(root, "broken.jpg");
        File.WriteAllText(broken, "not an image");
        var failedId = queue.Enqueue(configuration, paths, "bulk", "LC", broken, 2);
        await CreateWorker().RunOnceAsync(assumeStable: true);
        var failedRequest = Path.Combine(MediaRequestQueue.FailedDirectory(lane), failedId);
        Check("failed requests retain their source and count",
            File.Exists(Path.Combine(failedRequest, "failed", "broken.jpg")) &&
            MediaRequestQueue.Read(failedRequest).Request.Copies == 2, failedRequest);
        // Repair the failed source and retry through the same command used by Activity.
        File.Copy(image, Path.Combine(failedRequest, "failed", "broken.jpg"), overwrite: true);
        Check("retries saved media requests", control.RequeueFailed("bulk", "LC") == 1, failedId);
        await CreateWorker().RunOnceAsync(assumeStable: true);
        Check("retry keeps the requested amount instead of using the preset amount",
            Directory.GetFiles(lane.Output, "*.JPG").Length == 6, lane.Output);
        var malformed = Path.Combine(MediaRequestQueue.QueueDirectory(lane), "malformed");
        Directory.CreateDirectory(malformed);
        File.WriteAllText(Path.Combine(malformed, "request.json"), "{broken");
        queue.Enqueue(configuration, paths, "bulk", "LC", image, 1);
        await CreateWorker().RunOnceAsync(assumeStable: true);
        Check("malformed requests do not block valid media",
            Directory.GetFiles(lane.Output, "*.JPG").Length == 7 &&
            Directory.Exists(Path.Combine(MediaRequestQueue.FailedDirectory(lane), "malformed")), malformed);
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

    private static async Task VerifyAsync(PipelinePaths paths, Toolchain tools)
    {
        Section("Flat image preset");
        var imageLane = paths.Lane("image-clean", "test");
        Check("drains image input", Directory.GetFiles(imageLane.Input).Length == 0,
            Directory.GetFiles(imageLane.Input).Length);
        var images = Directory.GetFiles(imageLane.Output, "*.JPG");
        Check("creates one image", images.Length == 1, images.Length);
        if (images.Length == 1)
        {
            await VerifyJpegMetadataAsync(tools, images[0]);
        }
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

        Section("Full-length video preset");
        var longLane = paths.Lane("video-long", "test");
        Check("creates one video variant", Directory.GetFiles(longLane.Output, "*.MP4").Length == 1,
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
            new PresetProcessor(engine, events, logger, control),
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
            new PresetProcessor(engine, events, logger, new WorkerControl(paths)),
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

    private static async Task VerifyAuxiliaryAudioAsync(
        string root, Toolchain tools, PipelineConfiguration configuration, PipelinePaths paths, FfmpegEngine engine)
    {
        Section("MOV with unsupported auxiliary audio");
        var mov = Path.Combine(root, "spatial-audio.mov");
        await RunRequiredAsync(tools.FFmpeg,
            ["-y", "-v", "error",
             "-f", "lavfi", "-i", "testsrc2=size=160x120:rate=10",
             "-f", "lavfi", "-i", "sine=frequency=440:sample_rate=48000",
             "-f", "lavfi", "-i", "anullsrc=r=48000:cl=quad",
             "-t", "1", "-map", "0:v", "-map", "1:a", "-map", "2:a",
             "-c:v", "libx264", "-preset", "ultrafast", "-c:a:0", "aac", "-c:a:1", "pcm_s16le",
             "-disposition:a:0", "0", "-disposition:a:1", "default", mov]);
        // Label the auxiliary sample entry as APAC to reproduce decoder selection
        // without storing a private iPhone recording or needing an APAC encoder.
        var bytes = await File.ReadAllBytesAsync(mov);
        var entry = bytes.AsSpan().LastIndexOf("sowt"u8);
        if (entry < 0) entry = bytes.AsSpan().LastIndexOf("lpcm"u8);
        if (entry < 0) throw new InvalidDataException("Synthetic MOV has no audio sample entry.");
        "apac"u8.CopyTo(bytes.AsSpan(entry, 4));
        await File.WriteAllBytesAsync(mov, bytes);
        var automatic = await ProcessRunner.RunAsync(tools.FFmpeg,
            ["-v", "error", "-i", mov, "-frames:v", "1", "-f", "null", "-"]);
        Check("fixture reproduces unsupported automatic audio selection",
            !automatic.Succeeded && automatic.CombinedOutput.Contains("decoder", StringComparison.OrdinalIgnoreCase),
            automatic.CombinedOutput);

        var preset = configuration.Presets.Single(item => item.Name == "video-clean");
        var source = await engine.PrepareAsync(preset, paths.Lane(preset.Name, "test"), mov, MediaKind.Video);
        try
        {
            var variant = await engine.CreateVideoVariantAsync(source.ProcessingPath,
                Path.Combine(root, "spatial-audio-output"), preset, await engine.DurationAsync(mov),
                new FileInfo(mov).Length, new string('b', 64), 1, source.SourceHash);
            Check("MOV with unsupported auxiliary audio produces video", File.Exists(variant.Path), variant.Path);
            var audio = Path.Combine(root, "spatial-audio-output.pcm");
            await RunRequiredAsync(tools.FFmpeg,
                ["-y", "-v", "error", "-i", variant.Path, "-map", "0:a:0",
                 "-c:a", "pcm_s16le", "-f", "s16le", audio]);
            var samples = await File.ReadAllBytesAsync(audio);
            Check("MOV output retains audible primary audio",
                samples.Length > 0 && samples.Any(value => value != 0), samples.Length);
        }
        finally
        {
            FfmpegEngine.RemoveTemporarySource(source);
        }
    }

    private static async Task VerifyExifJpegAsync(
        string root, Toolchain tools, PipelineConfiguration configuration, PipelinePaths paths, FfmpegEngine engine)
    {
        Section("JPEG EXIF orientation");
        var jpeg = Path.Combine(root, "portrait-source.jpg");
        await CreateExifPortraitJpegAsync(tools, jpeg);
        var preset = configuration.Presets.Single(item => item.Name == "image-clean");
        var source = await engine.PrepareAsync(preset, paths.Lane(preset.Name, "test"), jpeg, MediaKind.Image);
        var variant = await engine.CreateImageVariantAsync(source, Path.Combine(root, "portrait-output"),
            preset, new string('a', 64), 1);
        Check("Image Clean keeps JPEG portrait dimensions after autorotation",
            variant.MediaInfo.Width == 1024 && variant.MediaInfo.Height == 1536,
            $"{variant.MediaInfo.Width}x{variant.MediaInfo.Height}");
        await VerifyJpegMetadataAsync(tools, variant.Path);

        // Point ExifTool at a tool that will reject its arguments after FFmpeg encodes the JPEG.
        var failingEngine = new FfmpegEngine(tools with { ExifTool = tools.FFprobe }, engine.Encoder);
        var failedOutput = Path.Combine(root, "failed-jpeg-output");
        var workDirectory = Path.Combine(root, "failed-jpeg-work");
        try
        {
            await failingEngine.CreateImageVariantAsync(source, failedOutput,
                preset, new string('a', 64), 1, workDirectory: workDirectory);
            Check("rejects JPEGs when metadata removal fails", false, failedOutput);
        }
        catch (InvalidOperationException exception)
        {
            Check("rejects JPEGs when metadata removal fails",
                exception.Message.Contains($"'{Path.GetFileName(tools.FFprobe)}' exited"), exception.Message);
        }
        Check("failed metadata removal leaves no output or working files",
            !Directory.EnumerateFileSystemEntries(failedOutput).Any() &&
            !Directory.EnumerateFileSystemEntries(workDirectory).Any(), failedOutput);
    }

    private static async Task VerifyJpegMetadataAsync(Toolchain tools, string jpeg)
    {
        var exifTool = tools.ExifTool ?? throw new InvalidOperationException("ExifTool is required for JPEG metadata tests.");
        var result = await ProcessRunner.RunAsync(exifTool,
            ["-config", "", "-j", "-G1", "-Comment", "-EXIF:All", "-XMP:All", "-IPTC:All", "-ICC_Profile:All", jpeg]);
        if (!result.Succeeded)
        {
            throw new InvalidOperationException(result.CombinedOutput);
        }

        using var metadata = JsonDocument.Parse(result.StandardOutput);
        Check("JPEG contains no encoder comment or embedded photo metadata",
            metadata.RootElement[0].EnumerateObject().All(tag => tag.Name == "SourceFile"), result.StandardOutput);
    }

    private static async Task CreateExifPortraitJpegAsync(Toolchain tools, string jpeg)
    {
        await RunRequiredAsync(tools.FFmpeg,
            ["-y", "-v", "error", "-f", "lavfi", "-i",
             "color=c=red:s=1536x1024,drawbox=x=768:y=0:w=768:h=1024:color=blue:t=fill",
             "-frames:v", "1", jpeg]);
        // A minimal EXIF Orientation=6 APP1 segment reproduces iPhone portrait
        // metadata without a dependency on ExifTool or a private photo fixture.
        byte[] orientation = [
            0xff, 0xe1, 0x00, 0x22, 0x45, 0x78, 0x69, 0x66, 0x00, 0x00,
            0x49, 0x49, 0x2a, 0x00, 0x08, 0x00, 0x00, 0x00, 0x01, 0x00,
            0x12, 0x01, 0x03, 0x00, 0x01, 0x00, 0x00, 0x00, 0x06, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
        ];
        var jpegBytes = await File.ReadAllBytesAsync(jpeg);
        await File.WriteAllBytesAsync(jpeg, [.. jpegBytes[..2], .. orientation, .. jpegBytes[2..]]);
    }

    private static async Task VerifyHeicImageAsync(
        string root, Toolchain tools, PipelineConfiguration configuration, PipelinePaths paths, FfmpegEngine engine)
    {
        if (!OperatingSystem.IsMacOS()) return;
        Section("Complete HEIC image decoding");
        var jpeg = Path.Combine(root, "heic-source.jpg");
        var heic = Path.Combine(root, "heic-source.heic");
        await CreateExifPortraitJpegAsync(tools, jpeg);
        await RunRequiredAsync("/usr/bin/sips", ["-s", "format", "heic", jpeg, "--out", heic]);
        // Routing must use content even when the extension is misleading.
        var renamed = Path.Combine(root, "heic-renamed.jpg");
        File.Move(heic, renamed);
        var preset = configuration.Presets.Single(item => item.Name == "image-clean");
        var source = await engine.PrepareAsync(preset, paths.Lane(preset.Name, "test"), renamed, MediaKind.Image);
        try
        {
            var variant = await engine.CreateImageVariantAsync(source, Path.Combine(root, "heic-output"),
                preset, new string('a', 64), 1);
            Check("HEIC retains the complete portrait dimensions",
                variant.MediaInfo.Width == 1024 && variant.MediaInfo.Height == 1536,
                $"{variant.MediaInfo.Width}x{variant.MediaInfo.Height}");
            Check("normalized HEIC still produces JPEG", Path.GetExtension(variant.Path) == ".JPG", variant.Path);
            // The original right half becomes the bottom half after rotation.
            // Sampling it detects accidentally repeated or enlarged first tiles.
            var sample = Path.Combine(root, "heic-bottom.rgb");
            await RunRequiredAsync(tools.FFmpeg,
                ["-y", "-v", "error", "-i", variant.Path, "-vf", "crop=32:32:500:1400,scale=1:1",
                 "-frames:v", "1", "-pix_fmt", "rgb24", "-f", "rawvideo", sample]);
            var rgb = await File.ReadAllBytesAsync(sample);
            Check("HEIC retains pixels beyond the first tile", rgb.Length == 3 && rgb[2] > rgb[0] + 80,
                string.Join(",", rgb));
        }
        finally
        {
            FfmpegEngine.RemoveTemporarySource(source);
        }
        Check("HEIC temporary image is removed", source.TemporaryPath is not null && !File.Exists(source.TemporaryPath),
            source.TemporaryPath ?? "No temporary image was created");
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
        Section("Explicit preset width cap");
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

        Check("output never exceeds the preset width",
            variant.MediaInfo.Width is > 0 and <= 160, variant.MediaInfo.Width ?? 0);
    }

    private static async Task VerifyLowBitrateVideoQualityAsync(
        string root, Toolchain tools, PipelineConfiguration configuration)
    {
        Section("Low-bitrate portrait video quality");
        var source = Path.Combine(root, "low-bitrate-portrait.mp4");
        await RunRequiredAsync(tools.FFmpeg,
        [
            "-y", "-hide_banner", "-loglevel", "error",
            "-f", "lavfi", "-i", "testsrc2=size=720x1280:rate=30:duration=2",
            "-c:v", "libx264", "-preset", "veryfast", "-b:v", "550k", "-pix_fmt", "yuv420p", source,
        ]);
        var sourceHash = MediaTransformSeed.HashFile(source);
        var preset = configuration.Presets.Single(item => item.Name == "video-clean") with { Crf = 18 };
        var encoders = new List<VideoEncoder> { new("libx264", "CPU quality regression") };
        if (OperatingSystem.IsMacOS())
        {
            var selected = await VideoEncoderSelector.SelectAsync(tools, configuration.Video);
            if (selected.Name == "h264_videotoolbox") { encoders.Add(selected); }
        }
        foreach (var encoder in encoders)
        {
            var output = Path.Combine(root, "quality-regression", encoder.Name);
            var engine = new FfmpegEngine(tools, encoder);
            var variant = await engine.CreateVideoVariantAsync(
                source, output, preset, 2, new FileInfo(source).Length,
                MediaTransformSeed.Derive(sourceHash, 0), 0, sourceHash);
            Check($"{encoder.Name} preserves 720x1280 on a small compressed source",
                variant.MediaInfo.Width == 720 && variant.MediaInfo.Height == 1280, variant.MediaInfo);
            Check($"{encoder.Name} allows more bytes to preserve quality",
                new FileInfo(variant.Path).Length > new FileInfo(source).Length, new FileInfo(variant.Path).Length);
            Check($"{encoder.Name} preserves the full clip apart from microtrim",
                variant.MediaInfo.DurationSeconds is > 1.9 and < 2.1, variant.MediaInfo.DurationSeconds ?? 0);
        }
        Check("quality processing leaves the source unchanged", MediaTransformSeed.HashFile(source) == sourceHash, source);
    }

    private static async Task VerifyHdrVideoColorAsync(
        string root, Toolchain tools, PipelineConfiguration configuration)
    {
        Section("HDR to SDR video color");
        var probe = new MediaProbe(tools);
        var preset = configuration.Presets.Single(item => item.Name == "video-clean");
        var encoders = new List<VideoEncoder> { new("libx264", "HDR CPU regression") };
        if (OperatingSystem.IsMacOS())
        {
            var selected = await VideoEncoderSelector.SelectAsync(tools, configuration.Video);
            if (selected.Name == "h264_videotoolbox") { encoders.Add(selected); }
        }
        foreach (var transfer in new[] { "arib-std-b67", "smpte2084" })
        {
            var source = Path.Combine(root, $"hdr-{transfer}.mp4");
            // Use 10-bit HDR frames so an accidental early conversion to 8-bit or
            // a metadata-only fix cannot satisfy the pixel comparison below.
            await RunRequiredAsync(tools.FFmpeg,
            [
                "-y", "-hide_banner", "-loglevel", "error",
                "-f", "lavfi", "-i", "testsrc2=size=320x240:rate=12:duration=0.5",
                "-vf", $"setparams=color_primaries=bt2020:color_trc={transfer}:colorspace=bt2020nc:range=limited",
                "-c:v", "libx264", "-preset", "ultrafast", "-pix_fmt", "yuv420p10le",
                "-colorspace", "bt2020nc", "-color_primaries", "bt2020",
                "-color_trc", transfer, "-color_range", "tv", source,
            ]);
            var info = await probe.ReadAsync(source);
            Check($"probe reads {transfer} HDR color metadata",
                info.IsHdr && info.ColorPrimaries == "bt2020" && info.PixelFormat == "yuv420p10le", info);
            var sourceHash = MediaTransformSeed.HashFile(source);
            foreach (var encoder in encoders)
            {
                var engine = new FfmpegEngine(tools, encoder);
                var variant = await engine.CreateVideoVariantAsync(
                    source, Path.Combine(root, "hdr-output", transfer, encoder.Name), preset,
                    info.DurationSeconds!.Value, new FileInfo(source).Length,
                    MediaTransformSeed.Derive(sourceHash, 0), 0, sourceHash);
                var output = variant.MediaInfo;
                Check($"{transfer} {encoder.Name} produces limited-range BT.709 SDR",
                    !output.IsHdr && output.ColorTransfer == "bt709" && output.ColorSpace == "bt709" &&
                    output.ColorPrimaries == "bt709" && output.ColorRange == "tv" &&
                    output.PixelFormat == "yuv420p", output);
                // Compare against the same HDR signal reduced to 8 bits without
                // tone mapping. A tag-only fix leaves these pixels almost equal.
                var difference = await ProcessRunner.RunAsync(tools.FFmpeg,
                [
                    "-v", "error", "-i", source, "-i", variant.Path,
                    "-filter_complex",
                    "[0:v]setparams=color_primaries=bt709:color_trc=bt709:colorspace=bt709,format=yuv420p[raw];" +
                    "[raw][1:v]blend=all_mode=difference,signalstats,metadata=mode=print:file=-",
                    "-frames:v", "1", "-f", "null", "-",
                ]);
                var average = difference.StandardOutput.Split('\n')
                    .FirstOrDefault(line => line.StartsWith("lavfi.signalstats.YAVG="))?.Split('=')[1];
                Check($"{transfer} {encoder.Name} converts pixels rather than only relabeling them",
                    difference.Succeeded && double.TryParse(average, System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out var delta) && delta > 5,
                    difference.CombinedOutput);
                var colors = await ProcessRunner.RunAsync(tools.FFmpeg,
                [
                    "-v", "error", "-i", variant.Path,
                    "-vf", "signalstats,metadata=mode=print:file=-",
                    "-frames:v", "1", "-f", "null", "-",
                ]);
                var saturation = colors.StandardOutput.Split('\n')
                    .FirstOrDefault(line => line.StartsWith("lavfi.signalstats.SATAVG="))?.Split('=')[1];
                Check($"{transfer} {encoder.Name} preserves saturated HDR chart colors",
                    colors.Succeeded && double.TryParse(saturation, System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out var chroma) && chroma > 40,
                    colors.CombinedOutput);
            }
        }
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
