using System.Text.Json;
using MediaPipeline.Core.Configuration;
using MediaPipeline.Core.IO;
using MediaPipeline.Core.Media;
using MediaPipeline.Core.Runtime;
using MediaPipeline.Core.Tools;
using MediaPipeline.Core.Uploads;

namespace MediaPipeline.Worker;

internal static class Program
{
    private static async Task<int> Main(string[] args)
    {
        if (args.Length == 0 || args[0] is "--help" or "-h")
        {
            PrintUsage();
            return args.Length == 0 ? 2 : 0;
        }

        try
        {
            var command = args[0].ToLowerInvariant();
            var configPath = ReadConfigPath(args[1..]);
            return command switch
            {
                "check" => await CheckAsync(configPath),
                "run" => await RunAsync(configPath),
                "once" => await RunOnceAsync(
                    configPath,
                    args.Contains("--assume-stable", StringComparer.OrdinalIgnoreCase)),
                "status" => Status(configPath),
                "pause" => SetPause(configPath, args, paused: true),
                "resume" => SetPause(configPath, args, paused: false),
                "stop" => Stop(configPath),
                "requeue" => Requeue(configPath, args),
                "archive" => await ArchiveAsync(configPath, args),
                "upload" => await UploadAsync(configPath, args),
                "upload-all" => await UploadAllAsync(configPath, args),
                "recompress" => await RecompressAsync(configPath, args),
                _ => UnknownCommand(args[0]),
            };
        }
        catch (Exception exception) when (
            exception is IOException or UnauthorizedAccessException or
            InvalidOperationException or ArgumentException)
        {
            Console.Error.WriteLine($"Worker failed: {exception.Message}");
            return 1;
        }
    }

    private static async Task<int> CheckAsync(string configPath)
    {
        var context = await CreateContextAsync(configPath);
        Console.WriteLine($"Configuration  {Path.GetFullPath(configPath)}");
        Console.WriteLine($"Pipeline root  {context.Paths.Root}");
        Console.WriteLine($"Presets        {context.Configuration.Presets.Count}");
        Console.WriteLine($"Workspaces     {string.Join(", ", context.Configuration.Workspaces)}");

        foreach (var warning in context.Configuration.Warnings)
        {
            Console.WriteLine($"Warning        {warning}");
        }

        Console.WriteLine($"FFmpeg         {context.Tools.FFmpeg}");
        Console.WriteLine($"FFprobe        {context.Tools.FFprobe}");
        Console.WriteLine($"ExifTool       {context.Tools.ExifTool ?? "(optional, not used – metadata stripped with -map_metadata -1)"}");
        Console.WriteLine($"Encoder        {context.Encoder.Name} ({context.Encoder.Description})");
        Console.WriteLine("Check passed.");
        return 0;
    }

    private static int Status(string configPath)
    {
        var (_, paths) = LoadBaseContext(configPath);
        if (!File.Exists(paths.StatusFile))
        {
            Console.WriteLine(JsonSerializer.Serialize(new
            {
                running = WorkerLock.IsHeld(paths),
                status = (object?)null,
            }));
            return 0;
        }

        using var status = JsonDocument.Parse(File.ReadAllText(paths.StatusFile));
        Console.WriteLine(JsonSerializer.Serialize(new
        {
            running = WorkerLock.IsHeld(paths),
            status = status.RootElement,
        }));
        return 0;
    }

    private static int SetPause(string configPath, string[] args, bool paused)
    {
        var (_, paths) = LoadBaseContext(configPath);
        var preset = Option(args, "--preset");
        var workspace = Option(args, "--workspace");
        new WorkerControl(paths).SetPaused(paused, preset, workspace);
        Console.WriteLine(paused ? "Paused." : "Resumed.");
        return 0;
    }

    private static int Stop(string configPath)
    {
        var (_, paths) = LoadBaseContext(configPath);
        new WorkerControl(paths).RequestStop();
        Console.WriteLine("Stop requested. The worker will finish its current operation first.");
        return 0;
    }

    private static int Requeue(string configPath, string[] args)
    {
        var (_, paths) = LoadBaseContext(configPath);
        var preset = RequiredOption(args, "--preset");
        var workspace = RequiredOption(args, "--workspace");
        var moved = new WorkerControl(paths).RequeueFailed(preset, workspace);
        Console.WriteLine($"Requeued {moved} file(s).");
        return 0;
    }

    private static async Task<int> ArchiveAsync(string configPath, string[] args)
    {
        var (configuration, paths) = LoadBaseContext(configPath);
        var preset = RequiredOption(args, "--preset");
        var workspace = RequiredOption(args, "--workspace");
        var outputs = Options(args, "--output");
        var result = new JobArchiveService(paths).Create(
            preset,
            workspace,
            outputs,
            Option(args, "--name"));
        Console.WriteLine(JsonSerializer.Serialize(result, JsonOptions));

        if (!HasFlag(args, "--upload"))
        {
            return 0;
        }

        return await RunUploadAsync(
            paths,
            configuration.Upload,
            result.Path,
            workspace,
            json: HasFlag(args, "--json"));
    }

    private static async Task<int> UploadAsync(string configPath, string[] args)
    {
        var (configuration, paths) = LoadBaseContext(configPath);
        var file = RequiredOption(args, "--file");
        return await RunUploadAsync(
            paths,
            configuration.Upload,
            file,
            Option(args, "--workspace"),
            HasFlag(args, "--json"));
    }

    private static async Task<int> UploadAllAsync(string configPath, string[] args)
    {
        var (configuration, paths) = LoadBaseContext(configPath);
        var requestedWorkspace = Option(args, "--workspace");
        var workspaces = requestedWorkspace is null
            ? configuration.Workspaces
            : configuration.Workspaces.Contains(requestedWorkspace, StringComparer.OrdinalIgnoreCase)
                ? [requestedWorkspace]
                : throw new ArgumentException($"Unknown workspace '{requestedWorkspace}'.");
        var failed = 0;

        foreach (var workspace in workspaces)
        {
            Directory.CreateDirectory(paths.Sync(workspace));
            foreach (var file in Directory.EnumerateFiles(paths.Sync(workspace)).OrderBy(path => path))
            {
                var exitCode = await RunUploadAsync(
                    paths,
                    configuration.Upload,
                    file,
                    workspace,
                    HasFlag(args, "--json"));
                if (exitCode != 0)
                {
                    failed++;
                }
            }
        }

        return failed == 0 ? 0 : 1;
    }

    private static async Task<int> RecompressAsync(string configPath, string[] args)
    {
        var context = await CreateContextAsync(configPath);
        using var workerLock = AcquireWorkerLock(context.Paths);
        using var legacyGuard = AcquireLegacyGuard(context.Paths);
        var onlyPreset = Option(args, "--preset");
        var onlyWorkspace = Option(args, "--workspace");
        var changed = 0;

        foreach (var workspace in context.Configuration.Workspaces)
        {
            if (onlyWorkspace is not null &&
                !workspace.Equals(onlyWorkspace, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            foreach (var preset in context.Configuration.Presets)
            {
                if (preset.SizeCapMB <= 0 ||
                    (onlyPreset is not null &&
                     !preset.Name.Equals(onlyPreset, StringComparison.OrdinalIgnoreCase)))
                {
                    continue;
                }

                var output = context.Paths.Lane(preset.Name, workspace).Output;
                if (!Directory.Exists(output))
                {
                    continue;
                }

                // Outputs are named IMG_####.MP4 (uppercase); match case-insensitively
                // so case-sensitive filesystems still find them.
                foreach (var path in Directory.EnumerateFiles(
                    output, "*", SearchOption.AllDirectories)
                    .Where(candidate => candidate.EndsWith(".mp4", StringComparison.OrdinalIgnoreCase)))
                {
                    if (await context.Engine.RecompressIfOversizedAsync(path, preset))
                    {
                        changed++;
                        Console.WriteLine($"Recompressed {path}");
                    }
                }
            }
        }

        // V2 fences every encode at min(10MiB, source size), so there is nothing
        // post-hoc to enforce; SizeCapMB is legacy. Say so instead of implying a cap.
        Console.WriteLine(changed == 0
            ? "Recompressed 0 oversized output(s). (V2 outputs are fenced at encode time; recompress is a no-op.)"
            : $"Recompressed {changed} oversized output(s).");
        return 0;
    }

    private static async Task<int> RunUploadAsync(
        PipelinePaths paths,
        UploadOptions options,
        string file,
        string? workspace,
        bool json)
    {
        var job = new UploadJob
        {
            SourcePath = Path.GetFullPath(file),
            WorkspaceOverride = workspace,
            Target = UploadTarget.FromConfiguration(options),
        };
        var service = new UploadService(paths);
        var lastPhase = UploadPhase.Queued;
        service.Progress += (_, progress) =>
        {
            if (json)
            {
                Console.WriteLine(JsonSerializer.Serialize(new
                {
                    type = "upload.progress",
                    file = progress.FileName,
                    workspace = progress.Workspace,
                    phase = progress.Phase.ToString(),
                    chunksSent = progress.ChunksSent,
                    chunks = progress.Chunks.Count,
                    bytesSent = progress.BytesSent,
                    bytes = progress.TotalBytes,
                    error = progress.Error,
                    sourceDeleted = progress.SourceDeleted,
                    retainedSourcePath = progress.RetainedSourcePath,
                }, JsonOptions));
            }
            else if (progress.Phase != lastPhase)
            {
                Console.WriteLine($"{progress.FileName}: {progress.Phase}");
                lastPhase = progress.Phase;
            }
        };

        using var cancellation = new CancellationTokenSource();
        ConsoleCancelEventHandler handler = (_, eventArgs) =>
        {
            eventArgs.Cancel = true;
            cancellation.Cancel();
        };
        Console.CancelKeyPress += handler;
        try
        {
            await service.RunAsync(job, cancellation.Token);
        }
        finally
        {
            Console.CancelKeyPress -= handler;
        }

        if (!json && job.Phase == UploadPhase.Done && job.RetainedSourcePath is not null)
        {
            Console.WriteLine($"Verified local source retained at: {job.RetainedSourcePath}");
        }

        return job.Phase switch
        {
            UploadPhase.Done => 0,
            UploadPhase.Cancelled => 130,
            _ => 1,
        };
    }

    private static async Task<int> RunAsync(string configPath)
    {
        var context = await CreateContextAsync(configPath);
        using var workerLock = AcquireWorkerLock(context.Paths);
        using var legacyGuard = AcquireLegacyGuard(context.Paths);

        using var cancellation = new CancellationTokenSource();
        Console.CancelKeyPress += (_, eventArgs) =>
        {
            eventArgs.Cancel = true;
            cancellation.Cancel();
        };

        await context.Worker.RunAsync(cancellation.Token);
        return 0;
    }

    private static async Task<int> RunOnceAsync(string configPath, bool assumeStable)
    {
        var context = await CreateContextAsync(configPath);
        using var workerLock = AcquireWorkerLock(context.Paths);
        using var legacyGuard = AcquireLegacyGuard(context.Paths);
        await context.Worker.RunOnceAsync(assumeStable);
        return 0;
    }

    private static async Task<WorkerContext> CreateContextAsync(string configPath)
    {
        var (configuration, paths) = LoadBaseContext(configPath);
        var tools = Toolchain.Discover();
        var encoder = await VideoEncoderSelector.SelectAsync(tools, configuration.Video);
        var engine = new FfmpegEngine(tools, encoder);
        var events = new EventWriter(paths);
        var logger = new PipelineLogger(paths);
        var processor = new PresetProcessor(engine, events, logger);
        var control = new WorkerControl(paths);
        var status = new StatusWriter(paths);
        var archive = new ArchiveManager(configuration, paths, logger);
        var worker = new PipelineWorker(
            configuration,
            paths,
            engine,
            processor,
            control,
            status,
            events,
            logger,
            archive);

        return new WorkerContext(configuration, paths, tools, encoder, engine, worker);
    }

    private static (PipelineConfiguration Configuration, PipelinePaths Paths) LoadBaseContext(
        string configPath)
    {
        var configuration = PipelineConfigurationLoader.Load(configPath);
        return (configuration, new PipelinePaths(configuration.PipelineRoot));
    }

    private static WorkerLock AcquireWorkerLock(PipelinePaths paths) =>
        WorkerLock.TryAcquire(paths)
        ?? throw new InvalidOperationException($"Another C# worker already owns '{paths.Root}'.");

    private static LegacyWatcherGuard AcquireLegacyGuard(PipelinePaths paths) =>
        LegacyWatcherGuard.TryAcquire(paths.Root)
        ?? throw new InvalidOperationException(
            $"The PowerShell watcher is already running for '{paths.Root}'. Stop it before starting the C# worker.");

    private static int UnknownCommand(string command)
    {
        Console.Error.WriteLine($"Unknown command '{command}'.");
        PrintUsage();
        return 2;
    }

    private static string ReadConfigPath(string[] args)
    {
        for (var index = 0; index < args.Length; index++)
        {
            if (!args[index].Equals("--config", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (index + 1 >= args.Length)
            {
                throw new ArgumentException("--config requires a path.");
            }

            return args[index + 1];
        }

        return Path.Combine(Environment.CurrentDirectory, "config.ini");
    }

    private static string RequiredOption(string[] args, string name) =>
        Option(args, name) ?? throw new ArgumentException($"{name} requires a value.");

    private static string? Option(string[] args, string name)
    {
        for (var index = 1; index < args.Length; index++)
        {
            if (!args[index].Equals(name, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (index + 1 >= args.Length || args[index + 1].StartsWith("--", StringComparison.Ordinal))
            {
                throw new ArgumentException($"{name} requires a value.");
            }

            return args[index + 1];
        }

        return null;
    }

    private static IReadOnlyList<string> Options(string[] args, string name)
    {
        var values = new List<string>();
        for (var index = 1; index < args.Length; index++)
        {
            if (!args[index].Equals(name, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (index + 1 >= args.Length || args[index + 1].StartsWith("--", StringComparison.Ordinal))
            {
                throw new ArgumentException($"{name} requires a value.");
            }

            values.Add(args[index + 1]);
        }

        return values;
    }

    private static bool HasFlag(string[] args, string name) =>
        args.Contains(name, StringComparer.OrdinalIgnoreCase);

    private static void PrintUsage()
    {
        Console.WriteLine("Usage:");
        Console.WriteLine("  media-pipeline-worker check [--config <path>]");
        Console.WriteLine("  media-pipeline-worker run [--config <path>]");
        Console.WriteLine("  media-pipeline-worker once [--config <path>] [--assume-stable]");
        Console.WriteLine("  media-pipeline-worker status [--config <path>]");
        Console.WriteLine("  media-pipeline-worker pause|resume [--preset <name>] [--workspace <name>]");
        Console.WriteLine("  media-pipeline-worker stop [--config <path>]");
        Console.WriteLine("  media-pipeline-worker requeue --preset <name> --workspace <name>");
        Console.WriteLine("  media-pipeline-worker archive --preset <name> --workspace <name> --output <path>...");
        Console.WriteLine("  media-pipeline-worker upload --file <path> [--workspace <name>] [--json]");
        Console.WriteLine("  media-pipeline-worker upload-all [--workspace <name>] [--json]");
        Console.WriteLine("  media-pipeline-worker recompress [--preset <name>] [--workspace <name>]");
    }

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
    };

    private sealed record WorkerContext(
        PipelineConfiguration Configuration,
        PipelinePaths Paths,
        Toolchain Tools,
        VideoEncoder Encoder,
        FfmpegEngine Engine,
        PipelineWorker Worker);
}
