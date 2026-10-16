using System.Text.Json;
using MediaPipeline.Core.Configuration;
using MediaPipeline.Core.IO;
using MediaPipeline.Core.Media;

namespace MediaPipeline.Core.Runtime;

public sealed record MediaRequest(string FileName, int Copies);

/// <summary>Stages an independent source copy before atomically publishing a request.</summary>
public sealed class MediaRequestQueue
{
    public const int MaximumCopies = 9999;
    private const string RequestFile = "request.json";

    public string Enqueue(
        PipelineConfiguration configuration,
        PipelinePaths paths,
        string presetName,
        string workspace,
        string source,
        int copies)
    {
        ValidateCopies(copies);
        var preset = configuration.Presets.FirstOrDefault(candidate =>
            candidate.Name.Equals(presetName, StringComparison.OrdinalIgnoreCase))
            ?? throw new ArgumentException($"Unknown preset '{presetName}'.");
        var resolvedWorkspace = configuration.Workspaces.FirstOrDefault(candidate =>
            candidate.Equals(workspace, StringComparison.OrdinalIgnoreCase))
            ?? throw new ArgumentException($"Unknown workspace '{workspace}'.");
        var kind = MediaClassifier.Classify(source);
        if (kind is not (MediaKind.Image or MediaKind.Video))
        {
            throw new ArgumentException("Choose a supported image or video file.");
        }
        var fileName = Path.GetFileName(source);
        if (fileName.Contains('\\'))
        {
            throw new ArgumentException("Media filenames cannot contain backslashes.");
        }

        var lane = paths.Lane(preset.Name, resolvedWorkspace);
        var root = QueueDirectory(lane);
        Directory.CreateDirectory(root);
        var id = Guid.NewGuid().ToString("n");
        var staging = Path.Combine(root, $".staging-{id}");
        var destination = Path.Combine(root, id);
        Directory.CreateDirectory(staging);
        try
        {
            Directory.CreateDirectory(Path.Combine(staging, "source"));
            File.Copy(source, Path.Combine(staging, "source", fileName));
            File.WriteAllText(Path.Combine(staging, RequestFile),
                JsonSerializer.Serialize(new MediaRequest(fileName, copies)));
            Directory.Move(staging, destination);
        }
        finally
        {
            if (Directory.Exists(staging)) { Directory.Delete(staging, recursive: true); }
        }
        return id;
    }

    public static string QueueDirectory(LanePaths lane) => Path.Combine(lane.Root, "requests");

    public static string FailedDirectory(LanePaths lane) => Path.Combine(lane.Root, "requests-failed");

    public static void PreserveFailed(LanePaths lane, string directory)
    {
        Directory.CreateDirectory(FailedDirectory(lane));
        Directory.Move(directory, OutputNameGenerator.UniqueDestination(FailedDirectory(lane), Path.GetFileName(directory)));
    }

    public static int RequeueFailed(LanePaths lane)
    {
        if (!Directory.Exists(FailedDirectory(lane))) { return 0; }
        var moved = 0;
        foreach (var directory in Directory.GetDirectories(FailedDirectory(lane)))
        {
            try
            {
                var (request, source) = Read(directory);
                var failed = Path.Combine(directory, "failed", request.FileName);
                if (!source.Exists && File.Exists(failed)) { File.Move(failed, source.FullName); }
                if (!File.Exists(source.FullName)) { continue; }
                Directory.CreateDirectory(QueueDirectory(lane));
                Directory.Move(directory, Path.Combine(QueueDirectory(lane), Path.GetFileName(directory)));
                moved++;
            }
            catch (Exception exception) when (exception is IOException or InvalidDataException or ArgumentException or JsonException)
            {
                // Malformed requests stay available for inspection; valid failures can retry.
            }
        }
        return moved;
    }

    public static IReadOnlyList<string> Pending(LanePaths lane) =>
        Directory.Exists(QueueDirectory(lane))
            ? Directory.GetDirectories(QueueDirectory(lane))
                .Where(directory => !Path.GetFileName(directory).StartsWith('.'))
                .OrderBy(Directory.GetCreationTimeUtc).ThenBy(directory => directory).ToArray()
            : [];

    public static (MediaRequest Request, FileInfo Source) Read(string directory)
    {
        var request = JsonSerializer.Deserialize<MediaRequest>(
            File.ReadAllText(Path.Combine(directory, RequestFile)))
            ?? throw new InvalidDataException("The media request is empty.");
        ValidateCopies(request.Copies);
        if (string.IsNullOrWhiteSpace(request.FileName) ||
            request.FileName != Path.GetFileName(request.FileName) ||
            request.FileName.IndexOfAny(['/', '\\']) >= 0 ||
            MediaClassifier.Classify(request.FileName) is not (MediaKind.Image or MediaKind.Video))
        {
            throw new InvalidDataException("The media request has an invalid source name.");
        }
        return (request, new FileInfo(Path.Combine(directory, "source", request.FileName)));
    }

    public static PresetOptions WithCopies(PresetOptions preset, int copies)
    {
        ValidateCopies(copies);
        // Retain all media transformations. Request counts are exact, independent of
        // the preset's alternating counts and set multiplication.
        return preset with
        {
            VideoCopies = copies,
            ImageCopies = copies,
            CopiesAlternate = 0,
            Batch = BatchMode.PerFile,
            SetCount = 1,
        };
    }

    private static void ValidateCopies(int copies)
    {
        if (copies is < 1 or > MaximumCopies)
        {
            throw new ArgumentException($"Copies must be a whole number from 1 to {MaximumCopies}.");
        }
    }
}
