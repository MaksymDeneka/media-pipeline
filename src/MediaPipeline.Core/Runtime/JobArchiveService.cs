using System.Globalization;
using System.IO.Compression;
using MediaPipeline.Core.IO;

namespace MediaPipeline.Core.Runtime;

public sealed record JobArchiveResult(
    string Path,
    int FileCount,
    long Bytes,
    IReadOnlyList<string> Missing);

/// <summary>Collects one completed job into a zip staged in its workspace sync folder.</summary>
public sealed class JobArchiveService(PipelinePaths paths)
{
    public JobArchiveResult Create(
        string preset,
        string workspace,
        IReadOnlyList<string> relativeOutputs,
        string? nameHint = null)
    {
        if (relativeOutputs.Count == 0)
        {
            throw new InvalidOperationException("That job recorded no output to collect.");
        }

        var outputRoot = Path.GetFullPath(paths.Lane(preset, workspace).Output);
        var syncFolder = paths.Sync(workspace);
        Directory.CreateDirectory(syncFolder);
        var stamp = DateTime.Now.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture);
        var baseName = string.IsNullOrWhiteSpace(nameHint)
            ? Sanitize(preset)
            : $"{Sanitize(preset)}-{Sanitize(nameHint)}";
        var zipPath = OutputNameGenerator.UniqueDestination(
            syncFolder,
            $"{baseName}-{stamp}.zip");
        var temporaryPath = zipPath + ".building";
        var missing = new List<string>();
        var added = 0;

        try
        {
            using (var stream = new FileStream(
                temporaryPath, FileMode.Create, FileAccess.Write, FileShare.None))
            using (var zip = new ZipArchive(stream, ZipArchiveMode.Create))
            {
                foreach (var relative in relativeOutputs.Distinct(StringComparer.OrdinalIgnoreCase))
                {
                    var source = ResolveOutput(outputRoot, relative);
                    if (!File.Exists(source))
                    {
                        missing.Add(relative);
                        continue;
                    }

                    var entryName = relative.Replace('\\', '/').TrimStart('/');
                    zip.CreateEntryFromFile(source, entryName, CompressionLevel.Fastest);
                    added++;
                }
            }

            if (added == 0)
            {
                throw new InvalidOperationException(
                    "None of that job's output is still on disk.");
            }

            File.Move(temporaryPath, zipPath, overwrite: true);
        }
        catch
        {
            File.Delete(temporaryPath);
            throw;
        }

        return new JobArchiveResult(
            zipPath,
            added,
            new FileInfo(zipPath).Length,
            missing);
    }

    private static string ResolveOutput(string outputRoot, string relative)
    {
        if (string.IsNullOrWhiteSpace(relative) || Path.IsPathRooted(relative))
        {
            throw new InvalidOperationException($"Invalid relative output path '{relative}'.");
        }

        var candidate = Path.GetFullPath(Path.Combine(outputRoot, relative));
        var comparison = OperatingSystem.IsWindows()
            ? StringComparison.OrdinalIgnoreCase
            : StringComparison.Ordinal;
        var prefix = outputRoot.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
        if (!candidate.StartsWith(prefix, comparison))
        {
            throw new InvalidOperationException($"Output path leaves its lane: '{relative}'.");
        }

        return candidate;
    }

    private static string Sanitize(string value)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var cleaned = new string(value
            .Select(character => invalid.Contains(character) ? '-' : character)
            .ToArray());
        return Path.GetFileNameWithoutExtension(cleaned).Trim();
    }
}
