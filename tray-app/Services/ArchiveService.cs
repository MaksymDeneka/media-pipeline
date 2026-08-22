using System.Globalization;
using System.IO;
using System.IO.Compression;

namespace MediaPipelineTray.Services;

public sealed record ArchiveResult
{
    public required string Path { get; init; }
    public required int FileCount { get; init; }
    public required long Bytes { get; init; }
    public required IReadOnlyList<string> Missing { get; init; }
}

/// <summary>
/// Collects a finished job's output into a zip, staged in that workspace's sync folder so it is
/// immediately ready to upload.
/// </summary>
public sealed class ArchiveService
{
    private readonly PipelinePaths _paths;

    public ArchiveService(PipelinePaths paths) => _paths = paths;

    /// <summary>
    /// Zips the given output paths, which are relative to the lane's output folder.
    ///
    /// Entries that have since been archived or deleted are reported rather than failing the
    /// whole operation: a partial archive of what still exists is more useful than nothing, and
    /// the caller can say what was missed.
    /// </summary>
    public ArchiveResult Create(
        string preset,
        string workspace,
        IReadOnlyList<string> relativeOutputs,
        string? nameHint = null)
    {
        if (relativeOutputs.Count == 0)
        {
            throw new InvalidOperationException("That job recorded no output to collect.");
        }

        var outputRoot = _paths.OutputDirectory(preset, workspace);
        var syncFolder = _paths.SyncDirectory(workspace);
        Directory.CreateDirectory(syncFolder);

        var stamp = DateTime.Now.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture);
        var baseName = string.IsNullOrWhiteSpace(nameHint) ? preset : $"{preset}-{Sanitize(nameHint)}";
        var zipPath = Path.Combine(syncFolder, $"{baseName}-{stamp}.zip");

        var missing = new List<string>();
        var added = 0;

        // Written to a temp name and moved into place, so a half-written zip is never visible
        // to the uploader watching that folder.
        var tempPath = zipPath + ".building";

        try
        {
            using (var stream = new FileStream(tempPath, FileMode.Create, FileAccess.Write, FileShare.None))
            using (var zip = new ZipArchive(stream, ZipArchiveMode.Create))
            {
                foreach (var relative in relativeOutputs)
                {
                    var source = Path.Combine(outputRoot, relative);

                    if (!File.Exists(source))
                    {
                        missing.Add(relative);
                        continue;
                    }

                    // Entry names use forward slashes so the archive unpacks correctly anywhere.
                    zip.CreateEntryFromFile(source, relative.Replace('\\', '/'), CompressionLevel.Fastest);
                    added++;
                }
            }

            if (added == 0)
            {
                File.Delete(tempPath);
                throw new InvalidOperationException(
                    "None of that job's output is still on disk. It may have been archived or cleaned up.");
            }

            File.Move(tempPath, zipPath, overwrite: true);
        }
        catch
        {
            if (File.Exists(tempPath))
            {
                File.Delete(tempPath);
            }

            throw;
        }

        return new ArchiveResult
        {
            Path = zipPath,
            FileCount = added,
            Bytes = new FileInfo(zipPath).Length,
            Missing = missing,
        };
    }

    /// <summary>Media is already compressed, so Fastest is the right trade for a zip of it.</summary>
    private static string Sanitize(string value)
    {
        var cleaned = new string(value
            .Select(c => Path.GetInvalidFileNameChars().Contains(c) ? '-' : c)
            .ToArray());

        return Path.GetFileNameWithoutExtension(cleaned).Trim();
    }
}
