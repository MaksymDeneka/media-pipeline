using MediaPipeline.Core.Configuration;
using MediaPipeline.Core.IO;

namespace MediaPipeline.Core.Runtime;

public sealed class ArchiveManager(
    PipelineConfiguration configuration,
    PipelinePaths paths,
    PipelineLogger logger)
{
    private DateTimeOffset? _lastCheck;

    public async Task RunIfDueAsync(
        DateTimeOffset now,
        CancellationToken cancellationToken = default)
    {
        if (!configuration.Archive.Enabled && configuration.Archive.AssetRetentionDays <= 0)
        {
            return;
        }

        if (_lastCheck is not null &&
            now - _lastCheck < TimeSpan.FromMinutes(configuration.Archive.CheckIntervalMinutes))
        {
            return;
        }

        _lastCheck = now;
        if (configuration.Archive.Enabled)
        {
            await ArchiveOutputsAsync(now, cancellationToken);
        }

        if (configuration.Archive.AssetRetentionDays > 0)
        {
            await DeleteExpiredAssetsAsync(now, cancellationToken);
        }
    }

    private async Task ArchiveOutputsAsync(
        DateTimeOffset now,
        CancellationToken cancellationToken)
    {
        var cutoff = now.LocalDateTime.AddHours(-configuration.Archive.AgeHours);
        foreach (var workspace in configuration.Workspaces)
        {
            foreach (var preset in configuration.Presets)
            {
                var lane = paths.Lane(preset.Name, workspace);
                if (!Directory.Exists(lane.Output))
                {
                    continue;
                }

                var archived = preset.Grouping == OutputGrouping.Flat
                    ? ArchiveFiles(lane.Output, lane.Archive, cutoff)
                    : ArchiveDirectories(lane.Output, lane.Archive, cutoff);
                if (archived > 0)
                {
                    await logger.InfoAsync(
                        $"Archived {archived} output entr{(archived == 1 ? "y" : "ies")} " +
                        $"from {preset.Name} [{workspace}].",
                        cancellationToken);
                }
            }
        }
    }

    private async Task DeleteExpiredAssetsAsync(
        DateTimeOffset now,
        CancellationToken cancellationToken)
    {
        var cutoff = now.LocalDateTime.AddDays(-configuration.Archive.AssetRetentionDays);
        var targets = new List<(string Path, string Label)>();
        if (Directory.Exists(paths.SyncParts))
        {
            // The workspace directory can be old while a child transfer is brand new. Retain the
            // stable workspace container and evaluate each transfer directory independently.
            targets.AddRange(new DirectoryInfo(paths.SyncParts)
                .EnumerateDirectories()
                .Select(directory => (directory.FullName, $"{directory.Name} sync parts")));
        }

        foreach (var workspace in configuration.Workspaces)
        {
            targets.Add((paths.Sync(workspace), $"{workspace} sync"));
            foreach (var preset in configuration.Presets)
            {
                if (preset.Name.Equals("image-bulk", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var lane = paths.Lane(preset.Name, workspace);
                targets.Add((lane.Archive, $"{preset.Name} archive"));
                targets.Add((lane.Original, $"{preset.Name} original"));
                targets.Add((lane.Failed, $"{preset.Name} failed"));
                if (preset.Segment)
                {
                    targets.Add((lane.Work, $"{preset.Name} work"));
                }
            }
        }

        foreach (var target in targets)
        {
            var deleted = DeleteExpired(target.Path, cutoff);
            if (deleted > 0)
            {
                await logger.InfoAsync(
                    $"Deleted {deleted} expired entr{(deleted == 1 ? "y" : "ies")} from {target.Label}.",
                    cancellationToken);
            }
        }
    }

    private static int ArchiveFiles(string source, string archive, DateTime cutoff)
    {
        var count = 0;
        foreach (var file in new DirectoryInfo(source).EnumerateFiles())
        {
            if (file.LastWriteTime > cutoff)
            {
                continue;
            }

            try
            {
                Directory.CreateDirectory(archive);
                var destination = OutputNameGenerator.UniqueDestination(archive, file.Name);
                var creation = file.CreationTime;
                File.Move(file.FullName, destination);
                TrySetCreationTime(destination, creation);
                count++;
            }
            catch (Exception exception) when (
                exception is IOException or UnauthorizedAccessException)
            {
            }
        }

        return count;
    }

    private static int ArchiveDirectories(string source, string archive, DateTime cutoff)
    {
        var count = 0;
        foreach (var directory in new DirectoryInfo(source).EnumerateDirectories())
        {
            if (directory.LastWriteTime > cutoff)
            {
                continue;
            }

            try
            {
                Directory.CreateDirectory(archive);
                var destination = OutputNameGenerator.UniqueDestination(archive, directory.Name);
                var creation = directory.CreationTime;
                Directory.Move(directory.FullName, destination);
                TrySetCreationTime(destination, creation);
                count++;
            }
            catch (Exception exception) when (
                exception is IOException or UnauthorizedAccessException)
            {
            }
        }

        return count;
    }

    private static int DeleteExpired(string directory, DateTime cutoff)
    {
        if (!Directory.Exists(directory))
        {
            return 0;
        }

        var count = 0;
        foreach (var entry in new DirectoryInfo(directory).EnumerateFileSystemInfos())
        {
            var entryTime = entry.CreationTime <= entry.LastWriteTime
                ? entry.CreationTime
                : entry.LastWriteTime;
            if (entryTime > cutoff)
            {
                continue;
            }

            try
            {
                if (entry is DirectoryInfo childDirectory)
                {
                    if (childDirectory.Name.EndsWith(".parts", StringComparison.OrdinalIgnoreCase) ||
                        File.Exists(childDirectory.FullName + ".upload.lock") ||
                        File.Exists(Path.Combine(childDirectory.FullName, "upload.lock")))
                    {
                        if (!TryClaimAndDeleteUploadDirectory(childDirectory.FullName))
                        {
                            continue;
                        }
                    }
                    else
                    {
                        childDirectory.Delete(recursive: true);
                    }
                }
                else
                {
                    // A sibling lease is removed only by the owner that atomically claimed its
                    // transfer directory. Unlinking it independently is unsafe on POSIX.
                    if (entry.Name.EndsWith(".parts.upload.lock", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }
                    entry.Delete();
                }

                count++;
            }
            catch (Exception exception) when (
                exception is IOException or UnauthorizedAccessException)
            {
            }
        }

        return count;
    }

    private static bool TryClaimAndDeleteUploadDirectory(string directory)
    {
        var lockPath = directory + ".upload.lock";
        FileStream? lease = null;
        string? claimedDirectory = null;
        try
        {
            // The retention lease denies a writer while permitting the directory rename.
            // Renaming is the atomic ownership boundary: a new uploader can safely recreate
            // the original path, while retention deletes only the directory it claimed.
            lease = new FileStream(
                lockPath,
                FileMode.OpenOrCreate,
                FileAccess.ReadWrite,
                FileShare.None);
            var parent = Path.GetDirectoryName(directory)!;
            claimedDirectory = Path.Combine(
                parent,
                $".{Path.GetFileName(directory)}.retention-{Guid.NewGuid():N}");
            Directory.Move(directory, claimedDirectory);
            Directory.Delete(claimedDirectory, recursive: true);
            return true;
        }
        catch (Exception exception) when (
            exception is IOException or UnauthorizedAccessException)
        {
            return false;
        }
        finally
        {
            lease?.Dispose();
        }
    }

    private static void TrySetCreationTime(string path, DateTime creation)
    {
        try
        {
            File.SetCreationTime(path, creation);
        }
        catch (Exception exception) when (
            exception is IOException or UnauthorizedAccessException or PlatformNotSupportedException)
        {
        }
    }
}
