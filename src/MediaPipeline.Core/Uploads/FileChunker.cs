using System.Security.Cryptography;

namespace MediaPipeline.Core.Uploads;

/// <summary>Creates resumable, fixed-size, SHA-256-addressed upload parts.</summary>
public static class FileChunker
{
    public static string PartName(int index) => $"part{index:D5}";

    public static IReadOnlyList<ChunkProgress> Plan(string sourcePath, int chunkSizeMB)
    {
        var source = new FileInfo(sourcePath);
        if (!source.Exists)
        {
            throw new FileNotFoundException($"Source file not found: {sourcePath}");
        }

        if (source.Length == 0)
        {
            throw new InvalidOperationException($"Source file is empty: {sourcePath}");
        }

        var chunkSize = checked((long)Math.Max(1, chunkSizeMB) * 1024 * 1024);
        var count = checked((int)Math.Ceiling((double)source.Length / chunkSize));
        var chunks = new List<ChunkProgress>(count);
        for (var index = 0; index < count; index++)
        {
            var offset = (long)index * chunkSize;
            chunks.Add(new ChunkProgress
            {
                Index = index + 1,
                Length = Math.Min(chunkSize, source.Length - offset),
                FileName = PartName(index + 1),
            });
        }

        return chunks;
    }

    public static async Task<bool> WritePartAsync(
        string sourcePath,
        ChunkProgress chunk,
        string partPath,
        int chunkSizeMB,
        CancellationToken cancellationToken = default)
    {
        var existing = new FileInfo(partPath);
        if (existing.Exists && existing.Length == chunk.Length &&
            await PartMatchesSourceAsync(
                sourcePath, chunk, partPath, chunkSizeMB, cancellationToken).ConfigureAwait(false))
        {
            return false;
        }

        Directory.CreateDirectory(Path.GetDirectoryName(partPath)
            ?? throw new InvalidOperationException($"No parent directory for '{partPath}'."));

        var chunkSize = checked((long)Math.Max(1, chunkSizeMB) * 1024 * 1024);
        var offset = checked((long)(chunk.Index - 1) * chunkSize);
        await using var source = new FileStream(
            sourcePath, FileMode.Open, FileAccess.Read, FileShare.Read, 1 << 20, useAsync: true);
        source.Seek(offset, SeekOrigin.Begin);

        var temporaryPath = partPath + ".writing";
        try
        {
            await using (var target = new FileStream(
                temporaryPath, FileMode.Create, FileAccess.Write, FileShare.None, 1 << 20, useAsync: true))
            {
                var buffer = new byte[1 << 20];
                var remaining = chunk.Length;
                while (remaining > 0)
                {
                    var wanted = (int)Math.Min(buffer.Length, remaining);
                    var read = await source.ReadAsync(
                        buffer.AsMemory(0, wanted), cancellationToken).ConfigureAwait(false);
                    if (read == 0)
                    {
                        throw new IOException($"Source ended early while writing {chunk.FileName}.");
                    }

                    await target.WriteAsync(
                        buffer.AsMemory(0, read), cancellationToken).ConfigureAwait(false);
                    remaining -= read;
                }
            }

            File.Move(temporaryPath, partPath, overwrite: true);
            return true;
        }
        finally
        {
            File.Delete(temporaryPath);
        }
    }

    public static async Task<string> HashAsync(
        string path,
        CancellationToken cancellationToken = default)
    {
        await using var stream = new FileStream(
            path, FileMode.Open, FileAccess.Read, FileShare.Read, 1 << 20, useAsync: true);
        return Convert.ToHexString(await SHA256.HashDataAsync(stream, cancellationToken));
    }

    public static async Task<string> HashPartsAsync(
        IEnumerable<string> partPaths,
        CancellationToken cancellationToken = default)
    {
        using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        var buffer = new byte[1 << 20];
        foreach (var partPath in partPaths)
        {
            await using var stream = new FileStream(
                partPath, FileMode.Open, FileAccess.Read, FileShare.Read, buffer.Length, useAsync: true);
            while (true)
            {
                var read = await stream.ReadAsync(buffer, cancellationToken).ConfigureAwait(false);
                if (read == 0)
                {
                    break;
                }
                hash.AppendData(buffer, 0, read);
            }
        }
        return Convert.ToHexString(hash.GetHashAndReset());
    }

    private static async Task<bool> PartMatchesSourceAsync(
        string sourcePath,
        ChunkProgress chunk,
        string partPath,
        int chunkSizeMB,
        CancellationToken cancellationToken)
    {
        var chunkSize = checked((long)Math.Max(1, chunkSizeMB) * 1024 * 1024);
        var offset = checked((long)(chunk.Index - 1) * chunkSize);
        await using var source = new FileStream(
            sourcePath, FileMode.Open, FileAccess.Read, FileShare.Read, 1 << 20, useAsync: true);
        await using var part = new FileStream(
            partPath, FileMode.Open, FileAccess.Read, FileShare.Read, 1 << 20, useAsync: true);
        source.Seek(offset, SeekOrigin.Begin);
        var sourceBuffer = new byte[1 << 20];
        var partBuffer = new byte[1 << 20];
        var remaining = chunk.Length;
        while (remaining > 0)
        {
            var wanted = (int)Math.Min(sourceBuffer.Length, remaining);
            var sourceRead = await source.ReadAsync(
                sourceBuffer.AsMemory(0, wanted), cancellationToken).ConfigureAwait(false);
            var partRead = await part.ReadAsync(
                partBuffer.AsMemory(0, wanted), cancellationToken).ConfigureAwait(false);
            if (sourceRead != partRead || sourceRead == 0 ||
                !sourceBuffer.AsSpan(0, sourceRead).SequenceEqual(partBuffer.AsSpan(0, partRead)))
            {
                return false;
            }
            remaining -= sourceRead;
        }
        return true;
    }
}
