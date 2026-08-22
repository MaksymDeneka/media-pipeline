using System.IO;
using System.Security.Cryptography;

namespace MediaPipelineTray.Services;

/// <summary>
/// Splits a file into fixed-size parts and hashes them.
///
/// Separate from <see cref="UploadService"/> so it can be exercised without a remote: the
/// split and the hashes are the part that would silently corrupt data if wrong, and testing
/// that should not require an SFTP host.
/// </summary>
public static class FileChunker
{
    public const string PartExtensionFormat = "{0}.part{1:D5}";

    public static string PartName(string fileName, int index) =>
        string.Format(PartExtensionFormat, fileName, index);

    /// <summary>Plans the parts for a file without writing anything.</summary>
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

        var chunkSize = (long)chunkSizeMB * 1024 * 1024;
        var count = (int)Math.Ceiling((double)source.Length / chunkSize);
        var chunks = new List<ChunkProgress>(count);

        for (var index = 0; index < count; index++)
        {
            var offset = (long)index * chunkSize;

            chunks.Add(new ChunkProgress
            {
                Index = index + 1,
                Length = Math.Min(chunkSize, source.Length - offset),
                FileName = PartName(source.Name, index + 1),
            });
        }

        return chunks;
    }

    /// <summary>
    /// Writes one part. A part that already exists at exactly the right length is left alone
    /// and reported as reused, which is what makes an interrupted split resumable.
    /// </summary>
    public static async Task<bool> WritePartAsync(
        string sourcePath,
        ChunkProgress chunk,
        string partPath,
        int chunkSizeMB,
        CancellationToken cancellationToken)
    {
        var existing = new FileInfo(partPath);
        if (existing.Exists && existing.Length == chunk.Length)
        {
            return false;
        }

        var chunkSize = (long)chunkSizeMB * 1024 * 1024;
        var offset = (long)(chunk.Index - 1) * chunkSize;

        await using var source = new FileStream(
            sourcePath, FileMode.Open, FileAccess.Read, FileShare.Read, 1 << 20, useAsync: true);

        source.Seek(offset, SeekOrigin.Begin);

        await using var target = new FileStream(
            partPath, FileMode.Create, FileAccess.Write, FileShare.None, 1 << 20, useAsync: true);

        var buffer = new byte[1 << 22];
        var remaining = chunk.Length;

        while (remaining > 0)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var wanted = (int)Math.Min(buffer.Length, remaining);
            var read = await source.ReadAsync(buffer.AsMemory(0, wanted), cancellationToken).ConfigureAwait(false);

            if (read == 0)
            {
                throw new IOException($"Source ended early while writing {chunk.FileName}.");
            }

            await target.WriteAsync(buffer.AsMemory(0, read), cancellationToken).ConfigureAwait(false);
            remaining -= read;
        }

        return true;
    }

    public static async Task<string> HashAsync(string path, CancellationToken cancellationToken)
    {
        await using var stream = new FileStream(
            path, FileMode.Open, FileAccess.Read, FileShare.Read, 1 << 20, useAsync: true);

        var hash = await SHA256.HashDataAsync(stream, cancellationToken).ConfigureAwait(false);
        return Convert.ToHexString(hash);
    }
}
