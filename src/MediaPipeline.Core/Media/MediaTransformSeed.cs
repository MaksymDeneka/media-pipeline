using System.Security.Cryptography;
using System.Text;

namespace MediaPipeline.Core.Media;

/// <summary>
/// Deterministic seeded planning ported from heatup sidecar/media-transform/profiles.ts.
/// Seed is a lowercase hex SHA256 string. All variant differences derive from it plus
/// the per-copy ordinal, so reruns produce the same argv for the same inputs.
/// </summary>
public static class MediaTransformSeed
{
    public static string Derive(string sourceContentHash, int ordinal)
    {
        var input = $"media-pipeline-v2\0{sourceContentHash.ToLowerInvariant()}\0{ordinal}";
        var bytes = SHA256.HashData(Encoding.UTF8.GetBytes(input));
        return Convert.ToHexString(bytes).ToLowerInvariant();
    }

    public static string HashFile(string path)
    {
        using var stream = File.OpenRead(path);
        return Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant();
    }

    public static async Task<string> HashFileAsync(string path, CancellationToken cancellationToken = default)
    {
        await using var stream = File.OpenRead(path);
        var hash = await SHA256.HashDataAsync(stream, cancellationToken);
        return Convert.ToHexString(hash).ToLowerInvariant();
    }

    /// <summary>Port of heatup seededInteger(seed, min, max, wordOffset).</summary>
    public static int SeededInteger(string seed, int min, int max, int wordOffset)
    {
        var normalized = seed.Length >= 64 ? seed[..64] : seed.PadRight(64, '0');
        var start = (wordOffset * 8) % 56;
        var slice = normalized.Substring(start, 8);
        var value = Convert.ToUInt32(slice, 16);
        return min + (int)(value % (uint)(max - min + 1));
    }

    public static double SeededSignedFraction(string seed, int rangeMillionths, int wordOffset) =>
        (SeededInteger(seed, 0, rangeMillionths * 2, wordOffset) - rangeMillionths) / 1_000_000.0;
}
