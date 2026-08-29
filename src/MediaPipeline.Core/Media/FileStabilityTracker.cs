namespace MediaPipeline.Core.Media;

/// <summary>
/// A candidate becomes ready only after two observations with the same length and write time.
/// This catches downloads that keep an old timestamp while their length is still changing.
/// </summary>
public sealed class FileStabilityTracker
{
    private readonly Dictionary<string, Observation> _observations = new(
        OperatingSystem.IsWindows()
            ? StringComparer.OrdinalIgnoreCase
            : StringComparer.Ordinal);

    public StabilityState Observe(
        string path,
        TimeSpan stableFor,
        TimeSpan timeout,
        DateTimeOffset observedAt)
    {
        FileInfo file;
        try
        {
            file = new FileInfo(path);
            if (!file.Exists)
            {
                _observations.Remove(path);
                return StabilityState.Waiting;
            }
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            return TimedOut(path, timeout, observedAt);
        }

        var current = new Observation(
            file.Length,
            file.LastWriteTimeUtc,
            observedAt,
            observedAt);
        if (_observations.TryGetValue(path, out var previous) && previous.Rejected)
        {
            if (previous.Length != current.Length ||
                previous.LastWriteUtc != current.LastWriteUtc)
            {
                _observations[path] = current with { Rejected = true };
                return StabilityState.Rejected;
            }

            if (file.Length > 0 && observedAt - previous.UnchangedSince >= stableFor &&
                CanOpenExclusively(path))
            {
                _observations[path] = current;
                return StabilityState.Waiting;
            }

            return StabilityState.Rejected;
        }

        if (previous is null ||
            previous.Length != current.Length ||
            previous.LastWriteUtc != current.LastWriteUtc)
        {
            _observations[path] = current with
            {
                FirstSeen = previous is { Rejected: false } ? previous.FirstSeen : observedAt,
            };
            return TimedOut(path, timeout, observedAt);
        }

        if (file.Length > 0 && observedAt - previous.UnchangedSince >= stableFor &&
            CanOpenExclusively(path))
        {
            return StabilityState.Ready;
        }

        return TimedOut(path, timeout, observedAt);
    }

    public bool IsReady(string path, TimeSpan stableFor, DateTimeOffset observedAt) =>
        Observe(path, stableFor, Timeout.InfiniteTimeSpan, observedAt) == StabilityState.Ready;

    public void ForgetMissingFiles(IEnumerable<string> existingPaths)
    {
        var existing = existingPaths.ToHashSet(_observations.Comparer);
        foreach (var path in _observations.Keys.Where(path => !existing.Contains(path)).ToArray())
        {
            _observations.Remove(path);
        }
    }

    public void ForgetMissingFilesInDirectory(
        string directory,
        IEnumerable<string> existingPaths)
    {
        var existing = existingPaths.ToHashSet(_observations.Comparer);
        foreach (var path in _observations.Keys.Where(path =>
            _observations.Comparer.Equals(Path.GetDirectoryName(path), directory) &&
            !existing.Contains(path)).ToArray())
        {
            _observations.Remove(path);
        }
    }

    public void Forget(string path) => _observations.Remove(path);

    private static bool CanOpenExclusively(string path)
    {
        try
        {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None);
            return stream.Length > 0;
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            return false;
        }
    }

    private StabilityState TimedOut(string path, TimeSpan timeout, DateTimeOffset observedAt)
    {
        if (timeout == Timeout.InfiniteTimeSpan ||
            !_observations.TryGetValue(path, out var observation) ||
            observedAt - observation.FirstSeen < timeout)
        {
            return StabilityState.Waiting;
        }

        _observations[path] = observation with { Rejected = true };
        return StabilityState.TimedOut;
    }

    private sealed record Observation(
        long Length,
        DateTime LastWriteUtc,
        DateTimeOffset UnchangedSince,
        DateTimeOffset FirstSeen,
        bool Rejected = false);
}

public enum StabilityState
{
    Waiting,
    Ready,
    TimedOut,
    Rejected,
}
