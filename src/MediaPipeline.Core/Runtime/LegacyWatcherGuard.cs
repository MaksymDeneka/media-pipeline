using System.Security.Cryptography;
using System.Text;

namespace MediaPipeline.Core.Runtime;

public sealed class LegacyWatcherGuard : IDisposable
{
    private readonly Mutex? _mutex;

    private LegacyWatcherGuard(Mutex? mutex)
    {
        _mutex = mutex;
    }

    /// <summary>
    /// Acquires the same root-scoped named mutex as the PowerShell watcher. Holding this lease
    /// closes both start orders: native-first and legacy-first.
    /// </summary>
    public static LegacyWatcherGuard? TryAcquire(string pipelineRoot)
    {
        if (!OperatingSystem.IsWindows())
        {
            return new LegacyWatcherGuard(null);
        }

        // The object's kernel lifetime is the cross-runtime claim. Do not take thread ownership:
        // async continuations can dispose on another thread, while named-object existence still
        // makes the legacy watcher's createdNew check reject a duplicate atomically.
        var mutex = new Mutex(initiallyOwned: false, MutexName(pipelineRoot), out var createdNew);
        if (!createdNew)
        {
            mutex.Dispose();
            return null;
        }

        return new LegacyWatcherGuard(mutex);
    }

    public static bool IsRunning(string pipelineRoot)
    {
        if (!OperatingSystem.IsWindows())
        {
            return false;
        }

        try
        {
            using var mutex = Mutex.OpenExisting(MutexName(pipelineRoot));
            return true;
        }
        catch (WaitHandleCannotBeOpenedException)
        {
            return false;
        }
        catch (UnauthorizedAccessException)
        {
            return true;
        }
    }

    private static string MutexName(string root)
    {
        var normalized = Path.GetFullPath(ExpandHome(root))
            .TrimEnd('\\', '/')
            .ToLowerInvariant();
        if (normalized == @"d:\mediapipeline")
        {
            return @"Global\MediaPipelineWatcher";
        }

        var hash = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(normalized)))[..16];
        return $@"Global\MediaPipelineWatcher_{hash}";
    }

    private static string ExpandHome(string path)
    {
        if (path == "~")
        {
            return Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        }
        if (path.StartsWith("~/", StringComparison.Ordinal) ||
            path.StartsWith("~\\", StringComparison.Ordinal))
        {
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                path[2..]);
        }
        return path;
    }

    public void Dispose()
    {
        _mutex?.Dispose();
    }
}
