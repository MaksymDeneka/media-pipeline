using System.Text;
using MediaPipeline.Core.IO;

namespace MediaPipeline.Core.Runtime;

/// <summary>
/// A cross-platform single-instance lock. The open file handle owns the lock, so a crash releases
/// it without relying on a Windows mutex name or a stale process identifier.
/// </summary>
public sealed class WorkerLock : IDisposable
{
    private readonly FileStream _stream;

    private WorkerLock(FileStream stream) => _stream = stream;

    public static WorkerLock? TryAcquire(PipelinePaths paths)
    {
        Directory.CreateDirectory(paths.Status);

        try
        {
            var stream = new FileStream(
                paths.WorkerLockFile,
                FileMode.OpenOrCreate,
                FileAccess.ReadWrite,
                FileShare.None);
            stream.SetLength(0);
            stream.Write(Encoding.UTF8.GetBytes(Environment.ProcessId.ToString()));
            stream.Flush(flushToDisk: true);
            return new WorkerLock(stream);
        }
        catch (IOException)
        {
            return null;
        }
    }

    public static bool IsHeld(PipelinePaths paths)
    {
        if (!File.Exists(paths.WorkerLockFile))
        {
            return false;
        }

        try
        {
            using var stream = new FileStream(
                paths.WorkerLockFile,
                FileMode.Open,
                FileAccess.Read,
                FileShare.ReadWrite | FileShare.Delete);
            return false;
        }
        catch (IOException)
        {
            return true;
        }
    }

    public void Dispose() => _stream.Dispose();
}
