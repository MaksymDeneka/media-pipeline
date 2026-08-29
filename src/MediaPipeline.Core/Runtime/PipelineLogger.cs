using System.Text;
using MediaPipeline.Core.IO;

namespace MediaPipeline.Core.Runtime;

public sealed class PipelineLogger(PipelinePaths paths)
{
    private readonly SemaphoreSlim _writeLock = new(1, 1);

    public Task InfoAsync(string message, CancellationToken cancellationToken = default) =>
        WriteAsync("INFO", message, cancellationToken);

    public Task WarningAsync(string message, CancellationToken cancellationToken = default) =>
        WriteAsync("WARN", message, cancellationToken);

    public Task ErrorAsync(string message, CancellationToken cancellationToken = default) =>
        WriteAsync("ERROR", message, cancellationToken);

    private async Task WriteAsync(
        string level,
        string message,
        CancellationToken cancellationToken)
    {
        var now = DateTimeOffset.Now;
        var line = $"[{now:yyyy-MM-dd HH:mm:ss}] [{level}] {message}";
        Console.WriteLine(line);

        await _writeLock.WaitAsync(cancellationToken);
        try
        {
            Directory.CreateDirectory(paths.Logs);
            var path = Path.Combine(paths.Logs, $"media-pipeline-{now:yyyyMMdd}.log");
            await File.AppendAllTextAsync(
                path,
                line + Environment.NewLine,
                new UTF8Encoding(false),
                cancellationToken);
        }
        finally
        {
            _writeLock.Release();
        }
    }
}
