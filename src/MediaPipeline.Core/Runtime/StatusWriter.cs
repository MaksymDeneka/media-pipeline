using System.Text.Json;
using System.Text.Json.Serialization;
using MediaPipeline.Core.Contracts;
using MediaPipeline.Core.IO;

namespace MediaPipeline.Core.Runtime;

public sealed class StatusWriter(PipelinePaths paths)
{
    private static readonly JsonSerializerOptions Options = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    };

    public async Task WriteAsync(WorkerStatus status, CancellationToken cancellationToken = default)
    {
        Directory.CreateDirectory(paths.Status);
        var temporary = paths.StatusFile + $".{Environment.ProcessId}.tmp";

        await using (var stream = new FileStream(
                         temporary,
                         FileMode.Create,
                         FileAccess.Write,
                         FileShare.None,
                         bufferSize: 16 * 1024,
                         FileOptions.Asynchronous))
        {
            await JsonSerializer.SerializeAsync(stream, status, Options, cancellationToken);
            await stream.FlushAsync(cancellationToken);
        }

        File.Move(temporary, paths.StatusFile, overwrite: true);
    }
}
