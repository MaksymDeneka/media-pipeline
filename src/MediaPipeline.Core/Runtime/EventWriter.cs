using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using MediaPipeline.Core.Contracts;
using MediaPipeline.Core.IO;

namespace MediaPipeline.Core.Runtime;

public sealed class EventWriter(PipelinePaths paths)
{
    private static readonly JsonSerializerOptions Options = new()
    {
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    };

    private readonly SemaphoreSlim _writeLock = new(1, 1);
    private readonly Dictionary<string, PipelineEvent> _activeJobs = new(
        StringComparer.OrdinalIgnoreCase);

    public async Task AppendAsync(PipelineEvent pipelineEvent, CancellationToken cancellationToken = default)
    {
        var path = Path.Combine(paths.Logs, $"events-{pipelineEvent.Timestamp.ToLocalTime():yyyyMMdd}.jsonl");
        var line = JsonSerializer.Serialize(pipelineEvent, Options) + Environment.NewLine;

        await _writeLock.WaitAsync(cancellationToken);
        try
        {
            if (ApplyActiveState(pipelineEvent))
            {
                await WriteActiveJobsAsync(cancellationToken);
            }
            Directory.CreateDirectory(paths.Logs);
            await File.AppendAllTextAsync(path, line, new UTF8Encoding(false), cancellationToken);
        }
        finally
        {
            _writeLock.Release();
        }
    }

    private bool ApplyActiveState(PipelineEvent pipelineEvent)
    {
        if (pipelineEvent.Name is "watcher.start" or "watcher.stop")
        {
            _activeJobs.Clear();
            return true;
        }

        if (pipelineEvent.JobId is null)
        {
            return false;
        }

        if (pipelineEvent.Name == "job.start")
        {
            _activeJobs[pipelineEvent.JobId] = pipelineEvent;
            return true;
        }

        if (pipelineEvent.Name is "job.done" or "job.failed" or "job.cancelled")
        {
            _activeJobs.Remove(pipelineEvent.JobId);
            return true;
        }

        return false;
    }

    private async Task WriteActiveJobsAsync(CancellationToken cancellationToken)
    {
        Directory.CreateDirectory(paths.Status);
        var temporary = paths.ActiveJobsFile + $".{Environment.ProcessId}.tmp";
        await File.WriteAllTextAsync(
            temporary,
            JsonSerializer.Serialize(_activeJobs.Values, Options),
            new UTF8Encoding(false),
            cancellationToken);
        File.Move(temporary, paths.ActiveJobsFile, overwrite: true);
    }
}
