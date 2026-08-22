using System.IO;
using System.Text;
using System.Text.Json;
using MediaPipelineTray.Models;

namespace MediaPipelineTray.Services;

/// <summary>
/// Tails logs\events-YYYYMMDD.jsonl, returning only lines added since the last call.
///
/// Reads from a remembered byte offset rather than re-parsing the file, so cost stays flat as
/// the day's file grows. Opened with full sharing because the watcher appends to it
/// continuously and must never be blocked by the UI reading.
///
/// Handles two awkward cases: the file rolling over at midnight, and a partial final line when
/// a read lands mid-append.
/// </summary>
public sealed class EventStreamReader
{
    private static readonly JsonSerializerOptions Options = new()
    {
        PropertyNameCaseInsensitive = true,
    };

    private readonly PipelinePaths _paths;
    private string? _currentFile;
    private long _offset;

    public EventStreamReader(PipelinePaths paths) => _paths = paths;

    /// <summary>
    /// Positions at the end of today's file without returning anything, so a freshly started UI
    /// does not replay the whole day as if it had just happened.
    /// </summary>
    public void SeekToEnd()
    {
        var file = _paths.EventFileFor(DateTimeOffset.Now);
        _currentFile = file;
        _offset = File.Exists(file) ? new FileInfo(file).Length : 0;
    }

    /// <summary>
    /// Reads everything appended since the previous call. Returns an empty list when there is
    /// nothing new, which is the common case.
    /// </summary>
    public IReadOnlyList<PipelineEvent> ReadNew()
    {
        var file = _paths.EventFileFor(DateTimeOffset.Now);

        if (!string.Equals(file, _currentFile, StringComparison.OrdinalIgnoreCase))
        {
            // Past midnight the watcher starts a new file; begin at the top of it.
            _currentFile = file;
            _offset = 0;
        }

        if (!File.Exists(file))
        {
            return [];
        }

        var length = new FileInfo(file).Length;
        if (length < _offset)
        {
            // The file shrank, so it was replaced rather than appended to.
            _offset = 0;
        }

        if (length == _offset)
        {
            return [];
        }

        var events = new List<PipelineEvent>();

        try
        {
            using var stream = new FileStream(
                file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);

            stream.Seek(_offset, SeekOrigin.Begin);

            using var reader = new StreamReader(stream, Encoding.UTF8);
            var consumed = _offset;

            while (reader.ReadLine() is { } line)
            {
                var bytes = Encoding.UTF8.GetByteCount(line) + Environment.NewLine.Length;

                // A line without a terminator may still be being written. Leave the offset
                // before it so the next pass re-reads it whole.
                if (consumed + bytes > length)
                {
                    break;
                }

                consumed += bytes;

                if (line.Trim().Length == 0)
                {
                    continue;
                }

                try
                {
                    var parsed = JsonSerializer.Deserialize<PipelineEvent>(line, Options);
                    if (parsed is not null && parsed.Name.Length > 0)
                    {
                        events.Add(parsed);
                    }
                }
                catch (JsonException)
                {
                    // A malformed line is not worth losing the rest of the batch over.
                }
            }

            _offset = consumed;
        }
        catch (IOException)
        {
            // Locked or vanished mid-read; the next tick tries again.
        }

        return events;
    }
}
