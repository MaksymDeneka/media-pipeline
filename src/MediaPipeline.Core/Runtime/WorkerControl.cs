using MediaPipeline.Core.IO;

namespace MediaPipeline.Core.Runtime;

public sealed class WorkerControl(PipelinePaths paths)
{
    public bool StopRequested => Exists("stop");

    public bool PausedAll => Exists("pause");

    public bool IsPaused(string preset, string workspace) =>
        PausedAll || Exists($"pause.{preset}") || Exists($"pause.{preset}.{workspace}");

    /// <summary>
    /// Aborts the current file or copy when the operator paused this lane or asked
    /// the worker to stop. Checked per file and per copy inside batch processing so
    /// pause/stop take effect within one encode instead of after the whole backlog.
    /// Throws <see cref="LaneInterruption"/> (not a bare cancellation) so engine
    /// timeouts and genuine shutdown stay distinguishable downstream.
    /// </summary>
    public void ThrowIfInterrupted(string preset, string workspace)
    {
        if (StopRequested || IsPaused(preset, workspace))
        {
            throw new LaneInterruption();
        }
    }

    public void ClearStopRequest()
    {
        var path = ControlPath("stop");
        if (File.Exists(path))
        {
            File.Delete(path);
        }
    }

    public void RequestStop() => SetFlag("stop", true);

    public void SetPaused(bool paused, string? preset = null, string? workspace = null)
    {
        if (workspace is not null && preset is null)
        {
            throw new ArgumentException("A workspace pause also requires a preset.");
        }

        ValidateScope(preset, nameof(preset));
        ValidateScope(workspace, nameof(workspace));
        var name = preset is null
            ? "pause"
            : workspace is null
                ? $"pause.{preset}"
                : $"pause.{preset}.{workspace}";
        SetFlag(name, paused);
    }

    public int RequeueFailed(string preset, string workspace)
    {
        ValidateScope(preset, nameof(preset));
        ValidateScope(workspace, nameof(workspace));
        var lane = paths.Lane(preset, workspace);
        var moved = MediaRequestQueue.RequeueFailed(lane);
        if (!Directory.Exists(lane.Failed))
        {
            return moved;
        }

        Directory.CreateDirectory(lane.Input);
        foreach (var source in Directory.EnumerateFiles(lane.Failed))
        {
            try
            {
                var destination = OutputNameGenerator.UniqueDestination(
                    lane.Input,
                    Path.GetFileName(source));
                File.Move(source, destination);
                moved++;
            }
            catch (IOException)
            {
            }
        }

        return moved;
    }

    private void SetFlag(string name, bool enabled)
    {
        Directory.CreateDirectory(paths.Control);
        var path = ControlPath(name);
        if (enabled)
        {
            if (!File.Exists(path))
            {
                File.WriteAllText(path, "");
            }
        }
        else if (File.Exists(path))
        {
            File.Delete(path);
        }
    }

    private static void ValidateScope(string? value, string parameter)
    {
        if (value is not null)
        {
            PipelinePaths.ValidateSegment(value, parameter);
        }
    }

    private bool Exists(string name) => File.Exists(ControlPath(name));

    private string ControlPath(string name) => Path.Combine(paths.Control, name);
}

/// <summary>
/// Operator-requested interruption (pause flag or stop request) surfacing from
/// inside batch processing. Distinct from <see cref="OperationCanceledException"/>
/// so engine-internal timeouts (converted to data failures) and genuine shutdown
/// (console token) can never be mistaken for a pause.
/// </summary>
public sealed class LaneInterruption(string? message = null)
    : OperationCanceledException(message ?? "Processing was interrupted.");
