using MediaPipeline.Core.IO;

namespace MediaPipeline.Core.Runtime;

public sealed class WorkerControl(PipelinePaths paths)
{
    public bool StopRequested => Exists("stop");

    public bool PausedAll => Exists("pause");

    public bool IsPaused(string preset, string workspace) =>
        PausedAll || Exists($"pause.{preset}") || Exists($"pause.{preset}.{workspace}");

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
        if (!Directory.Exists(lane.Failed))
        {
            return 0;
        }

        Directory.CreateDirectory(lane.Input);
        var moved = 0;
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
