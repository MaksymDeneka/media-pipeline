using System.Globalization;
using System.Text.Json;
using MediaPipeline.Core.Tools;

namespace MediaPipeline.Core.Media;

public sealed record MediaInfo
{
    public int? Width { get; init; }
    public int? Height { get; init; }
    public double? DurationSeconds { get; init; }
    public string FormatName { get; init; } = "";
}

public sealed class MediaProbe(Toolchain tools)
{
    public async Task<MediaInfo> ReadAsync(string path, CancellationToken cancellationToken = default)
    {
        var result = await ProcessRunner.RunAsync(
            tools.FFprobe,
            [
                "-v",
                "error",
                "-select_streams",
                "v:0",
                "-read_intervals",
                "%+#1",
                "-show_frames",
                "-show_entries",
                "frame=width,height:frame_side_data=rotation:format=duration,format_name",
                "-of",
                "json",
                path,
            ],
            cancellationToken);

        if (!result.Succeeded)
        {
            var detail = result.CombinedOutput.Trim();
            throw new InvalidDataException(
                detail.Length > 0
                    ? $"FFprobe could not read '{path}': {detail}"
                    : $"FFprobe could not read '{path}'.");
        }

        using var document = JsonDocument.Parse(result.StandardOutput);
        var root = document.RootElement;
        var frame = root.TryGetProperty("frames", out var frames) &&
                    frames.ValueKind == JsonValueKind.Array &&
                    frames.GetArrayLength() > 0
            ? frames[0]
            : default;
        var format = root.TryGetProperty("format", out var formatElement)
            ? formatElement
            : default;

        var width = ReadInt(frame, "width");
        var height = ReadInt(frame, "height");
        var rotation = ReadRotation(frame);
        var normalizedRotation = (rotation % 360 + 360) % 360;
        if (width is not null && height is not null &&
            (Math.Abs(normalizedRotation - 90) < 0.5 ||
             Math.Abs(normalizedRotation - 270) < 0.5))
        {
            (width, height) = (height, width);
        }

        return new MediaInfo
        {
            Width = width,
            Height = height,
            DurationSeconds = ReadDouble(format, "duration"),
            FormatName = ReadString(format, "format_name") ?? "",
        };
    }

    private static double ReadRotation(JsonElement frame)
    {
        if (frame.ValueKind != JsonValueKind.Object ||
            !frame.TryGetProperty("side_data_list", out var sideData) ||
            sideData.ValueKind != JsonValueKind.Array)
        {
            return 0;
        }

        foreach (var item in sideData.EnumerateArray())
        {
            if (ReadDouble(item, "rotation") is { } rotation)
            {
                return rotation;
            }
        }

        return 0;
    }

    private static int? ReadInt(JsonElement element, string property)
    {
        if (element.ValueKind != JsonValueKind.Object || !element.TryGetProperty(property, out var value))
        {
            return null;
        }

        if (value.ValueKind == JsonValueKind.Number && value.TryGetInt32(out var number))
        {
            return number;
        }

        return value.ValueKind == JsonValueKind.String &&
               int.TryParse(value.GetString(), NumberStyles.Integer, CultureInfo.InvariantCulture, out number)
            ? number
            : null;
    }

    private static double? ReadDouble(JsonElement element, string property)
    {
        if (element.ValueKind != JsonValueKind.Object || !element.TryGetProperty(property, out var value))
        {
            return null;
        }

        if (value.ValueKind == JsonValueKind.Number && value.TryGetDouble(out var number))
        {
            return number;
        }

        return value.ValueKind == JsonValueKind.String &&
               double.TryParse(value.GetString(), NumberStyles.Float, CultureInfo.InvariantCulture, out number)
            ? number
            : null;
    }

    private static string? ReadString(JsonElement element, string property) =>
        element.ValueKind == JsonValueKind.Object &&
        element.TryGetProperty(property, out var value) &&
        value.ValueKind == JsonValueKind.String
            ? value.GetString()
            : null;
}
