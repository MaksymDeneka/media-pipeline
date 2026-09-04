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
    public string? CodecName { get; init; }
    public string? MajorBrand { get; init; }
    public double? FrameCount { get; init; }
    public int? RotationDegrees { get; init; }
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
                "-show_entries",
                "stream=codec_type,codec_name,width,height,duration,nb_frames:stream_tags=rotate:stream_side_data=rotation:format=duration,format_name:format_tags=major_brand",
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

        JsonDocument document;
        try
        {
            document = JsonDocument.Parse(result.StandardOutput);
        }
        catch (JsonException exception)
        {
            throw new InvalidDataException(
                $"FFprobe returned unparseable output for '{path}'. {exception.Message}", exception);
        }

        using (document)
        {
            var root = document.RootElement;

            JsonElement? videoStream = null;
            if (root.TryGetProperty("streams", out var streams) &&
                streams.ValueKind == JsonValueKind.Array)
            {
                foreach (var stream in streams.EnumerateArray())
                {
                    if (ReadString(stream, "codec_type") == "video")
                    {
                        videoStream = stream;
                        break;
                    }
                }
            }

            var format = root.TryGetProperty("format", out var formatElement)
                ? formatElement
                : default;

            var width = videoStream.HasValue ? ReadInt(videoStream.Value, "width") : null;
            var height = videoStream.HasValue ? ReadInt(videoStream.Value, "height") : null;
            var codecName = videoStream.HasValue ? ReadString(videoStream.Value, "codec_name") : null;
            var duration = ReadDouble(format, "duration")
                ?? (videoStream.HasValue ? ReadDouble(videoStream.Value, "duration") : null);
            var formatName = ReadString(format, "format_name") ?? "";
            string? majorBrand = null;
            if (format.ValueKind == JsonValueKind.Object &&
                format.TryGetProperty("tags", out var tags) &&
                tags.ValueKind == JsonValueKind.Object)
            {
                majorBrand = ReadString(tags, "major_brand");
            }

            double? frameCount = videoStream.HasValue ? ReadDouble(videoStream.Value, "nb_frames") : null;
            var rotation = videoStream.HasValue ? ReadRotation(videoStream.Value) : null;

            return new MediaInfo
            {
                Width = width,
                Height = height,
                DurationSeconds = duration,
                FormatName = formatName,
                CodecName = codecName,
                MajorBrand = majorBrand,
                FrameCount = frameCount,
                RotationDegrees = rotation,
            };
        }
    }

    /// <summary>
    /// Content-based identity ported from heatup detectMediaSourceIdentity.
    /// Filename is ignored; duration is validation evidence, never type evidence.
    /// </summary>
    public static (string MediaKind, string SourceExtension) DetectIdentity(MediaInfo probe)
    {
        var formatName = (probe.FormatName ?? "").ToLowerInvariant();
        var stillImageContainer = System.Text.RegularExpressions.Regex.IsMatch(
            $",{formatName},",
            ",(image2|image2pipe|png_pipe|jpeg_pipe|webp_pipe|apng|heif|avif),");
        var codecName = (probe.CodecName ?? "").ToLowerInvariant();
        var videoContainer = System.Text.RegularExpressions.Regex.IsMatch(
            $",{formatName},",
            ",(matroska|webm|mov|mp4|m4a|3gp|3g2|mj2|avi|flv|mpeg|mpegts|ogg),");
        var stillImageCodec = System.Text.RegularExpressions.Regex.IsMatch(
            codecName, "^(png|apng|webp|mjpeg|jpeg|heic|heif|av1|hevc)$");
        var majorBrand = (probe.MajorBrand ?? "").ToLowerInvariant();
        var stillImageBrand = System.Text.RegularExpressions.Regex.IsMatch(
            majorBrand, "^(avif|avis|heic|heix|hevc|hevx|heim|heis|mif1|msf1)$");
        var stillImageStream = probe.FrameCount == 1 && stillImageCodec && (!videoContainer || stillImageBrand);
        string mediaKind;
        if (stillImageContainer || stillImageStream)
        {
            mediaKind = "image";
        }
        else if (videoContainer || (probe.FrameCount ?? 0) > 1)
        {
            mediaKind = "video";
        }
        else
        {
            mediaKind = stillImageCodec ? "image" : "video";
        }

        var sourceExtension = mediaKind == "video"
            ? ".mp4"
            : codecName is "png" or "apng" ? ".png"
            : codecName == "webp" ? ".webp"
            : ".jpg";
        return (mediaKind, sourceExtension);
    }

    private static int? ReadRotation(JsonElement stream)
    {
        object? raw = null;
        if (stream.ValueKind == JsonValueKind.Object)
        {
            if (stream.TryGetProperty("side_data_list", out var sideData) &&
                sideData.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in sideData.EnumerateArray())
                {
                    if (ReadDouble(item, "rotation") is { } rotation)
                    {
                        raw = rotation;
                        break;
                    }
                }
            }

            raw ??= stream.TryGetProperty("tags", out var tags) &&
                    tags.ValueKind == JsonValueKind.Object &&
                    ReadString(tags, "rotate") is { } rotateTag &&
                    double.TryParse(rotateTag, NumberStyles.Float, CultureInfo.InvariantCulture, out var rotate)
                ? rotate
                : null;
        }

        return MediaTransformPlanner.NormalizeRotationDegrees(raw);
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
