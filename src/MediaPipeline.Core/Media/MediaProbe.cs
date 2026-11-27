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
    public bool HasTileGrid { get; init; }
    public bool IsHeif => MajorBrand?.Trim().ToLowerInvariant() is
        "heic" or "heix" or "hevc" or "hevx" or "heim" or "heis" or "mif1" or "msf1";
    public double? FrameCount { get; init; }
    public int? RotationDegrees { get; init; }
    public string? PixelFormat { get; init; }
    public string? ColorSpace { get; init; }
    public string? ColorTransfer { get; init; }
    public string? ColorPrimaries { get; init; }
    public string? ColorRange { get; init; }
    public bool IsHdr => ColorTransfer is "arib-std-b67" or "smpte2084";
}

public sealed class MediaProbe(Toolchain tools)
{
    public Task<MediaInfo> ReadAsync(string path, CancellationToken cancellationToken = default) =>
        ReadAsync(path, readFirstFrame: false, cancellationToken);

    public Task<MediaInfo> ReadImageAsync(string path, CancellationToken cancellationToken = default) =>
        ReadAsync(path, readFirstFrame: true, cancellationToken);

    private async Task<MediaInfo> ReadAsync(string path, bool readFirstFrame, CancellationToken cancellationToken)
    {
        const string streamEntries = "stream=codec_type,codec_name,width,height,duration,nb_frames,pix_fmt,color_space,color_transfer,color_primaries,color_range:stream_tags=rotate:stream_side_data=rotation:stream_group=type:stream_group_component=width,height:stream_group_side_data=rotation:format=duration,format_name:format_tags=major_brand";
        string[] frameArgs = readFirstFrame ? ["-select_streams", "v:0", "-read_intervals", "%+#1"] : [];
        var result = await ProcessRunner.RunAsync(
            tools.FFprobe,
            [
                "-v",
                "error",
                .. frameArgs,
                "-show_entries",
                readFirstFrame ? $"{streamEntries}:frame_side_data=rotation" : streamEntries,
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
            var hasTileGrid = false;
            int? gridRotation = null;
            if (root.TryGetProperty("stream_groups", out var groups))
            {
                foreach (var group in groups.EnumerateArray())
                {
                    if (ReadString(group, "type") != "Tile Grid") continue;
                    hasTileGrid = true;
                    // A HEIF stream is only one tile. Validate the primary grid's
                    // dimensions before asking a full-image decoder to allocate it.
                    if (group.TryGetProperty("components", out var components))
                    {
                        foreach (var component in components.EnumerateArray())
                        {
                            width = ReadInt(component, "width") ?? width;
                            height = ReadInt(component, "height") ?? height;
                            gridRotation = ReadRotation(component);
                            break;
                        }
                    }
                    break;
                }
            }
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
            // JPEG EXIF orientation is exposed on the decoded frame. FFmpeg
            // applies it before the crop filter receives the image.
            var frameRotation = root.TryGetProperty("frames", out var frames) && frames.GetArrayLength() > 0
                ? ReadRotation(frames[0])
                : null;
            var rotation = gridRotation ?? frameRotation ?? (videoStream.HasValue ? ReadRotation(videoStream.Value) : null);

            return new MediaInfo
            {
                Width = width,
                Height = height,
                DurationSeconds = duration,
                FormatName = formatName,
                CodecName = codecName,
                MajorBrand = majorBrand,
                HasTileGrid = hasTileGrid,
                FrameCount = frameCount,
                RotationDegrees = rotation,
                PixelFormat = videoStream.HasValue ? ReadString(videoStream.Value, "pix_fmt") : null,
                ColorSpace = videoStream.HasValue ? ReadString(videoStream.Value, "color_space") : null,
                ColorTransfer = videoStream.HasValue ? ReadString(videoStream.Value, "color_transfer") : null,
                ColorPrimaries = videoStream.HasValue ? ReadString(videoStream.Value, "color_primaries") : null,
                ColorRange = videoStream.HasValue ? ReadString(videoStream.Value, "color_range") : null,
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
