namespace MediaPipeline.Core.Media;

public enum MediaKind
{
    Unsupported,
    Video,
    Image,
    Temporary,
}

public static class MediaClassifier
{
    private static readonly HashSet<string> VideoExtensions = new(
        [".mp4", ".mov", ".mkv", ".webm", ".avi"],
        StringComparer.OrdinalIgnoreCase);

    private static readonly HashSet<string> ImageExtensions = new(
        [".jpg", ".jpeg", ".png", ".webp", ".heic", ".heif"],
        StringComparer.OrdinalIgnoreCase);

    private static readonly HashSet<string> TemporaryExtensions = new(
        [".crdownload", ".tmp", ".part", ".download"],
        StringComparer.OrdinalIgnoreCase);

    public static MediaKind Classify(string path)
    {
        var extension = Path.GetExtension(path);
        if (TemporaryExtensions.Contains(extension))
        {
            return MediaKind.Temporary;
        }

        if (VideoExtensions.Contains(extension))
        {
            return MediaKind.Video;
        }

        return ImageExtensions.Contains(extension)
            ? MediaKind.Image
            : MediaKind.Unsupported;
    }

    public static bool IsHeic(string path) =>
        Path.GetExtension(path).Equals(".heic", StringComparison.OrdinalIgnoreCase) ||
        Path.GetExtension(path).Equals(".heif", StringComparison.OrdinalIgnoreCase);
}
