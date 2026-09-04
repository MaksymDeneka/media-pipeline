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
    // Lane admission is filename-based (cheap); final routing always uses probed
    // content (see MediaProbe.DetectIdentity). Keep this list broad so valid
    // heatup-style containers are admitted to a lane instead of silently skipped.
    private static readonly HashSet<string> VideoExtensions = new(
        [".mp4", ".mov", ".m4v", ".mkv", ".webm", ".avi", ".flv", ".m4a", ".3gp",
         ".3g2", ".mpeg", ".mpg", ".mpegts", ".ts", ".ogg", ".ogv", ".mts"],
        StringComparer.OrdinalIgnoreCase);

    private static readonly HashSet<string> ImageExtensions = new(
        [".jpg", ".jpeg", ".png", ".webp", ".heic", ".heif", ".avif", ".apng"],
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
