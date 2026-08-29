using System.Security.Cryptography;

namespace MediaPipeline.Core.IO;

public static class OutputNameGenerator
{
    private static readonly string[] Descriptors =
    [
        "autumn",
        "bright",
        "calm",
        "cedar",
        "clear",
        "coastal",
        "daily",
        "evening",
        "fresh",
        "garden",
        "golden",
        "harbor",
        "local",
        "maple",
        "meadow",
        "modern",
        "morning",
        "natural",
        "open",
        "quiet",
        "river",
        "silver",
        "simple",
        "spring",
        "studio",
        "summer",
        "sunny",
        "travel",
        "urban",
        "warm",
        "weekend",
        "winter",
    ];

    private static readonly string[] Subjects =
    [
        "album",
        "capture",
        "clip",
        "collection",
        "frame",
        "gallery",
        "image",
        "media",
        "memory",
        "moment",
        "photo",
        "picture",
        "post",
        "project",
        "scene",
        "shot",
        "snapshot",
        "story",
        "take",
        "update",
        "upload",
        "video",
        "view",
        "work",
    ];

    private static readonly string[] Contexts =
    [
        "archive",
        "backup",
        "camera",
        "desktop",
        "draft",
        "edit",
        "export",
        "folder",
        "home",
        "inbox",
        "library",
        "mobile",
        "notes",
        "phone",
        "review",
        "share",
        "social",
        "temp",
        "today",
        "trip",
        "week",
        "workshop",
    ];

    private static readonly string[] Separators = ["-", "_", " "];

    public static string NewFilePath(string directory, string extension)
    {
        var normalizedExtension = extension.StartsWith('.') ? extension : "." + extension;
        normalizedExtension = normalizedExtension.ToUpperInvariant();

        while (true)
        {
            var number = RandomNumberGenerator.GetInt32(1, 10_000);
            var path = Path.Combine(directory, $"IMG_{number:D4}{normalizedExtension}");
            if (!File.Exists(path))
            {
                return path;
            }
        }
    }

    public static string NewDirectory(string parent)
    {
        while (true)
        {
            var path = Path.Combine(parent, NewRegularName());
            if (!Directory.Exists(path))
            {
                return Directory.CreateDirectory(path).FullName;
            }
        }
    }

    public static string UniqueDestination(string directory, string originalFileName)
    {
        var destination = Path.Combine(directory, originalFileName);
        if (!File.Exists(destination) && !Directory.Exists(destination))
        {
            return destination;
        }

        var baseName = Path.GetFileNameWithoutExtension(originalFileName);
        var extension = Path.GetExtension(originalFileName);
        do
        {
            destination = Path.Combine(directory, $"{baseName}-{NewRegularName()}{extension}");
        }
        while (File.Exists(destination) || Directory.Exists(destination));

        return destination;
    }

    public static string NewRegularName()
    {
        var capitalize = RandomNumberGenerator.GetInt32(2) == 1;
        var descriptor = Pick(Descriptors, capitalize);
        var subject = Pick(Subjects, capitalize);
        var context = Pick(Contexts, capitalize);
        var number = NewNumber();

        return RandomNumberGenerator.GetInt32(12) switch
        {
            0 => Join(descriptor, subject),
            1 => Join(subject, number),
            2 => Join(descriptor, subject, number),
            3 => Join(context, subject),
            4 => Join(subject, context, number),
            5 => Join(descriptor, context, subject),
            6 => Join(context, number),
            7 => Join(subject, descriptor),
            8 => Join(context, descriptor, number),
            9 => Join(descriptor, number),
            10 => Join(subject, context),
            _ => Join(descriptor, context, subject, number),
        };
    }

    private static string Pick(IReadOnlyList<string> values, bool capitalize)
    {
        var value = values[RandomNumberGenerator.GetInt32(values.Count)];
        return capitalize ? char.ToUpperInvariant(value[0]) + value[1..] : value;
    }

    private static string NewNumber()
    {
        var digits = RandomNumberGenerator.GetInt32(2, 7);
        var minimum = (int)Math.Pow(10, digits - 1);
        var maximum = (int)Math.Pow(10, digits);
        return RandomNumberGenerator.GetInt32(minimum, maximum).ToString();
    }

    private static string Join(params string[] parts) =>
        string.Join(Separators[RandomNumberGenerator.GetInt32(Separators.Length)], parts);
}
