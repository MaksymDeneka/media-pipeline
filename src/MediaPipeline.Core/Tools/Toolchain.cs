using System.Runtime.InteropServices;

namespace MediaPipeline.Core.Tools;

public sealed record Toolchain
{
    public required string FFmpeg { get; init; }
    public required string FFprobe { get; init; }
    public string? ExifTool { get; init; }

    public static Toolchain Discover(string? applicationDirectory = null)
    {
        applicationDirectory ??= AppContext.BaseDirectory;

        return new Toolchain
        {
            FFmpeg = FindRequired("ffmpeg", applicationDirectory),
            FFprobe = FindRequired("ffprobe", applicationDirectory),
            // Heatup approach strips metadata with -map_metadata -1; ExifTool is legacy/optional.
            ExifTool = FindOptional("exiftool", applicationDirectory),
        };
    }

    public static string? FindOptional(string name, string? applicationDirectory = null)
    {
        try
        {
            return FindRequired(name, applicationDirectory);
        }
        catch (FileNotFoundException)
        {
            return null;
        }
    }

    public static string FindRequired(string name, string? applicationDirectory = null)
    {
        applicationDirectory ??= AppContext.BaseDirectory;
        var executableName = OperatingSystem.IsWindows() ? name + ".exe" : name;
        var runtime = RuntimeInformation.RuntimeIdentifier;
        var candidates = new List<string>
        {
            Path.Combine(applicationDirectory, "tools", runtime, executableName),
            Path.Combine(applicationDirectory, "tools", executableName),
        };

        var path = Environment.GetEnvironmentVariable("PATH") ?? "";
        candidates.AddRange(
            path.Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Select(directory => Path.Combine(directory, executableName)));

        if (OperatingSystem.IsWindows())
        {
            var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            candidates.Add(Path.Combine(localAppData, "Microsoft", "WinGet", "Links", executableName));
            candidates.Add(Path.Combine(@"C:\Tools", name, name == "exiftool" ? "exiftool.exe" : $"bin\\{executableName}"));
        }
        else
        {
            candidates.Add(Path.Combine("/opt/homebrew/bin", executableName));
            candidates.Add(Path.Combine("/usr/local/bin", executableName));
            candidates.Add(Path.Combine("/usr/bin", executableName));
        }

        var found = candidates.FirstOrDefault(File.Exists);
        return found is not null
            ? Path.GetFullPath(found)
            : throw new FileNotFoundException(
                $"Required tool '{name}' was not found in the app bundle or PATH.");
    }
}
