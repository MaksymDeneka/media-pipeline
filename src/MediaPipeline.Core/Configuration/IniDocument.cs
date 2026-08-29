namespace MediaPipeline.Core.Configuration;

/// <summary>
/// Reads the existing comment-heavy INI format during the worker migration. Ordinary section
/// headings are decorative. Only a heading beginning with "preset " changes key scope.
/// </summary>
public sealed class IniDocument
{
    private IniDocument(
        IReadOnlyDictionary<string, string> globals,
        IReadOnlyDictionary<string, IReadOnlyDictionary<string, string>> presets)
    {
        Globals = globals;
        Presets = presets;
    }

    public IReadOnlyDictionary<string, string> Globals { get; }

    public IReadOnlyDictionary<string, IReadOnlyDictionary<string, string>> Presets { get; }

    public static IniDocument Load(string path) => Parse(File.ReadLines(path));

    public static IniDocument Parse(IEnumerable<string> lines)
    {
        var globals = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var presets = new Dictionary<string, IReadOnlyDictionary<string, string>>(
            StringComparer.OrdinalIgnoreCase);
        Dictionary<string, string>? currentPreset = null;

        foreach (var rawLine in lines)
        {
            var line = rawLine.Trim();
            if (line.Length == 0 || line[0] is ';' or '#')
            {
                continue;
            }

            if (line[0] == '[')
            {
                var header = line.Trim('[', ']').Trim();
                if (header.StartsWith("preset ", StringComparison.OrdinalIgnoreCase))
                {
                    var name = header[7..].Trim();
                    if (name.Length == 0)
                    {
                        currentPreset = null;
                        continue;
                    }

                    if (!presets.TryGetValue(name, out var existing))
                    {
                        currentPreset = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        presets[name] = currentPreset;
                    }
                    else
                    {
                        currentPreset = (Dictionary<string, string>)existing;
                    }
                }
                else
                {
                    currentPreset = null;
                }

                continue;
            }

            var equals = line.IndexOf('=');
            if (equals < 1)
            {
                continue;
            }

            var key = line[..equals].Trim();
            if (key.Length == 0)
            {
                continue;
            }

            var value = CleanValue(line[(equals + 1)..]);
            (currentPreset ?? globals)[key] = value;
        }

        return new IniDocument(globals, presets);
    }

    public static string CleanValue(string raw)
    {
        var value = raw.Trim();
        char? quote = null;
        for (var index = 0; index < value.Length; index++)
        {
            var character = value[index];
            if (quote is not null)
            {
                if (character == quote)
                {
                    quote = null;
                }
                continue;
            }

            if (character is '\'' or '"')
            {
                quote = character;
                continue;
            }

            if (index > 0 && char.IsWhiteSpace(value[index - 1]) && character is ';' or '#')
            {
                value = value[..index].TrimEnd();
                break;
            }
        }

        if (value.Length >= 2 &&
            ((value[0] == '"' && value[^1] == '"') ||
             (value[0] == '\'' && value[^1] == '\'')))
        {
            return value[1..^1];
        }

        return value;
    }
}
