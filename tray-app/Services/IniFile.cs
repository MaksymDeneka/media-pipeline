using System.IO;
using System.Text;

namespace MediaPipelineTray.Services;

/// <summary>
/// A comment-preserving reader and writer for config.ini.
///
/// The file is heavily commented and those comments are the documentation, so this keeps the
/// original lines and edits values in place rather than parsing to a model and reserialising.
/// Setting a key that already exists rewrites just that line, including its trailing comment.
///
/// Parsing deliberately matches Read-IniDocument in watch-media.ps1: ordinary section headers
/// are decorative and their keys are global, while a [preset name] header scopes the keys that
/// follow it to that preset.
/// </summary>
public sealed class IniFile
{
    private readonly List<string> _lines;

    private IniFile(List<string> lines) => _lines = lines;

    public static IniFile Load(string path) =>
        new(File.Exists(path) ? [.. File.ReadAllLines(path)] : []);

    /// <summary>Strips a quoted wrapper, or an unquoted trailing comment, like the watcher does.</summary>
    public static string CleanValue(string raw)
    {
        var value = raw.Trim();

        if (value.Length >= 2 &&
            ((value[0] == '"' && value[^1] == '"') || (value[0] == '\'' && value[^1] == '\'')))
        {
            return value[1..^1];
        }

        // An inline comment starts at the first whitespace followed by ; or #.
        for (var i = 1; i < value.Length; i++)
        {
            if (char.IsWhiteSpace(value[i - 1]) && (value[i] == ';' || value[i] == '#'))
            {
                return value[..(i - 1)].TrimEnd();
            }
        }

        return value;
    }

    private static bool TryReadHeader(string line, out string? presetName)
    {
        presetName = null;

        var trimmed = line.Trim();
        if (trimmed.Length < 2 || trimmed[0] != '[')
        {
            return false;
        }

        var header = trimmed.Trim('[', ']').Trim();
        if (header.StartsWith("preset ", StringComparison.OrdinalIgnoreCase))
        {
            presetName = header[7..].Trim();
        }

        return true;
    }

    private static bool TryReadPair(string line, out string key, out string value)
    {
        key = "";
        value = "";

        var trimmed = line.Trim();
        if (trimmed.Length == 0 || trimmed[0] is '#' or ';' or '[')
        {
            return false;
        }

        var equals = trimmed.IndexOf('=');
        if (equals < 1)
        {
            return false;
        }

        key = trimmed[..equals].Trim();
        value = CleanValue(trimmed[(equals + 1)..]);
        return key.Length > 0;
    }

    /// <summary>Global settings, meaning every key outside a [preset ...] section.</summary>
    public Dictionary<string, string> ReadGlobals()
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        string? preset = null;

        foreach (var line in _lines)
        {
            if (TryReadHeader(line, out var name))
            {
                preset = name;
                continue;
            }

            if (preset is null && TryReadPair(line, out var key, out var value))
            {
                result[key] = value;
            }
        }

        return result;
    }

    /// <summary>Each preset's overrides, in the order they appear in the file.</summary>
    public Dictionary<string, Dictionary<string, string>> ReadPresets()
    {
        var result = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);
        string? preset = null;

        foreach (var line in _lines)
        {
            if (TryReadHeader(line, out var name))
            {
                preset = name;
                if (preset is not null && !result.ContainsKey(preset))
                {
                    result[preset] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                }

                continue;
            }

            if (preset is not null && TryReadPair(line, out var key, out var value))
            {
                result[preset][key] = value;
            }
        }

        return result;
    }

    /// <summary>
    /// Sets a key, either global (<paramref name="preset"/> null) or inside one preset.
    ///
    /// An existing key keeps its position and any trailing comment. A new key is appended to
    /// the end of its section rather than the end of the file, so it lands under the heading a
    /// reader would expect.
    /// </summary>
    public void Set(string key, string value, string? preset = null)
    {
        var inTargetSection = preset is null;
        var lastLineOfSection = -1;
        string? current = null;

        for (var i = 0; i < _lines.Count; i++)
        {
            if (TryReadHeader(_lines[i], out var name))
            {
                current = name;
                inTargetSection = preset is null
                    ? name is null
                    : string.Equals(name, preset, StringComparison.OrdinalIgnoreCase);

                if (inTargetSection)
                {
                    lastLineOfSection = i;
                }

                continue;
            }

            if (!inTargetSection)
            {
                continue;
            }

            if (_lines[i].Trim().Length > 0)
            {
                lastLineOfSection = i;
            }

            if (TryReadPair(_lines[i], out var existingKey, out _) &&
                existingKey.Equals(key, StringComparison.OrdinalIgnoreCase))
            {
                _lines[i] = RewriteValue(_lines[i], value);
                return;
            }
        }

        _ = current;

        var newLine = $"{key} = {value}";

        if (lastLineOfSection >= 0)
        {
            _lines.Insert(lastLineOfSection + 1, newLine);
        }
        else
        {
            if (preset is not null)
            {
                if (_lines.Count > 0 && _lines[^1].Trim().Length > 0)
                {
                    _lines.Add("");
                }

                _lines.Add($"[preset {preset}]");
            }

            _lines.Add(newLine);
        }
    }

    /// <summary>Replaces the value on an existing line, keeping indentation and any comment.</summary>
    private static string RewriteValue(string line, string value)
    {
        var equals = line.IndexOf('=');
        var prefix = line[..(equals + 1)];
        var rest = line[(equals + 1)..];

        // Preserve a trailing comment so the documentation survives an edit.
        var comment = "";
        for (var i = 1; i < rest.Length; i++)
        {
            if (char.IsWhiteSpace(rest[i - 1]) && (rest[i] == ';' || rest[i] == '#'))
            {
                comment = rest[(i - 1)..];
                break;
            }
        }

        return $"{prefix} {value}{comment}";
    }

    /// <summary>
    /// Writes to a temporary file and moves it into place, so a reader never sees a partial
    /// config, and a failure part-way through leaves the original intact.
    /// </summary>
    public void Save(string path)
    {
        var temp = path + ".tmp";
        File.WriteAllLines(temp, _lines, new UTF8Encoding(false));
        File.Move(temp, path, overwrite: true);
    }
}
