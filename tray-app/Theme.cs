using System.IO;
using System.Windows;
using System.Windows.Media;
using Microsoft.Win32;

namespace MediaPipelineTray;

/// <summary>
/// The application palette.
///
/// Deliberately monochrome: every ordinary state is carried by contrast and shape, and the two
/// colours are spent only on the two things worth interrupting for. Green means a job finished,
/// red means one failed. Nothing else is ever coloured, including work in progress, which is
/// drawn at full ink contrast instead.
/// </summary>
public static class Theme
{
    public static bool IsDark { get; private set; }

    private static readonly Dictionary<string, string> Light = new()
    {
        ["Bg"] = "#F4F4F5",
        ["Surface"] = "#FFFFFF",
        ["Surface2"] = "#FAFAFB",
        ["Sunk"] = "#ECECEE",
        ["Track"] = "#E4E4E7",
        ["Ink"] = "#101012",
        ["Muted"] = "#66666D",
        ["Faint"] = "#96969D",
        ["Line"] = "#E2E2E6",
        ["LineSoft"] = "#EDEDF0",
        ["Titlebar"] = "#EDEDEF",
        ["Ok"] = "#1A7F42",
        ["OkBg"] = "#E4F1E8",
        ["Fail"] = "#B62419",
        ["FailBg"] = "#FBE6E4",
    };

    private static readonly Dictionary<string, string> Dark = new()
    {
        ["Bg"] = "#0C0C0E",
        ["Surface"] = "#161619",
        ["Surface2"] = "#1C1C20",
        ["Sunk"] = "#121215",
        ["Track"] = "#2A2A30",
        ["Ink"] = "#F1F1F3",
        ["Muted"] = "#9B9BA2",
        ["Faint"] = "#6C6C74",
        ["Line"] = "#26262B",
        ["LineSoft"] = "#1F1F24",
        ["Titlebar"] = "#1A1A1E",
        ["Ok"] = "#4FB877",
        ["OkBg"] = "#14251A",
        ["Fail"] = "#E8695C",
        ["FailBg"] = "#2A1614",
    };

    /// <summary>
    /// Applies the palette matching the user's Windows setting. Brushes are replaced in place
    /// on the existing resource keys, so a theme switch repaints without rebuilding the UI.
    /// </summary>
    public static void Apply(ResourceDictionary resources, bool? forceDark = null)
    {
        IsDark = forceDark ?? SystemPrefersDark();
        var palette = IsDark ? Dark : Light;

        foreach (var (key, hex) in palette)
        {
            var brush = new SolidColorBrush((Color)ColorConverter.ConvertFromString(hex));
            brush.Freeze();
            resources[key + "Brush"] = brush;
        }
    }

    private static bool SystemPrefersDark()
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize");

            // The value is "apps use light theme", so zero means dark.
            return key?.GetValue("AppsUseLightTheme") is int value && value == 0;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return false;
        }
    }
}
