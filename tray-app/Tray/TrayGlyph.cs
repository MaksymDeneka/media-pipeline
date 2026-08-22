using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media.Imaging;
using MediaPipelineTray.Models;

namespace MediaPipelineTray.Tray;

/// <summary>
/// The application glyph, drawn rather than shipped as a resource so it can be tinted per
/// state without carrying an .ico per state.
///
/// One definition serves both the notification area and the window, which is what makes the
/// taskbar button match the tray icon.
///
/// Neutral states follow the system theme, because the same white that reads well on a dark
/// taskbar disappears on a light one. Only failure and success are fixed colours, and those
/// two carry enough contrast either way.
/// </summary>
public static class TrayGlyph
{
    private const int Size = 32;

    public static Icon Render(ActivityState state, bool isDark)
    {
        var (stroke, fill) = Palette(state, isDark);

        using var bitmap = new Bitmap(Size, Size);

        using (var graphics = Graphics.FromImage(bitmap))
        {
            graphics.SmoothingMode = SmoothingMode.AntiAlias;
            graphics.Clear(Color.Transparent);

            var bounds = new Rectangle(5, 5, Size - 11, Size - 11);
            using var path = RoundedRect(bounds, 6);

            if (fill != Color.Transparent)
            {
                using var brush = new SolidBrush(fill);
                graphics.FillPath(brush, path);
            }

            using var pen = new Pen(stroke, 3f);
            graphics.DrawPath(pen, path);
        }

        return Icon.FromHandle(bitmap.GetHicon());
    }

    /// <summary>The same glyph as a WPF image source, for Window.Icon and the taskbar.</summary>
    public static BitmapSource RenderForWindow(ActivityState state, bool isDark)
    {
        using var icon = Render(state, isDark);

        var source = Imaging.CreateBitmapSourceFromHIcon(
            icon.Handle, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());

        source.Freeze();
        return source;
    }

    private static (Color Stroke, Color Fill) Palette(ActivityState state, bool isDark)
    {
        // Failure and success keep their colour in either theme: they are the only two states
        // the whole application spends colour on.
        if (state == ActivityState.Failed)
        {
            var red = isDark ? Color.FromArgb(0xE8, 0x69, 0x5C) : Color.FromArgb(0xB6, 0x24, 0x19);
            return (red, red);
        }

        if (state == ActivityState.Finished)
        {
            var green = isDark ? Color.FromArgb(0x4F, 0xB8, 0x77) : Color.FromArgb(0x1A, 0x7F, 0x42);
            return (green, green);
        }

        var ink = isDark ? Color.FromArgb(0xF1, 0xF1, 0xF3) : Color.FromArgb(0x10, 0x10, 0x12);
        var muted = isDark ? Color.FromArgb(0x9B, 0x9B, 0xA2) : Color.FromArgb(0x66, 0x66, 0x6D);

        return state switch
        {
            // Running fills solid, which is the loudest this palette gets without colour.
            ActivityState.Running => (ink, ink),
            ActivityState.Paused => (muted, Color.Transparent),
            _ => (muted, Color.Transparent),
        };
    }

    private static GraphicsPath RoundedRect(Rectangle bounds, int radius)
    {
        var diameter = radius * 2;
        var path = new GraphicsPath();

        path.AddArc(bounds.X, bounds.Y, diameter, diameter, 180, 90);
        path.AddArc(bounds.Right - diameter, bounds.Y, diameter, diameter, 270, 90);
        path.AddArc(bounds.Right - diameter, bounds.Bottom - diameter, diameter, diameter, 0, 90);
        path.AddArc(bounds.X, bounds.Bottom - diameter, diameter, diameter, 90, 90);
        path.CloseFigure();

        return path;
    }
}
