using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using MediaPipelineTray.Models;

namespace MediaPipelineTray.Tray;

/// <summary>
/// The notification-area icon.
///
/// The glyph is drawn rather than shipped as a resource, so it can be tinted per state without
/// carrying four .ico files. It stays monochrome while everything is normal and only takes on
/// colour when a job finishes or fails, which is the whole point: a glance at the tray should
/// tell you whether anything needs you.
/// </summary>
public sealed class TrayIcon : IDisposable
{
    private readonly NotifyIcon _icon;
    private readonly Dictionary<ActivityState, Icon> _icons = [];

    public TrayIcon()
    {
        _icon = new NotifyIcon
        {
            Visible = true,
            Text = "Media Pipeline",
            ContextMenuStrip = new ContextMenuStrip(),
        };

        _icon.DoubleClick += (_, _) => OpenRequested?.Invoke(this, EventArgs.Empty);
        SetState(ActivityState.Idle, "Media Pipeline");
    }

    public event EventHandler? OpenRequested;
    public event EventHandler? PauseRequested;
    public event EventHandler? QuitRequested;

    /// <summary>Rebuilds the menu. Called on each state change so labels stay truthful.</summary>
    public void BuildMenu(bool paused, IReadOnlyList<string> lines)
    {
        var menu = _icon.ContextMenuStrip!;
        menu.Items.Clear();

        foreach (var line in lines)
        {
            menu.Items.Add(new ToolStripMenuItem(line) { Enabled = false });
        }

        if (lines.Count > 0)
        {
            menu.Items.Add(new ToolStripSeparator());
        }

        var open = new ToolStripMenuItem("Open window");
        open.Click += (_, _) => OpenRequested?.Invoke(this, EventArgs.Empty);
        menu.Items.Add(open);

        var pause = new ToolStripMenuItem(paused ? "Resume all" : "Pause all");
        pause.Click += (_, _) => PauseRequested?.Invoke(this, EventArgs.Empty);
        menu.Items.Add(pause);

        menu.Items.Add(new ToolStripSeparator());

        var quit = new ToolStripMenuItem("Quit");
        quit.Click += (_, _) => QuitRequested?.Invoke(this, EventArgs.Empty);
        menu.Items.Add(quit);
    }

    public void SetState(ActivityState state, string tooltip)
    {
        _icon.Icon = GetIcon(state);

        // The tray truncates past 63 characters and throws in some Windows versions if longer.
        _icon.Text = tooltip.Length > 62 ? tooltip[..62] : tooltip;
    }

    public void Notify(string title, string message, bool isFailure)
    {
        _icon.BalloonTipTitle = title;
        _icon.BalloonTipText = message;
        _icon.BalloonTipIcon = isFailure ? ToolTipIcon.Error : ToolTipIcon.Info;
        _icon.ShowBalloonTip(5000);
    }

    private Icon GetIcon(ActivityState state)
    {
        if (_icons.TryGetValue(state, out var cached))
        {
            return cached;
        }

        var icon = Render(state);
        _icons[state] = icon;
        return icon;
    }

    /// <summary>
    /// Draws a rounded square outline. Idle and paused are a hollow outline, running fills it,
    /// and only finished and failed are tinted.
    /// </summary>
    private static Icon Render(ActivityState state)
    {
        const int size = 32;

        var (stroke, fill) = state switch
        {
            ActivityState.Failed => (Color.FromArgb(0xE8, 0x69, 0x5C), Color.FromArgb(0xE8, 0x69, 0x5C)),
            ActivityState.Finished => (Color.FromArgb(0x4F, 0xB8, 0x77), Color.FromArgb(0x4F, 0xB8, 0x77)),
            ActivityState.Running => (Color.White, Color.White),
            ActivityState.Paused => (Color.FromArgb(0x9B, 0x9B, 0xA2), Color.Transparent),
            _ => (Color.FromArgb(0xC8, 0xC8, 0xCC), Color.Transparent),
        };

        using var bitmap = new Bitmap(size, size);
        using (var graphics = Graphics.FromImage(bitmap))
        {
            graphics.SmoothingMode = SmoothingMode.AntiAlias;
            graphics.Clear(Color.Transparent);

            var rect = new Rectangle(5, 5, size - 11, size - 11);

            using var path = RoundedRect(rect, 6);

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

    public void Dispose()
    {
        _icon.Visible = false;
        _icon.Dispose();

        foreach (var icon in _icons.Values)
        {
            icon.Dispose();
        }
    }
}
