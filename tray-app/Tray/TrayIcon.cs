using System.Drawing;
using System.Windows.Forms;
using MediaPipelineTray.Models;

namespace MediaPipelineTray.Tray;

/// <summary>
/// The notification-area icon.
///
/// The glyph comes from <see cref="TrayGlyph"/>, which the window also uses, so the taskbar
/// button and the tray icon are always the same picture in the same state.
/// </summary>
public sealed class TrayIcon : IDisposable
{
    private readonly NotifyIcon _icon;
    private readonly Dictionary<(ActivityState State, bool IsDark), Icon> _icons = [];

    public TrayIcon()
    {
        _icon = new NotifyIcon
        {
            Visible = true,
            Text = "Media Pipeline",
            ContextMenuStrip = new ContextMenuStrip(),
        };

        // A single left click opens the window. Right click is left to the context menu, which
        // NotifyIcon shows on its own.
        _icon.MouseClick += (_, e) =>
        {
            if (e.Button == MouseButtons.Left)
            {
                OpenRequested?.Invoke(this, EventArgs.Empty);
            }
        };

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
        open.Font = new System.Drawing.Font(open.Font, System.Drawing.FontStyle.Bold);
        open.Click += (_, _) => OpenRequested?.Invoke(this, EventArgs.Empty);
        menu.Items.Add(open);

        var pause = new ToolStripMenuItem(paused ? "Resume all" : "Pause all");
        pause.Click += (_, _) => PauseRequested?.Invoke(this, EventArgs.Empty);
        menu.Items.Add(pause);

        menu.Items.Add(new ToolStripSeparator());

        // Named so it is clear this is the only thing that actually ends the app: closing the
        // window only hides it.
        var quit = new ToolStripMenuItem("Quit Media Pipeline");
        quit.Click += (_, _) => QuitRequested?.Invoke(this, EventArgs.Empty);
        menu.Items.Add(quit);
    }

    public void SetState(ActivityState state, string tooltip)
    {
        _icon.Icon = GetIcon(state);

        // The tray truncates past 63 characters and throws on some Windows versions if longer.
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
        // The tray sits on the taskbar, so it follows the taskbar's theme rather than the
        // app's window theme. They are the same setting in practice.
        var key = (state, Theme.IsDark);

        if (_icons.TryGetValue(key, out var cached))
        {
            return cached;
        }

        var icon = TrayGlyph.Render(state, Theme.IsDark);
        _icons[key] = icon;
        return icon;
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
