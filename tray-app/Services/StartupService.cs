using System.Diagnostics;
using System.IO;
using Microsoft.Win32;

namespace MediaPipelineTray.Services;

/// <summary>
/// Two pieces of Windows integration the app needs to behave like a tray app people can find:
/// starting with Windows, and not being buried in the notification-area overflow.
/// </summary>
public static class StartupService
{
    private const string RunKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";
    private const string ValueName = "MediaPipelineTray";

    /// <summary>The executable, resolved so a shortcut or relative launch still records a real path.</summary>
    public static string ExecutablePath =>
        Environment.ProcessPath ?? Process.GetCurrentProcess().MainModule?.FileName ?? "";

    public static bool IsEnabled
    {
        get
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(RunKey);
                return key?.GetValue(ValueName) is string value && value.Length > 0;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                return false;
            }
        }
    }

    /// <summary>
    /// Adds or removes the per-user startup entry. Per-user rather than machine-wide, so it
    /// needs no elevation and follows the account that actually uses the pipeline.
    /// </summary>
    public static bool SetEnabled(bool enabled)
    {
        try
        {
            using var key = Registry.CurrentUser.CreateSubKey(RunKey, writable: true);
            if (key is null)
            {
                return false;
            }

            if (enabled)
            {
                var path = ExecutablePath;
                if (path.Length == 0)
                {
                    return false;
                }

                key.SetValue(ValueName, $"\"{path}\"");
            }
            else
            {
                key.DeleteValue(ValueName, throwOnMissingValue: false);
            }

            return true;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return false;
        }
    }

    /// <summary>
    /// Turns startup on the first time the app runs, then never touches it again.
    ///
    /// A tray app that monitors a background service is expected to come back after a reboot,
    /// so on is the useful default. The marker is what makes it a default rather than a policy:
    /// once it exists, switching startup off in Settings stays off.
    /// </summary>
    public static void ApplyFirstRunDefault()
    {
        const string appKey = @"SOFTWARE\MediaPipelineTray";
        const string marker = "StartupInitialised";

        try
        {
            using var key = Registry.CurrentUser.CreateSubKey(appKey, writable: true);
            if (key is null || key.GetValue(marker) is not null)
            {
                return;
            }

            SetEnabled(true);
            key.SetValue(marker, 1, RegistryValueKind.DWord);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            // Startup is a convenience; failing to set it must not stop the app.
        }
    }

    /// <summary>
    /// Asks Windows 11 to show this app's notification icon on the taskbar rather than inside
    /// the hidden overflow.
    ///
    /// New tray icons are hidden by default, which makes a tray-first app look like it failed
    /// to start: the window closes to a tray icon nobody can see. Windows records a per-icon
    /// preference under Control Panel\NotifyIconSettings, keyed by the executable, and honours
    /// IsPromoted there.
    ///
    /// The entry only appears after the icon has been shown at least once, so a first run
    /// promotes itself for the next one. Returns whether anything was changed.
    /// </summary>
    public static bool PromoteTrayIcon()
    {
        var executable = ExecutablePath;
        if (executable.Length == 0)
        {
            return false;
        }

        try
        {
            using var settings = Registry.CurrentUser.OpenSubKey(
                @"Control Panel\NotifyIconSettings", writable: true);

            if (settings is null)
            {
                return false;
            }

            var changed = false;

            foreach (var name in settings.GetSubKeyNames())
            {
                using var entry = settings.OpenSubKey(name, writable: true);

                if (entry?.GetValue("ExecutablePath") is not string path)
                {
                    continue;
                }

                if (!string.Equals(path, executable, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                // Leave an explicit choice alone: if it is already promoted, the user may have
                // dragged it there themselves, and rewriting it would be pointless churn.
                if (entry.GetValue("IsPromoted") is int promoted && promoted == 1)
                {
                    continue;
                }

                entry.SetValue("IsPromoted", 1, RegistryValueKind.DWord);
                changed = true;
            }

            return changed;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Security.SecurityException)
        {
            // Not being able to promote is cosmetic; the icon still works in the overflow.
            return false;
        }
    }
}
