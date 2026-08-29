using System.Diagnostics;
using System.IO;
using Microsoft.Win32;

namespace MediaPipelineTray.Services;

/// <summary>
/// Registers the tray app to start for the current user when they sign in.
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

                key.SetValue(ValueName, $"\"{path}\" --startup");
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
            if (key is null)
            {
                return;
            }

            if (key.GetValue(marker) is null)
            {
                if (SetEnabled(true))
                {
                    key.SetValue(marker, 1, RegistryValueKind.DWord);
                }

                return;
            }

            // Rewrite older entries once so Windows passes the tray-only startup flag.
            if (IsEnabled)
            {
                SetEnabled(true);
            }
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            // Startup is a convenience; failing to set it must not stop the app.
        }
    }
}
