<#
.SYNOPSIS
    End-to-end check of the tray app's lifecycle.

.DESCRIPTION
    Covers the tray app lifecycle: Windows startup opens only the tray icon, closing the window
    hides it rather than exiting, a single click brings the window back, and only the tray menu
    quits.

    Drives the real UI. The context menu is operated by keyboard, because a WinForms menu
    hosted in a WPF app is not introspectable through UI Automation.

    This starts and stops the app, so do not run it while relying on the tray icon.
#>

# End-to-end check of the tray app's lifecycle:
#   closing the window hides it, the tray icon brings it back, and only the tray menu quits.

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class W {
    [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] public static extern void mouse_event(uint f, uint dx, uint dy, uint d, IntPtr e);
    [DllImport("user32.dll")] public static extern void keybd_event(byte vk, byte scan, uint flags, IntPtr extra);
    public const uint WM_CLOSE = 0x0010;
    public const byte VK_UP = 0x26, VK_RETURN = 0x0D;
    public const uint KEYUP = 0x0002;
    public const uint RIGHTDOWN = 0x08, RIGHTUP = 0x10;
}
'@

$failures = 0
function Check($label, $ok, $detail) {
    if ($ok) { Write-Host "   PASS  $label" -ForegroundColor Green }
    else { Write-Host "   FAIL  $label :: $detail" -ForegroundColor Red; $script:failures++ }
}

$exe = 'D:\Projects\media-pipeline\tray-app\bin\Release\net8.0-windows\MediaPipelineTray.exe'

Write-Host "`n== Tray lifecycle"

Get-Process MediaPipelineTray -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 2

Start-Process $exe -ArgumentList '--startup'
Start-Sleep -Seconds 6

$app = Get-Process MediaPipelineTray -ErrorAction SilentlyContinue
Check "the app starts" ($null -ne $app) "not running"
if (-not $app) { exit 1 }
Check "startup does not open the window" `
    ($app.MainWindowHandle -eq [IntPtr]::Zero) "a window is visible"

# --- autostart ---
$run = Get-ItemProperty 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run' -ErrorAction SilentlyContinue
Check "it registered itself to start with Windows" `
    ($null -ne $run.MediaPipelineTray) "no Run entry"

if ($run.MediaPipelineTray) {
    Check "the startup entry points at this executable" `
        ($run.MediaPipelineTray -like "*MediaPipelineTray.exe*") $run.MediaPipelineTray
    Check "the startup entry requests tray-only mode" `
        ($run.MediaPipelineTray -like "*--startup*") $run.MediaPipelineTray
}

# --- find the tray icon ---
$root = [System.Windows.Automation.AutomationElement]::RootElement

$chevrons = $root.FindAll(
    [System.Windows.Automation.TreeScope]::Descendants,
    (New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::NameProperty, 'Show Hidden Icons')))

if ($chevrons.Count -gt 0) {
    try {
        $chevrons.Item(0).GetCurrentPattern(
            [System.Windows.Automation.InvokePattern]::Pattern).Invoke()
        Start-Sleep -Milliseconds 900
    } catch { }
}

$icon = $null
$buttons = $root.FindAll(
    [System.Windows.Automation.TreeScope]::Descendants,
    (New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
        [System.Windows.Automation.ControlType]::Button)))

foreach ($b in $buttons) {
    if ($b.Current.Name -like '*Media Pipeline*') { $icon = $b; break }
}

Check "the tray icon exists" ($null -ne $icon) "not found in the automation tree"
if (-not $icon) { exit 1 }

# --- single click reopens ---
$icon.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()
Start-Sleep -Seconds 2
$app.Refresh()

Check "a single click reopens the window" ([W]::IsWindowVisible($app.MainWindowHandle)) "still hidden"

# --- close hides ---
[void][W]::SendMessage($app.MainWindowHandle, [W]::WM_CLOSE, [IntPtr]::Zero, [IntPtr]::Zero)
Start-Sleep -Seconds 2
$app.Refresh()

Check "closing the window does not exit" (-not $app.HasExited) "process exited"
Check "the window is hidden" (-not [W]::IsWindowVisible($app.MainWindowHandle)) "still visible"

# --- right click offers quit ---
# Re-find the icon: clicking it closed the overflow flyout, so the earlier element is stale
# and reports a NaN rectangle.
function Find-TrayIcon {
    $found = $null
    $all = $root.FindAll(
        [System.Windows.Automation.TreeScope]::Descendants,
        (New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
            [System.Windows.Automation.ControlType]::Button)))

    foreach ($b in $all) {
        if ($b.Current.Name -like '*Media Pipeline*') {
            $r = $b.Current.BoundingRectangle
            if (-not [double]::IsNaN($r.X)) { return $b }
            $found = $b
        }
    }

    return $found
}

$icon = Find-TrayIcon
$rect = $icon.Current.BoundingRectangle

if ([double]::IsNaN($rect.X)) {
    # Still off-screen, so reveal the overflow and look again.
    $chevrons = $root.FindAll(
        [System.Windows.Automation.TreeScope]::Descendants,
        (New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::NameProperty, 'Show Hidden Icons')))

    if ($chevrons.Count -gt 0) {
        try {
            $chevrons.Item(0).GetCurrentPattern(
                [System.Windows.Automation.InvokePattern]::Pattern).Invoke()
            Start-Sleep -Milliseconds 900
        } catch { }
    }

    $icon = Find-TrayIcon
    $rect = $icon.Current.BoundingRectangle
}

Check "the icon has a screen position" (-not [double]::IsNaN($rect.X)) "rectangle is NaN"
if ([double]::IsNaN($rect.X)) { Write-Host ""; exit 1 }
$x = [int]($rect.X + $rect.Width / 2)
$y = [int]($rect.Y + $rect.Height / 2)

[void][W]::SetCursorPos($x, $y)
Start-Sleep -Milliseconds 250
[W]::mouse_event([W]::RIGHTDOWN, 0, 0, 0, [IntPtr]::Zero)
Start-Sleep -Milliseconds 80
[W]::mouse_event([W]::RIGHTUP, 0, 0, 0, [IntPtr]::Zero)
Start-Sleep -Milliseconds 1200

# A WinForms context menu hosted in a WPF app is not introspectable through UI Automation,
# so drive it the way a keyboard user would: Up selects the last item, which is Quit.
Start-Sleep -Milliseconds 400

[W]::keybd_event([W]::VK_UP, 0, 0, [IntPtr]::Zero)
[W]::keybd_event([W]::VK_UP, 0, [W]::KEYUP, [IntPtr]::Zero)
Start-Sleep -Milliseconds 300
[W]::keybd_event([W]::VK_RETURN, 0, 0, [IntPtr]::Zero)
[W]::keybd_event([W]::VK_RETURN, 0, [W]::KEYUP, [IntPtr]::Zero)

Start-Sleep -Seconds 3

$gone = $null -eq (Get-Process MediaPipelineTray -ErrorAction SilentlyContinue)
Check "quitting from the tray menu exits the app" $gone "still running"

Write-Host ""
if ($failures -eq 0) { Write-Host "TRAY LIFECYCLE OK" -ForegroundColor Green; exit 0 }
Write-Host "TRAY LIFECYCLE FAILED: $failures" -ForegroundColor Red
exit 1
