using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using MediaPipelineTray.ViewModels;

namespace MediaPipelineTray.Views;

public partial class ActivityView : UserControl
{
    public ActivityView() => InitializeComponent();

    private MainViewModel? Model => DataContext as MainViewModel;

    private void OnPauseLane(object sender, RoutedEventArgs e)
    {
        if (sender is FrameworkElement { DataContext: LaneRow lane })
        {
            Model?.PauseLane(lane);
        }
    }

    private void OnOpenFailedFolder(object sender, RoutedEventArgs e)
    {
        if (sender is FrameworkElement { DataContext: FailureRow failure })
        {
            Model?.OpenLaneFolder(failure.Preset, failure.Workspace);
        }
    }

    private void OnRequeue(object sender, RoutedEventArgs e)
    {
        if (Model is null)
        {
            return;
        }

        var moved = Model.RequeueFailures();

        MessageBox.Show(
            Window.GetWindow(this),
            moved == 0
                ? "Nothing to requeue. The failed folder is empty."
                : $"Moved {moved} file(s) back to the input folder.",
            "Requeue",
            MessageBoxButton.OK,
            MessageBoxImage.Information);
    }

    private void OnDismissFailures(object sender, RoutedEventArgs e) => Model?.DismissFailures();

    private void OnZip(object sender, RoutedEventArgs e) => Archive(sender, thenUpload: false);

    private void OnZipAndUpload(object sender, RoutedEventArgs e) => Archive(sender, thenUpload: true);

    private void Archive(object sender, bool thenUpload)
    {
        if (sender is not FrameworkElement { DataContext: FinishedRow row } || Model is null)
        {
            return;
        }

        try
        {
            var result = Model.ArchiveFinished(row, thenUpload);

            var message = new StringBuilder();
            message.AppendLine($"{result.FileCount} file(s) collected into:");
            message.AppendLine(result.Path);
            message.AppendLine();
            message.AppendLine(UploadRow.Format(result.Bytes));

            if (thenUpload)
            {
                message.AppendLine();
                message.AppendLine("Queued for upload.");
            }

            // Say what was skipped rather than quietly producing a smaller archive.
            if (result.Missing.Count > 0)
            {
                message.AppendLine();
                message.AppendLine(
                    $"{result.Missing.Count} file(s) were no longer on disk and were skipped.");
            }

            MessageBox.Show(
                Window.GetWindow(this),
                message.ToString(),
                thenUpload ? "Zipped and queued" : "Zipped",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
        catch (Exception ex) when (ex is IOException or InvalidOperationException or UnauthorizedAccessException)
        {
            MessageBox.Show(
                Window.GetWindow(this),
                ex.Message,
                "Could not zip that output",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }
    }
}
