using System.Windows;
using System.Windows.Controls;
using MediaPipelineTray.ViewModels;

namespace MediaPipelineTray.Views;

public partial class UploadsView : UserControl
{
    public UploadsView() => InitializeComponent();

    private UploadsViewModel? Model => DataContext as UploadsViewModel;

    private void OnRefresh(object sender, RoutedEventArgs e) => Model?.Refresh();

    private void OnUploadFile(object sender, RoutedEventArgs e)
    {
        if (sender is FrameworkElement { DataContext: SyncFileRow file })
        {
            Model?.QueueFile(file);
        }
    }

    private void OnUploadWorkspace(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement { DataContext: SyncWorkspaceRow workspace } || Model is null)
        {
            return;
        }

        var pending = workspace.Files.Count(f => !f.IsQueued);

        if (pending == 0)
        {
            return;
        }

        var total = UploadRow.Format(workspace.Files.Where(f => !f.IsQueued).Sum(f => f.Length));

        var answer = MessageBox.Show(
            Window.GetWindow(this),
            $"Upload {pending} file(s) from {workspace.Name}?\n\n{total} in total.\n\n"
            + (Model.DeleteAfterUpload
                ? "Each local file is deleted once its remote copy is verified."
                : "Local files are kept."),
            "Upload workspace",
            MessageBoxButton.OKCancel,
            MessageBoxImage.Question);

        if (answer == MessageBoxResult.OK)
        {
            Model.QueueWorkspace(workspace);
        }
    }

    private void OnCancel(object sender, RoutedEventArgs e) => Model?.Cancel();
}
