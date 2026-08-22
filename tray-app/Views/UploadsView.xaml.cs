using System.IO;
using System.Windows;
using System.Windows.Controls;
using MediaPipelineTray.ViewModels;
using Microsoft.Win32;

namespace MediaPipelineTray.Views;

public partial class UploadsView : UserControl
{
    public UploadsView() => InitializeComponent();

    private UploadsViewModel? Model => DataContext as UploadsViewModel;

    private async void OnChooseFile(object sender, RoutedEventArgs e)
    {
        if (Model is null)
        {
            return;
        }

        var dialog = new OpenFileDialog
        {
            Title = "Choose a file to upload",
            InitialDirectory = Directory.Exists(Model.SyncFolder) ? Model.SyncFolder : null,
            CheckFileExists = true,
        };

        if (dialog.ShowDialog(Window.GetWindow(this)) == true)
        {
            await Model.StartAsync(dialog.FileName);
        }
    }

    /// <summary>
    /// The common case: the sync folder holds one big archive waiting to go. Picking the
    /// largest matches what the old script did when run with no arguments.
    /// </summary>
    private async void OnUploadLargest(object sender, RoutedEventArgs e)
    {
        if (Model is null)
        {
            return;
        }

        var candidates = Model.FindCandidates();

        if (candidates.Count == 0)
        {
            MessageBox.Show(
                Window.GetWindow(this),
                $"Nothing to upload. Put a file in {Model.SyncFolder} first.",
                "Uploads",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        var largest = candidates[0];

        var answer = MessageBox.Show(
            Window.GetWindow(this),
            $"Upload {largest.Name}?\n\n{UploadRow.Format(largest.Length)}",
            "Uploads",
            MessageBoxButton.OKCancel,
            MessageBoxImage.Question);

        if (answer == MessageBoxResult.OK)
        {
            await Model.StartAsync(largest.FullName);
        }
    }

    private void OnCancel(object sender, RoutedEventArgs e) => Model?.Cancel();
}
